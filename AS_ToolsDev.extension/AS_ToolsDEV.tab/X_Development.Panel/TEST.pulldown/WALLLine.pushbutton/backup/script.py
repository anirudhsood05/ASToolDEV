# -*- coding: utf-8 -*-
"""
Script:   Wall Face Model Lines
Desc:     Creates model lines on both exterior and interior faces of selected
          walls in the active plan view. Works inside group edit mode.
Author:   AUK BIM Team
Usage:    Select one or more walls (or enter group edit mode first), then click.
Result:   Model lines created on both sides of each wall, matching wall length.
"""
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms

import clr
clr.AddReference("RevitAPI")

logger = script.get_logger()
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

# ── Constants ────────────────────────────────────────────────────────────────
TOOL_NAME = "Wall Face Model Lines"
LINE_STYLE_NAME = "Wide Lines"  # fallback line style; change as needed


# ── Helpers ──────────────────────────────────────────────────────────────────
def eid_int(eid):
    """Cross-version ElementId to int helper (2023-2026)."""
    try:
        return eid.Value          # Revit 2026+
    except AttributeError:
        return eid.IntegerValue   # Revit 2023-2025


def get_line_style_id(style_name):
    """Return GraphicsStyle Id for the named line style, or None."""
    try:
        styles = DB.FilteredElementCollector(doc) \
            .OfClass(DB.GraphicsStyle) \
            .ToElements()
        for gs in styles:
            try:
                cat = gs.GraphicsStyleCategory
                if cat and cat.Parent and cat.Parent.Name == "Lines":
                    if DB.Element.Name.GetValue(gs) == style_name:
                        return gs.Id
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_wall_face_curves(wall):
    """
    Return two lists of DB.Line segments representing the exterior and
    interior faces of a wall projected onto the XY plane (plan view).

    Uses wall location curve + compound structure offsets for robustness
    inside group edit mode where HostObjectUtils may not be available.
    """
    loc = wall.Location
    if not loc or not isinstance(loc, DB.LocationCurve):
        return None, None

    wall_curve = loc.Curve

    # Get wall direction and perpendicular normal in XY plane
    p0 = wall_curve.GetEndPoint(0)
    p1 = wall_curve.GetEndPoint(1)
    direction = (p1 - p0).Normalize()
    normal = DB.XYZ(-direction.Y, direction.X, 0).Normalize()

    # Determine half-thicknesses from compound structure offsets
    wall_type = doc.GetElement(wall.GetTypeId())
    if wall_type is None:
        return None, None

    cs = wall_type.GetCompoundStructure()
    if cs is None:
        # Basic wall — use Width / 2
        half_ext = wall.Width / 2.0
        half_int = wall.Width / 2.0
    else:
        # Use the location-line offset to compute correct face positions
        # GetOffsetForLocationLine gives offset from centre to each edge
        total_width = cs.GetWidth()
        half_ext = total_width / 2.0
        half_int = total_width / 2.0

    # Account for wall location line offset (wall may not be centred)
    loc_line = wall.get_Parameter(DB.BuiltInParameter.WALL_KEY_REF_PARAM)
    if loc_line and loc_line.HasValue:
        loc_line_val = loc_line.AsInteger()
        # 0=WallCenterline, 1=CoreCenterline, 2=FinishFaceExt,
        # 3=FinishFaceInt, 4=CoreFaceExt, 5=CoreFaceInt
        if cs:
            try:
                offset_ext = cs.GetOffsetForLocationLine(
                    DB.WallLocationLine.FinishFaceExterior)
                offset_int = cs.GetOffsetForLocationLine(
                    DB.WallLocationLine.FinishFaceInterior)
            except Exception:
                offset_ext = half_ext
                offset_int = -half_int
        else:
            offset_ext = half_ext
            offset_int = -half_int
    else:
        offset_ext = half_ext
        offset_int = -half_int

    # Build exterior and interior curves
    # Flatten to plan view Z
    z = p0.Z

    def offset_line(crv, offset_dist):
        """Offset a line by distance along the wall normal, keep at Z."""
        s = crv.GetEndPoint(0)
        e = crv.GetEndPoint(1)
        off = normal * offset_dist
        ns = DB.XYZ(s.X + off.X, s.Y + off.Y, z)
        ne = DB.XYZ(e.X + off.X, e.Y + off.Y, z)
        if ns.DistanceTo(ne) < 0.001:
            return None
        return DB.Line.CreateBound(ns, ne)

    ext_line = offset_line(wall_curve, offset_ext)
    int_line = offset_line(wall_curve, offset_int)

    return ext_line, int_line


def is_group_edit_mode():
    """Check if the document is currently in group edit mode."""
    try:
        # Revit 2025.2+ / 2026
        if hasattr(doc, "IsInEditMode"):
            return doc.IsInEditMode()
    except Exception:
        pass
    # Fallback: check if the active view is inside a group context
    # by seeing if we can find an open edit group
    try:
        sel = uidoc.Selection.GetElementIds()
        for eid in sel:
            el = doc.GetElement(eid)
            if el and isinstance(el, DB.Group):
                return True
    except Exception:
        pass
    return False


# ── Validation ───────────────────────────────────────────────────────────────
def validate():
    """Pre-condition checks."""
    view = doc.ActiveView
    if view is None:
        forms.alert("No active view found.", title=TOOL_NAME, exitscript=True)

    # Must be a plan view for model lines in plan
    allowed = [
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.EngineeringPlan,
        DB.ViewType.AreaPlan,
    ]
    if view.ViewType not in allowed:
        forms.alert(
            "This tool works in plan views only.\n\n"
            "Current view type: {}".format(view.ViewType),
            title=TOOL_NAME,
            exitscript=True
        )
    return view


def get_walls():
    """Retrieve walls from current selection."""
    selection = revit.get_selection()
    elements = selection.elements

    if not elements:
        forms.alert(
            "Please select one or more walls first.",
            title=TOOL_NAME,
            exitscript=True
        )

    walls = [e for e in elements if isinstance(e, DB.Wall)]
    if not walls:
        forms.alert(
            "No walls found in selection.\n"
            "Please select walls and try again.",
            title=TOOL_NAME,
            exitscript=True
        )
    return walls


# ── Main Logic ───────────────────────────────────────────────────────────────
def main():
    view = validate()
    walls = get_walls()

    # Optional: find a named line style
    line_style_id = get_line_style_id(LINE_STYLE_NAME)

    succeeded = 0
    skipped = []
    created_ids = []

    # Determine sketch plane for model line creation
    # In group edit mode we rely on the active sketch plane already set
    # by Revit when the user enters group edit mode.
    try:
        with revit.Transaction("Create Wall Face Lines"):
            for wall in walls:
                try:
                    ext_line, int_line = get_wall_face_curves(wall)

                    if ext_line is None and int_line is None:
                        logger.warning(
                            "Skipped wall {}: could not compute face curves"
                            .format(eid_int(wall.Id)))
                        skipped.append(eid_int(wall.Id))
                        continue

                    lines_created = 0

                    for face_line in [ext_line, int_line]:
                        if face_line is None:
                            continue
                        try:
                            # Create model line on the active sketch plane
                            sp = DB.SketchPlane.Create(
                                doc,
                                DB.Plane.CreateByNormalAndOrigin(
                                    DB.XYZ.BasisZ,
                                    face_line.GetEndPoint(0)
                                )
                            )
                            model_line = doc.Create.NewModelCurve(
                                face_line, sp
                            )

                            # Apply line style if available
                            if line_style_id and model_line:
                                try:
                                    model_line.LineStyle = doc.GetElement(
                                        line_style_id)
                                except Exception:
                                    pass  # non-critical; keep default style

                            if model_line:
                                created_ids.append(eid_int(model_line.Id))
                                lines_created += 1

                        except Exception as le:
                            logger.warning(
                                "Could not create line for wall {}: {}"
                                .format(eid_int(wall.Id), str(le)))

                    if lines_created > 0:
                        succeeded += 1
                    else:
                        skipped.append(eid_int(wall.Id))

                except Exception as we:
                    logger.warning(
                        "Skipped wall {}: {}".format(
                            eid_int(wall.Id), str(we)))
                    skipped.append(eid_int(wall.Id))

    except Exception as e:
        logger.error("Transaction failed: {}".format(str(e)))
        forms.alert(
            "Operation failed:\n{}".format(str(e)),
            title=TOOL_NAME,
            exitscript=True
        )

    # ── Summary ──────────────────────────────────────────────────────────
    total_lines = len(created_ids)
    msg = "{} model lines created for {} wall(s).".format(
        total_lines, succeeded)
    if skipped:
        msg += "\n{} wall(s) skipped.".format(len(skipped))
        logger.info("Skipped wall IDs: {}".format(
            ", ".join(str(i) for i in skipped)))

    forms.alert(msg, title=TOOL_NAME)


# ── Entry Point ──────────────────────────────────────────────────────────────
main()
