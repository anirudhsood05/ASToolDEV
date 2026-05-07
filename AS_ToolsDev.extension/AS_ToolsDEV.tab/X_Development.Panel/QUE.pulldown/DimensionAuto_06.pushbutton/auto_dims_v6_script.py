# -*- coding: utf-8 -*-
__title__ = u"Auto Dims SC v6"
__doc__ = u"""Auto-dimensioning for walls, pylons, and structural columns.
v6 - Multi-grid support per element.

Usage:
1. Click the button
2. Box-select grids and elements (walls, columns)
3. Press Enter

Grid logic per element per axis:
- Finds ALL perpendicular grids that intersect/touch the element
- If none intersect, falls back to 2 nearest grids (one each side of centre)
- Builds a single combined chain: E_lo -> G1 -> G2 -> ... -> E_hi (row 1)
- Adds separate overall E->E (row 2) when any grid is strictly inside the element
- Grids on edges replace the edge ref in the chain (no duplicate at same position)
"""

import clr
import math

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Dimension, DimensionType, Grid, Wall, FamilyInstance,
    XYZ, Line, Arc, Reference, ReferenceArray,
    ElementId, Transaction, TransactionGroup,
    Options, ViewPlan,
    PlanarFace, Solid, GeometryInstance,
    FailureProcessingResult, IFailuresPreprocessor,
    BuiltInFailures, DatumEnds,
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, script


# -----------------------------------------
#  VERSION-SAFE ELEMENTID HELPER
# -----------------------------------------

def eid_int(element_id):
    """Return integer value of an ElementId. Compatible with Revit 2023-2026."""
    try:
        return element_id.Value          # Revit 2026+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2023-2025

INVALID_ID_INT = eid_int(ElementId.InvalidElementId)


# -----------------------------------------
#  FAILURE HANDLER (suppresses "not parallel")
# -----------------------------------------

class DimFailureSwallower(IFailuresPreprocessor):
    """Automatically deletes problematic dimensions instead of showing a dialog."""

    def __init__(self):
        self.had_errors = []

    def PreprocessFailures(self, failuresAccessor):
        failures = failuresAccessor.GetFailureMessages()
        for f in failures:
            try:
                sev = f.GetSeverity()
                desc = f.GetDescriptionText()
                if sev == sev.Error:
                    self.had_errors.append(desc)
                    ids = f.GetFailingElementIds()
                    if ids and ids.Count > 0:
                        failuresAccessor.DeleteElements(ids)
                    else:
                        failuresAccessor.ResolveFailure(f)
                elif sev == sev.Warning:
                    failuresAccessor.DeleteWarning(f)
            except Exception:
                try:
                    failuresAccessor.ResolveFailure(f)
                except Exception:
                    pass
        return FailureProcessingResult.Continue


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView
output = script.get_output()

# -----------------------------------------
#  SETTINGS
# -----------------------------------------

OFFSET_1_MM = 800   # first row (main dimensions/chains)
OFFSET_2_MM = 1400  # second row (overall E->E when grids inside)

OFFSET_CHAIN_1_MM = 1500   # pairwise grid chain
OFFSET_CHAIN_GAP_MM = 700  # gap between pairwise and overall grid chain

ZERO_TOL_MM = 5       # edge "on grid" tolerance
INTERSECT_TOL_MM = 50  # grid "intersects" element tolerance
MAX_SNAP_DIST_MM = 10000  # max distance to grid for fallback snapping

DEBUG = True


def mm_to_ft(mm):
    return mm / 304.8


def ft_to_mm(ft):
    return ft * 304.8


# -----------------------------------------
#  DATA COLLECTION
# -----------------------------------------

def collect_grids_from_selection(selected_elements):
    """Collects grids from selected elements. Determines bubble side."""
    grids = []
    for g in selected_elements:
        if not isinstance(g, Grid):
            continue
        try:
            crv = g.Curve
            if not isinstance(crv, Line):
                continue
            d = crv.Direction.Normalize()
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            if abs(d.Y) < 0.1:
                orientation = "horizontal"
                coord = (p0.Y + p1.Y) / 2.0
            elif abs(d.X) < 0.1:
                orientation = "vertical"
                coord = (p0.X + p1.X) / 2.0
            else:
                continue

            bubble_end = _get_bubble_end(g, p0, p1)

            grids.append({
                "element": g, "name": g.Name,
                "orientation": orientation, "coord_ft": coord,
                "p0": p0, "p1": p1,
                "bubble_end": bubble_end,
            })
        except Exception:
            continue
    return grids


def _get_bubble_end(grid, p0, p1):
    """Determines which end of the grid has the bubble."""
    try:
        b0 = grid.IsBubbleVisibleInView(DatumEnds.End0, view)
        b1 = grid.IsBubbleVisibleInView(DatumEnds.End1, view)
        if b1 and not b0:
            return "p1"
        return "p0"
    except Exception:
        return "p0"


def collect_elements_from_selection(selected_elements):
    """Collects walls and columns from selected elements."""
    elements = []
    seen_ids = set()
    wall_cat_id = eid_int(ElementId(BuiltInCategory.OST_Walls))
    str_col_cat_id = eid_int(ElementId(BuiltInCategory.OST_StructuralColumns))
    col_cat_id = eid_int(ElementId(BuiltInCategory.OST_Columns))

    for e in selected_elements:
        if isinstance(e, Grid):
            continue
        eid = eid_int(e.Id)
        if eid in seen_ids:
            continue
        seen_ids.add(eid)

        try:
            cat = e.Category
            if cat is None:
                continue
            cat_id = eid_int(cat.Id)
        except Exception:
            continue

        if cat_id == wall_cat_id:
            info = _bbox(e, "Wall")
            if info:
                elements.append(info)
        elif cat_id == str_col_cat_id or cat_id == col_cat_id:
            if isinstance(e, FamilyInstance) and e.SuperComponent is not None:
                continue
            info = _bbox(e, "Column")
            if info:
                elements.append(info)
    return elements


def _bbox(elem, cat):
    try:
        bb = elem.get_BoundingBox(view) or elem.get_BoundingBox(None)
        if not bb:
            return None
        w = abs(bb.Max.X - bb.Min.X)
        d = abs(bb.Max.Y - bb.Min.Y)
        if ft_to_mm(w) < 50 or ft_to_mm(d) < 50:
            return None
        return {
            "element": elem, "category": cat,
            "min_x": bb.Min.X, "max_x": bb.Max.X,
            "min_y": bb.Min.Y, "max_y": bb.Max.Y,
            "cx": (bb.Min.X + bb.Max.X) / 2.0,
            "cy": (bb.Min.Y + bb.Max.Y) / 2.0,
            "w_ft": w, "d_ft": d,
        }
    except Exception:
        return None


# -----------------------------------------
#  REFERENCES
# -----------------------------------------

def get_faces(elem, axis):
    """Gets face references of an element.

    Wall: references directly from Solid.
    FamilyInstance: GetSymbolGeometry() -> stable ref -> replace type id with instance id.
    """
    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = False
    opt.View = view
    geo = elem.get_Geometry(opt)
    if not geo:
        if DEBUG:
            output.print_md(u"   \u26a0 get_faces: no geometry (id={})".format(eid_int(elem.Id)))
        return None, None, None, None

    is_family = isinstance(elem, FamilyInstance)
    faces = []

    for item in geo:
        try:
            if isinstance(item, GeometryInstance) and is_family:
                xform = item.Transform
                sym_geo = item.GetSymbolGeometry()
                if not sym_geo:
                    continue
                for sym_item in sym_geo:
                    if not isinstance(sym_item, Solid) or sym_item.Faces.Size == 0:
                        continue
                    for face in sym_item.Faces:
                        if not isinstance(face, PlanarFace):
                            continue
                        sym_ref = face.Reference
                        if sym_ref is None:
                            continue
                        wn = xform.OfVector(face.FaceNormal)
                        wo = xform.OfPoint(face.Origin)
                        inst_ref = _symbol_to_instance_ref(sym_ref, elem)
                        if inst_ref is None:
                            continue
                        if axis == "x" and abs(wn.X) > 0.9:
                            faces.append((inst_ref, wo.X))
                        elif axis == "y" and abs(wn.Y) > 0.9:
                            faces.append((inst_ref, wo.Y))

            elif isinstance(item, Solid) and item.Faces.Size > 0:
                for face in item.Faces:
                    if not isinstance(face, PlanarFace):
                        continue
                    ref = face.Reference
                    if ref is None:
                        continue
                    n = face.FaceNormal
                    if axis == "x" and abs(n.X) > 0.9:
                        faces.append((ref, face.Origin.X))
                    elif axis == "y" and abs(n.Y) > 0.9:
                        faces.append((ref, face.Origin.Y))
        except Exception as ex:
            if DEBUG:
                output.print_md(u"   \u26a0 scan error: {}".format(str(ex)))
            continue

    if DEBUG:
        output.print_md(u"   \U0001f50d get_faces axis={}: {} faces, family={} (id={})".format(
            axis, len(faces), is_family, eid_int(elem.Id)))

    if len(faces) < 2:
        if DEBUG:
            _dump_normals(geo, is_family)
        return None, None, None, None

    faces.sort(key=lambda x: x[1])
    if DEBUG:
        output.print_md(u"   \U0001f4cf lo={:.0f}mm, hi={:.0f}mm".format(
            ft_to_mm(faces[0][1]), ft_to_mm(faces[-1][1])))
    return faces[0][0], faces[-1][0], faces[0][1], faces[-1][1]


def _symbol_to_instance_ref(sym_ref, instance):
    """Converts a reference from GetSymbolGeometry() to an instance reference."""
    try:
        stable = sym_ref.ConvertToStableRepresentation(doc)
        colon_idx = stable.index(":")
        new_stable = str(eid_int(instance.Id)) + stable[colon_idx:]
        inst_ref = Reference.ParseFromStableRepresentation(doc, new_stable)
        if DEBUG:
            output.print_md(u"      \U0001f517 ref: {} \u2192 {}".format(
                stable[:60], new_stable[:60]))
        return inst_ref
    except Exception as ex:
        if DEBUG:
            output.print_md(u"      \u274c ref convert failed: {}".format(str(ex)))
        return None


def _dump_normals(geo, is_family):
    """Dumps all normals for debugging."""
    all_n = []
    for item in geo:
        try:
            if isinstance(item, GeometryInstance) and is_family:
                xf = item.Transform
                for si in item.GetSymbolGeometry():
                    if isinstance(si, Solid):
                        for f in si.Faces:
                            if isinstance(f, PlanarFace):
                                wn = xf.OfVector(f.FaceNormal)
                                all_n.append(u"({:.2f},{:.2f},{:.2f})".format(wn.X, wn.Y, wn.Z))
            elif isinstance(item, Solid):
                for f in item.Faces:
                    if isinstance(f, PlanarFace):
                        all_n.append(u"({:.2f},{:.2f},{:.2f})".format(
                            f.FaceNormal.X, f.FaceNormal.Y, f.FaceNormal.Z))
        except Exception:
            pass
    output.print_md(u"   \U0001f9ca Normals: {}".format(u", ".join(all_n[:12])))


def get_grid_ref(grid):
    try:
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        opt.View = view
        geo = grid.get_Geometry(opt)
        if geo:
            for item in geo:
                if isinstance(item, Line) and item.Reference:
                    return item.Reference
        crv = grid.Curve
        if crv and crv.Reference:
            return crv.Reference
    except Exception:
        pass
    return None


# -----------------------------------------
#  DIMENSION CREATION
# -----------------------------------------

def make_dim(refs, p0, p1, label=""):
    if len(refs) < 2:
        return None
    ra = ReferenceArray()
    for r in refs:
        ra.Append(r)
    try:
        ln = Line.CreateBound(p0, p1)
        if DEBUG:
            d = ln.Direction.Normalize()
            output.print_md(
                u"   \U0001f4d0 make_dim [{}]: {} refs, line ({:.0f},{:.0f})->({:.0f},{:.0f}), dir=({:.2f},{:.2f})".format(
                    label, len(refs),
                    ft_to_mm(p0.X), ft_to_mm(p0.Y),
                    ft_to_mm(p1.X), ft_to_mm(p1.Y),
                    d.X, d.Y))
        dim = doc.Create.NewDimension(view, ln, ra)
        if DEBUG and dim:
            output.print_md(u"   \u2705 Dimension created (id={})".format(eid_int(dim.Id)))
        return dim
    except Exception as e:
        if DEBUG:
            output.print_md(u"   \u274c make_dim ERROR [{}]: **{}**".format(label, str(e)))
        return None


def _displace_small_texts(dim):
    try:
        scale = view.Scale
    except Exception:
        scale = 100

    text_width_mm = 5.0 * scale
    displace_mm = text_width_mm

    try:
        crv = dim.Curve
        if not crv or not isinstance(crv, Line):
            return
        direction = crv.Direction.Normalize()
    except Exception:
        return

    try:
        segs = list(dim.Segments)
        if segs and len(segs) > 0:
            for i, seg in enumerate(segs):
                try:
                    val = seg.Value
                    if val is None:
                        continue
                    val_mm = ft_to_mm(val)
                    if val_mm >= text_width_mm:
                        continue
                    if not seg.IsTextPositionAdjustable():
                        continue
                    tp = seg.TextPosition
                    if tp is None:
                        continue
                    sign = -1.0 if i == 0 else 1.0
                    offset_ft = mm_to_ft(displace_mm)
                    new_tp = XYZ(
                        tp.X + direction.X * offset_ft * sign,
                        tp.Y + direction.Y * offset_ft * sign,
                        tp.Z,
                    )
                    seg.TextPosition = new_tp
                except Exception:
                    continue
            return
    except Exception:
        pass

    try:
        val = dim.Value
        if val is None:
            return
        val_mm = ft_to_mm(val)
        if val_mm >= text_width_mm:
            return
        if not dim.IsTextPositionAdjustable():
            return
        tp = dim.TextPosition
        if tp is None:
            return
        offset_ft = mm_to_ft(displace_mm)
        new_tp = XYZ(
            tp.X + direction.X * offset_ft,
            tp.Y + direction.Y * offset_ft,
            tp.Z,
        )
        dim.TextPosition = new_tp
    except Exception:
        pass


# -----------------------------------------
#  MULTI-GRID FINDING
# -----------------------------------------

def _find_relevant_grids(ei, axis, grids_perpendicular):
    """Finds all relevant grids for an element along a given axis.

    Strategy:
    1. Collect ALL grids that intersect/touch the element (within INTERSECT_TOL_MM)
    2. If none intersect, fall back to 2 nearest grids (one each side of element centre)
    3. Filter out grids beyond MAX_SNAP_DIST_MM

    Returns list of grid dicts sorted by coord_ft, or empty list.
    """
    if not grids_perpendicular:
        return []

    if axis == "x":
        elem_center = ei["cx"]
        c_lo = ei["min_x"]
        c_hi = ei["max_x"]
    else:
        elem_center = ei["cy"]
        c_lo = ei["min_y"]
        c_hi = ei["max_y"]

    tol_inter = mm_to_ft(INTERSECT_TOL_MM)
    max_snap = mm_to_ft(MAX_SNAP_DIST_MM)

    # Pass 1: find all grids that intersect the element
    intersecting = []
    for g in grids_perpendicular:
        gc = g["coord_ft"]
        if (c_lo - tol_inter) < gc < (c_hi + tol_inter):
            intersecting.append(g)

    if intersecting:
        intersecting.sort(key=lambda g: g["coord_ft"])
        if DEBUG:
            names = u", ".join(g["name"] for g in intersecting)
            output.print_md(u"   \U0001f3af Intersecting grids: [{}]".format(names))
        return intersecting

    # Pass 2: no intersecting grids - find nearest on each side of centre
    best_lo = None  # nearest grid below/left of centre
    best_lo_dist = None
    best_hi = None  # nearest grid above/right of centre
    best_hi_dist = None

    for g in grids_perpendicular:
        gc = g["coord_ft"]
        dist = abs(gc - elem_center)
        if dist > max_snap:
            continue
        if gc <= elem_center:
            if best_lo_dist is None or dist < best_lo_dist:
                best_lo = g
                best_lo_dist = dist
        else:
            if best_hi_dist is None or dist < best_hi_dist:
                best_hi = g
                best_hi_dist = dist

    fallback = []
    if best_lo:
        fallback.append(best_lo)
    if best_hi:
        fallback.append(best_hi)

    # Deduplicate if same grid ended up on both sides (exactly at centre)
    seen = set()
    deduped = []
    for g in fallback:
        gid = eid_int(g["element"].Id)
        if gid not in seen:
            seen.add(gid)
            deduped.append(g)
    fallback = deduped

    fallback.sort(key=lambda g: g["coord_ft"])
    if DEBUG:
        if fallback:
            names = u", ".join(g["name"] for g in fallback)
            output.print_md(u"   \U0001f3af Fallback grids (nearest each side): [{}]".format(names))
        else:
            output.print_md(u"   \u26a0 No grids within {}mm".format(MAX_SNAP_DIST_MM))
    return fallback


# -----------------------------------------
#  CORE LOGIC
# -----------------------------------------

def dim_along_axis(ei, axis, grids_perpendicular, grids_parallel, all_elems, dims_to_adjust, forced_side=None,
                   occupied_zones=None):
    """Build dimensions for one element along one axis, using multiple grids."""
    elem = ei["element"]
    created = 0

    elem_name = elem.Name if hasattr(elem, 'Name') else '?'
    if DEBUG:
        output.print_md(u"---")
        output.print_md(u"### {} (id={}) axis={}  cat={}".format(
            elem_name, eid_int(elem.Id), axis, ei["category"]))
        output.print_md(u"   bbox: X[{:.0f}..{:.0f}] Y[{:.0f}..{:.0f}] mm".format(
            ft_to_mm(ei["min_x"]), ft_to_mm(ei["max_x"]),
            ft_to_mm(ei["min_y"]), ft_to_mm(ei["max_y"])))

    ref_lo, ref_hi, c_lo, c_hi = get_faces(elem, axis)
    if ref_lo is None:
        if DEBUG:
            output.print_md(u"   \u23ed Skipped \u2014 no faces found")
        return 0

    if axis == "x":
        perp_lo = ei["min_y"]
        perp_hi = ei["max_y"]
    else:
        perp_lo = ei["min_x"]
        perp_hi = ei["max_x"]

    side = _pick_side(ei, axis, grids_parallel, forced_side)

    off1 = mm_to_ft(OFFSET_1_MM)
    off2 = mm_to_ft(OFFSET_2_MM)

    if side < 0:
        line_row1 = perp_lo - off1
        line_row2 = perp_lo - off2
    else:
        line_row1 = perp_hi + off1
        line_row2 = perp_hi + off2

    # --- Find all relevant grids ---
    relevant_grids = _find_relevant_grids(ei, axis, grids_perpendicular)

    # No grids at all: overall only
    if not relevant_grids:
        if DEBUG:
            output.print_md(u"   \u2192 Overall only (no grids found)")
        dim_g = _dim_overall(ref_lo, ref_hi, c_lo, c_hi, axis, line_row1)
        if dim_g:
            dims_to_adjust.append(dim_g)
            created += 1
        return created

    # --- Resolve grid references ---
    tol_zero = mm_to_ft(ZERO_TOL_MM)
    tol_inter = mm_to_ft(INTERSECT_TOL_MM)

    grid_entries = []  # list of (coord_ft, reference, name, relationship)
    for g in relevant_grids:
        gref = get_grid_ref(g["element"])
        if gref is None:
            if DEBUG:
                output.print_md(u"   \u274c Failed to get Reference for grid {}".format(g["name"]))
            continue

        gc = g["coord_ft"]
        # Classify relationship to element
        inside = (c_lo - tol_inter) < gc < (c_hi + tol_inter)
        on_lo = inside and abs(c_lo - gc) <= tol_zero
        on_hi = inside and abs(c_hi - gc) <= tol_zero

        if on_lo:
            rel = "on_lo"
        elif on_hi:
            rel = "on_hi"
        elif inside:
            rel = "inside"
        else:
            rel = "outside"

        grid_entries.append({
            "coord": gc, "ref": gref, "name": g["name"], "rel": rel
        })

    if not grid_entries:
        # All grid refs failed
        dim_g = _dim_overall(ref_lo, ref_hi, c_lo, c_hi, axis, line_row1)
        if dim_g:
            dims_to_adjust.append(dim_g)
            created += 1
        return created

    grid_entries.sort(key=lambda x: x["coord"])

    if DEBUG:
        for ge in grid_entries:
            output.print_md(u"   \U0001f4cc grid **{}**: coord={:.0f}mm, rel={}".format(
                ge["name"], ft_to_mm(ge["coord"]), ge["rel"]))

    # --- Build combined chain: E_lo -> [G1 -> G2 -> ...] -> E_hi ---
    # Grids on lo edge replace ref_lo; grids on hi edge replace ref_hi
    has_inside_grids = any(ge["rel"] == "inside" for ge in grid_entries)
    has_outside_grids = any(ge["rel"] == "outside" for ge in grid_entries)

    # Start with lo edge ref (unless a grid sits exactly on it)
    lo_replaced = any(ge["rel"] == "on_lo" for ge in grid_entries)
    hi_replaced = any(ge["rel"] == "on_hi" for ge in grid_entries)

    # Build ordered ref list for the chain
    chain_refs = []
    chain_coords = []  # track all coords for line span

    # Add lo edge (or the grid that replaces it)
    if not lo_replaced:
        chain_refs.append(ref_lo)
        chain_coords.append(c_lo)
    # On-lo grids go first (they replace the lo edge)
    for ge in grid_entries:
        if ge["rel"] == "on_lo":
            chain_refs.append(ge["ref"])
            chain_coords.append(ge["coord"])

    # Add outside grids below element, then inside grids, then outside grids above
    for ge in grid_entries:
        if ge["rel"] == "outside" and ge["coord"] < c_lo:
            chain_refs.insert(0, ge["ref"])
            chain_coords.append(ge["coord"])
        elif ge["rel"] == "inside":
            chain_refs.append(ge["ref"])
            chain_coords.append(ge["coord"])
        elif ge["rel"] == "outside" and ge["coord"] > c_hi:
            # Will be added after hi edge
            pass  # handled below

    # Add hi edge (or the grid that replaces it)
    if not hi_replaced:
        chain_refs.append(ref_hi)
        chain_coords.append(c_hi)
    for ge in grid_entries:
        if ge["rel"] == "on_hi":
            chain_refs.append(ge["ref"])
            chain_coords.append(ge["coord"])

    # Add outside grids above element at the end
    for ge in grid_entries:
        if ge["rel"] == "outside" and ge["coord"] > c_hi:
            chain_refs.append(ge["ref"])
            chain_coords.append(ge["coord"])

    # Re-sort the chain by coordinate to ensure correct order
    # We need to pair refs with their coordinates for sorting
    # Rebuild properly: collect all (coord, ref) pairs and sort
    ref_coord_pairs = []

    # Lo edge or on_lo grids
    if not lo_replaced:
        ref_coord_pairs.append((c_lo, ref_lo, "edge_lo"))
    for ge in grid_entries:
        if ge["rel"] == "on_lo":
            ref_coord_pairs.append((ge["coord"], ge["ref"], "grid_" + ge["name"]))

    # Inside grids
    for ge in grid_entries:
        if ge["rel"] == "inside":
            ref_coord_pairs.append((ge["coord"], ge["ref"], "grid_" + ge["name"]))

    # Hi edge or on_hi grids
    if not hi_replaced:
        ref_coord_pairs.append((c_hi, ref_hi, "edge_hi"))
    for ge in grid_entries:
        if ge["rel"] == "on_hi":
            ref_coord_pairs.append((ge["coord"], ge["ref"], "grid_" + ge["name"]))

    # Outside grids (below or above element)
    for ge in grid_entries:
        if ge["rel"] == "outside":
            ref_coord_pairs.append((ge["coord"], ge["ref"], "grid_" + ge["name"]))

    # Sort by coordinate
    ref_coord_pairs.sort(key=lambda x: x[0])

    # Deduplicate refs at same coordinate (within tolerance)
    deduped_pairs = []
    for rcp in ref_coord_pairs:
        if deduped_pairs and abs(rcp[0] - deduped_pairs[-1][0]) < tol_zero:
            # Same position - prefer grid ref over edge ref
            if rcp[2].startswith("grid_"):
                deduped_pairs[-1] = rcp
            continue
        deduped_pairs.append(rcp)

    chain_refs = [p[1] for p in deduped_pairs]
    all_coords = [p[0] for p in deduped_pairs]

    if DEBUG:
        labels = [p[2] for p in deduped_pairs]
        output.print_md(u"   \U0001f517 Chain order: [{}] ({} refs)".format(
            u" \u2192 ".join(labels), len(chain_refs)))

    if len(chain_refs) < 2:
        if DEBUG:
            output.print_md(u"   \u23ed Skipped \u2014 fewer than 2 refs in chain")
        return 0

    # --- Span for dimension line ---
    span_lo = min(all_coords)
    span_hi = max(all_coords)

    # Collision check row1
    if occupied_zones is not None:
        line_row1 = _adjust_perp_for_collisions(
            axis, span_lo, span_hi, line_row1, side, occupied_zones)

    # Check collisions with other elements (for outside grids)
    if has_outside_grids:
        line_row1 = _avoid_collision(
            ei, line_row1, span_lo, span_hi,
            axis, side, all_elems or []
        )

    # Create the combined chain dimension (row 1)
    p0, p1 = _line_pts_from_span(span_lo, span_hi, axis, line_row1)
    label_parts = []
    for ge in grid_entries:
        label_parts.append(ge["name"])
    chain_label = u"chain[{}]".format(u"+".join(label_parts))

    dim_chain = make_dim(chain_refs, p0, p1, chain_label)
    if dim_chain:
        dims_to_adjust.append(dim_chain)
        created += 1
        if occupied_zones is not None:
            _register_zone(axis, span_lo, span_hi, line_row1, occupied_zones)

    # --- Row 2: overall E->E when grids are strictly inside ---
    if has_inside_grids:
        min_gap = mm_to_ft(OFFSET_2_MM - OFFSET_1_MM)
        if side < 0:
            line_row2 = min(line_row1 - min_gap, line_row2)
        else:
            line_row2 = max(line_row1 + min_gap, line_row2)
        if occupied_zones is not None:
            line_row2 = _adjust_perp_for_collisions(
                axis, c_lo, c_hi, line_row2, side, occupied_zones)

        dim_overall = _dim_overall(ref_lo, ref_hi, c_lo, c_hi, axis, line_row2)
        if dim_overall:
            dims_to_adjust.append(dim_overall)
            created += 1
            if occupied_zones is not None:
                _register_zone(axis, c_lo, c_hi, line_row2, occupied_zones)

    return created


def _dim_overall(ref_lo, ref_hi, c_lo, c_hi, axis, perp_pos):
    p0, p1 = _line_pts_from_span(min(c_lo, c_hi), max(c_lo, c_hi), axis, perp_pos)
    return make_dim([ref_lo, ref_hi], p0, p1, "overall")


def _line_pts_from_span(span_lo, span_hi, axis, perp_pos):
    """Creates line endpoints from a coordinate span and perpendicular position."""
    if axis == "x":
        return XYZ(span_lo, perp_pos, 0), XYZ(span_hi, perp_pos, 0)
    else:
        return XYZ(perp_pos, span_lo, 0), XYZ(perp_pos, span_hi, 0)


def _pick_side(ei, axis, grids_parallel, forced_side=None):
    """Chooses the side for the dimension line."""
    if forced_side is not None:
        if DEBUG:
            output.print_md(u"   \U0001f4cd _pick_side axis={}: forced side={}".format(axis, forced_side))
        return forced_side

    if not grids_parallel:
        return -1

    if axis == "x":
        elem_center = ei["cy"]
    else:
        elem_center = ei["cx"]

    best_grid = None
    best_d = None
    best_sign = 0
    for g in grids_parallel:
        d = g["coord_ft"] - elem_center
        abs_d = abs(d)
        if best_d is None or abs_d < best_d:
            best_d = abs_d
            best_grid = g
            best_sign = d

    if best_grid is None:
        return -1

    side = -1 if best_sign < 0 else +1
    if DEBUG:
        output.print_md(u"   \U0001f4cd _pick_side axis={}: parallel grid **{}** ({:.0f}mm), center={:.0f}mm \u2192 side={}".format(
            axis, best_grid["name"], ft_to_mm(best_grid["coord_ft"]),
            ft_to_mm(elem_center), side))
    return side


def _avoid_collision(ei, perp_pos, coord_lo, coord_hi, axis, side, all_elems):
    margin = mm_to_ft(200)
    my_id = eid_int(ei["element"].Id)
    lo = min(coord_lo, coord_hi)
    hi = max(coord_lo, coord_hi)

    for other in all_elems:
        if eid_int(other["element"].Id) == my_id:
            continue

        if axis == "x":
            if other["max_x"] < lo or other["min_x"] > hi:
                continue
            if other["min_y"] - margin < perp_pos < other["max_y"] + margin:
                if side < 0:
                    perp_pos = min(perp_pos, other["min_y"] - margin)
                else:
                    perp_pos = max(perp_pos, other["max_y"] + margin)
        else:
            if other["max_y"] < lo or other["min_y"] > hi:
                continue
            if other["min_x"] - margin < perp_pos < other["max_x"] + margin:
                if side < 0:
                    perp_pos = min(perp_pos, other["min_x"] - margin)
                else:
                    perp_pos = max(perp_pos, other["max_x"] + margin)

    return perp_pos


# -----------------------------------------
#  DIMENSION COLLISION PREVENTION
# -----------------------------------------

COLLISION_SHIFT_MM = 300
COLLISION_MAX_PASSES = 3


def _make_zone(axis, coord_lo, coord_hi, perp_pos, height_ft=None):
    if height_ft is None:
        height_ft = mm_to_ft(200)
    lo = min(coord_lo, coord_hi)
    hi = max(coord_lo, coord_hi)
    if axis == "x":
        return (lo, perp_pos - height_ft, hi, perp_pos + height_ft)
    else:
        return (perp_pos - height_ft, lo, perp_pos + height_ft, hi)


def _zone_overlaps(zone, occupied):
    for oz in occupied:
        if zone[2] <= oz[0] or oz[2] <= zone[0]:
            continue
        if zone[3] <= oz[1] or oz[3] <= zone[1]:
            continue
        return True
    return False


def _adjust_perp_for_collisions(axis, coord_lo, coord_hi, perp_pos, side, occupied):
    shift = mm_to_ft(COLLISION_SHIFT_MM)
    original_perp = perp_pos
    for attempt in range(COLLISION_MAX_PASSES):
        zone = _make_zone(axis, coord_lo, coord_hi, perp_pos)
        if not _zone_overlaps(zone, occupied):
            break
        if DEBUG:
            output.print_md(u"   \u26a0 COLLISION pass {}: perp={:.0f}mm, shifting".format(
                attempt + 1, ft_to_mm(perp_pos)))
        perp_pos += shift * side
    if DEBUG and perp_pos != original_perp:
        output.print_md(u"   \u2194 SHIFTED: {:.0f}mm \u2192 {:.0f}mm".format(
            ft_to_mm(original_perp), ft_to_mm(perp_pos)))
    return perp_pos


def _register_zone(axis, coord_lo, coord_hi, perp_pos, occupied):
    zone = _make_zone(axis, coord_lo, coord_hi, perp_pos)
    occupied.append(zone)


# -----------------------------------------
#  GRID CHAINS
# -----------------------------------------

def _grid_chain_exists(grids_sorted, measure_axis):
    """Checks if a dimension chain between the given grids already exists."""
    grid_ids = set()
    for g in grids_sorted:
        grid_ids.add(eid_int(g["element"].Id))

    if len(grid_ids) < 2:
        return False

    try:
        dims_on_view = FilteredElementCollector(doc, view.Id).OfClass(Dimension).ToElements()
    except Exception:
        return False

    for dim in dims_on_view:
        try:
            refs = dim.References
            if refs is None or refs.Size < 2:
                continue
            dim_ref_ids = set()
            for ref in refs:
                dim_ref_ids.add(eid_int(ref.ElementId))
            if grid_ids.issubset(dim_ref_ids):
                if DEBUG:
                    output.print_md(u"   \u23ed Grid chain already exists (dim id={})".format(
                        eid_int(dim.Id)))
                return True
        except Exception:
            continue

    return False


def make_grid_chain(grids_sorted, measure_axis, offset_mm):
    """Creates a dimension chain between grids."""
    if len(grids_sorted) < 2:
        return 0

    if _grid_chain_exists(grids_sorted, measure_axis):
        return 0

    refs = []
    for g in grids_sorted:
        r = get_grid_ref(g["element"])
        if r:
            refs.append(r)
    if len(refs) < 2:
        return 0

    bubble_coord, bubble_side = _get_bubble_baseline(grids_sorted, measure_axis)
    existing_offset_ft = _find_existing_grid_dim_offset(grids_sorted, measure_axis, bubble_side)

    off = mm_to_ft(offset_mm)

    if bubble_side > 0:
        base = max(bubble_coord, existing_offset_ft) if existing_offset_ft is not None else bubble_coord
        perp = base + off
    else:
        base = min(bubble_coord, existing_offset_ft) if existing_offset_ft is not None else bubble_coord
        perp = base - off

    if measure_axis == "x":
        p0 = XYZ(grids_sorted[0]["coord_ft"], perp, 0)
        p1 = XYZ(grids_sorted[-1]["coord_ft"], perp, 0)
    else:
        p0 = XYZ(perp, grids_sorted[0]["coord_ft"], 0)
        p1 = XYZ(perp, grids_sorted[-1]["coord_ft"], 0)

    if DEBUG:
        output.print_md(u"   \U0001f4cf chain {}: perp={:.0f}mm, bubble_side={}, base={:.0f}mm (exist={})".format(
            measure_axis, ft_to_mm(perp), bubble_side, ft_to_mm(bubble_coord),
            u"{:.0f}mm".format(ft_to_mm(existing_offset_ft)) if existing_offset_ft is not None else "none"))

    ra = ReferenceArray()
    for r in refs:
        ra.Append(r)
    try:
        dim = doc.Create.NewDimension(view, Line.CreateBound(p0, p1), ra)
        return 1 if dim else 0
    except Exception as e:
        if DEBUG:
            output.print_md(u"\u26a0 chain: {}".format(str(e)))
        return 0


def _get_bubble_baseline(grids_sorted, measure_axis):
    """Determines the bubble-end coordinate and shift direction."""
    bubble_coords = []
    non_bubble_coords = []

    for g in grids_sorted:
        be = g.get("bubble_end", "p0")
        bp = g[be]
        nbp = g["p1"] if be == "p0" else g["p0"]

        if measure_axis == "x":
            bubble_coords.append(bp.Y)
            non_bubble_coords.append(nbp.Y)
        else:
            bubble_coords.append(bp.X)
            non_bubble_coords.append(nbp.X)

    avg_bubble = sum(bubble_coords) / len(bubble_coords)
    avg_non_bubble = sum(non_bubble_coords) / len(non_bubble_coords)

    if avg_bubble > avg_non_bubble:
        bubble_edge = max(bubble_coords)
        return bubble_edge, -1
    else:
        bubble_edge = min(bubble_coords)
        return bubble_edge, +1


def _find_existing_grid_dim_offset(grids_sorted, measure_axis, side=1):
    """Finds the position of existing grid dimension chains to avoid overlap."""
    grid_ids = set(eid_int(g["element"].Id) for g in grids_sorted)

    try:
        dims_on_view = FilteredElementCollector(doc, view.Id).OfClass(Dimension).ToElements()
    except Exception:
        return None

    best_perp = None

    for dim in dims_on_view:
        try:
            refs = dim.References
            if refs is None or refs.Size < 2:
                continue
            match_count = 0
            for ref in refs:
                if eid_int(ref.ElementId) in grid_ids:
                    match_count += 1
            if match_count < 2:
                continue

            crv = dim.Curve
            if crv and isinstance(crv, Line):
                if measure_axis == "x":
                    if side > 0:
                        perp = max(crv.GetEndPoint(0).Y, crv.GetEndPoint(1).Y)
                        if best_perp is None or perp > best_perp:
                            best_perp = perp
                    else:
                        perp = min(crv.GetEndPoint(0).Y, crv.GetEndPoint(1).Y)
                        if best_perp is None or perp < best_perp:
                            best_perp = perp
                else:
                    if side > 0:
                        perp = max(crv.GetEndPoint(0).X, crv.GetEndPoint(1).X)
                        if best_perp is None or perp > best_perp:
                            best_perp = perp
                    else:
                        perp = min(crv.GetEndPoint(0).X, crv.GetEndPoint(1).X)
                        if best_perp is None or perp < best_perp:
                            best_perp = perp
        except Exception:
            continue

    return best_perp


# -----------------------------------------
#  MAIN
# -----------------------------------------

def main():
    if not isinstance(view, ViewPlan):
        forms.alert(u"Please open a plan view.", title=__title__)
        return

    # --- Box selection ---
    try:
        sel_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            u"Box-select grids, walls, and columns, then press Enter"
        )
    except Exception:
        return

    if not sel_refs:
        forms.alert(u"Nothing selected.", title=__title__)
        return

    selected_elements = []
    for r in sel_refs:
        try:
            el = doc.GetElement(r.ElementId)
            if el is not None:
                selected_elements.append(el)
        except Exception:
            continue

    if not selected_elements:
        forms.alert(u"Could not resolve any selected elements.", title=__title__)
        return

    all_grids = collect_grids_from_selection(selected_elements)
    all_elems = collect_elements_from_selection(selected_elements)

    h_grids = sorted([g for g in all_grids if g["orientation"] == "horizontal"],
                     key=lambda g: g["coord_ft"])
    v_grids = sorted([g for g in all_grids if g["orientation"] == "vertical"],
                     key=lambda g: g["coord_ft"])

    if DEBUG:
        n_walls = sum(1 for e in all_elems if e["category"] == "Wall")
        n_cols = sum(1 for e in all_elems if e["category"] == "Column")
        output.print_md(u"## Data (from selection)")
        output.print_md(u"- H grids: **{}** ({})".format(
            len(h_grids), u", ".join(g["name"] for g in h_grids)))
        output.print_md(u"- V grids: **{}** ({})".format(
            len(v_grids), u", ".join(g["name"] for g in v_grids)))
        output.print_md(u"- Elements: **{}** (walls: {}, columns: {})".format(
            len(all_elems), n_walls, n_cols))

    if not all_elems:
        forms.alert(u"No walls or columns in the selection.", title=__title__)
        return

    if not all_grids:
        forms.alert(u"No grids in the selection. Please select at least one grid.", title=__title__)
        return

    tg = TransactionGroup(doc, u"Auto Dims SC v6")
    tg.Start()
    total = 0
    failure_handler = DimFailureSwallower()

    try:
        # 1. Build grid chains
        t1 = Transaction(doc, u"Chains")
        opts1 = t1.GetFailureHandlingOptions()
        opts1.SetFailuresPreprocessor(failure_handler)
        t1.SetFailureHandlingOptions(opts1)
        t1.Start()
        try:
            n_chains = 0
            if len(v_grids) >= 2:
                n_chains += make_grid_chain(v_grids, "x", OFFSET_CHAIN_1_MM)
                if len(v_grids) > 2:
                    n_chains += make_grid_chain(
                        [v_grids[0], v_grids[-1]], "x",
                        OFFSET_CHAIN_1_MM + OFFSET_CHAIN_GAP_MM)
            if len(h_grids) >= 2:
                n_chains += make_grid_chain(h_grids, "y", OFFSET_CHAIN_1_MM)
                if len(h_grids) > 2:
                    n_chains += make_grid_chain(
                        [h_grids[0], h_grids[-1]], "y",
                        OFFSET_CHAIN_1_MM + OFFSET_CHAIN_GAP_MM)
            t1.Commit()
            total += n_chains
            if DEBUG:
                output.print_md(u"\u2705 Grid chains: **{}**".format(n_chains))
        except Exception as e:
            if t1.HasStarted() and not t1.HasEnded():
                t1.RollBack()
            if DEBUG:
                output.print_md(u"\u274c Grid chain transaction failed: {}".format(str(e)))

        # 2. Build element dimension chains
        t2 = Transaction(doc, u"Element Dims")
        opts2 = t2.GetFailureHandlingOptions()
        opts2.SetFailuresPreprocessor(failure_handler)
        t2.SetFailureHandlingOptions(opts2)
        t2.Start()
        try:
            n_x = 0
            n_y = 0
            dims_to_adjust = []
            occupied_zones = []

            for ei in all_elems:
                side_x = _pick_side(ei, "x", h_grids)
                side_y = side_x

                try:
                    n_x += dim_along_axis(ei, "x", v_grids, h_grids, all_elems, dims_to_adjust,
                                          forced_side=side_x, occupied_zones=occupied_zones)
                except Exception as e:
                    if DEBUG:
                        output.print_md(u"\u26a0 Error on X axis: {}".format(str(e)))
                try:
                    n_y += dim_along_axis(ei, "y", h_grids, v_grids, all_elems, dims_to_adjust,
                                          forced_side=side_y, occupied_zones=occupied_zones)
                except Exception as e:
                    if DEBUG:
                        output.print_md(u"\u26a0 Error on Y axis: {}".format(str(e)))

            doc.Regenerate()

            for d in dims_to_adjust:
                try:
                    _displace_small_texts(d)
                except Exception:
                    pass

            t2.Commit()
            total += n_x + n_y
            if DEBUG:
                output.print_md(u"\u2705 Dimensions along X: **{}**, along Y: **{}**".format(n_x, n_y))
                if occupied_zones:
                    output.print_md(u"\U0001f4e6 Reserved zones: **{}**".format(len(occupied_zones)))
                if failure_handler.had_errors:
                    output.print_md(u"\u26a0 Revit errors (auto-resolved): **{}**".format(len(failure_handler.had_errors)))
                    for err_msg in failure_handler.had_errors:
                        output.print_md(u"   - {}".format(err_msg))
        except Exception as e:
            if t2.HasStarted() and not t2.HasEnded():
                t2.RollBack()
            if DEBUG:
                output.print_md(u"\u274c Element dims transaction failed: {}".format(str(e)))

        tg.Assimilate()

    except Exception as e:
        try:
            tg.RollBack()
        except Exception:
            pass
        forms.alert(u"Error:\n{}".format(str(e)), title=__title__)
        return

    output.print_md(u"---")
    output.print_md(u"## Result: **{}** dimensions created".format(total))


if __name__ == "__main__":
    main()
