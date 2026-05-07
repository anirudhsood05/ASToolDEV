# -*- coding: utf-8 -*-
"""
Script:   Paint Material Tagger
Desc:     Annotates painted wall/floor/ceiling surfaces in plan view with material name TextNotes.
          Revit's native Material Tags cannot tag painted surfaces in plan view — this tool
          works around that limitation by detecting painted faces via the API, resolving the
          painted material name, and placing a TextNote at each element's centre.
Author:   Anirudh Sood — Aukett Swanke
Usage:    Open a Floor Plan or Ceiling Plan view. Optionally pre-select elements.
          If nothing is selected, the tool processes all visible walls in the active view.
Result:   TextNote annotations showing painted material names placed at element centres.
"""
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms

import clr
clr.AddReference("System")
from System.Collections.Generic import List

logger = script.get_logger()
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

# ── Constants ─────────────────────────────────────────────────────────────────
SUPPORTED_VIEW_TYPES = [
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.AreaPlan,
    DB.ViewType.EngineeringPlan,
]

# Categories to scan for paint
PAINTABLE_CATEGORIES = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Columns,
    DB.BuiltInCategory.OST_StructuralColumns,
]

# Maximum elements to process (safety limit)
MAX_ELEMENTS = 2000

# Tag offset from element centre (feet)
TAG_OFFSET_Y = 0.3


# ── Helpers ───────────────────────────────────────────────────────────────────
def eid_int(element_id):
    """Return integer value of an ElementId. Compatible Revit 2023-2026."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


INVALID_ID_INT = eid_int(DB.ElementId.InvalidElementId)


def get_element_name(element):
    """Safe element name access for IronPython."""
    try:
        return DB.Element.Name.GetValue(element) or ""
    except Exception:
        return ""


def get_element_centre(element, view):
    """Get the centre point of an element's bounding box in the active view."""
    try:
        bbox = element.get_BoundingBox(view)
        if bbox is None:
            return None
        centre = DB.XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
        return centre
    except Exception:
        return None


def get_painted_materials(element):
    """
    Detect all painted faces on an element and return a dict of
    {material_name: material_id} for unique painted materials.
    Uses Fine detail level to capture all geometry faces.
    """
    painted = {}
    try:
        opts = DB.Options()
        opts.DetailLevel = DB.ViewDetailLevel.Fine
        geo_elem = element.get_Geometry(opts)
        if geo_elem is None:
            return painted

        solids = []
        for geo_obj in geo_elem:
            if isinstance(geo_obj, DB.Solid) and geo_obj.Volume > 0:
                solids.append(geo_obj)
            elif isinstance(geo_obj, DB.GeometryInstance):
                inst_geo = geo_obj.GetInstanceGeometry()
                if inst_geo:
                    for inst_obj in inst_geo:
                        if isinstance(inst_obj, DB.Solid) and inst_obj.Volume > 0:
                            solids.append(inst_obj)

        for solid in solids:
            if solid.Faces is None:
                continue
            for face in solid.Faces:
                try:
                    if doc.IsPainted(element.Id, face):
                        mat_id = doc.GetPaintedMaterial(element.Id, face)
                        if mat_id and eid_int(mat_id) != INVALID_ID_INT:
                            mat_elem = doc.GetElement(mat_id)
                            if mat_elem:
                                mat_name = get_element_name(mat_elem)
                                if mat_name and mat_name not in painted:
                                    painted[mat_name] = mat_id
                except Exception:
                    continue
    except Exception as ex:
        logger.debug("Geometry error on {}: {}".format(eid_int(element.Id), str(ex)))

    return painted


def get_text_note_type(preferred_name=None):
    """
    Find a suitable TextNoteType. Tries preferred name first,
    then falls back to the project default.
    """
    all_types = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.TextNoteType)
        .WhereElementIsElementType()
        .ToElements()
    )
    if not all_types:
        return None

    # Try preferred name
    if preferred_name:
        for tnt in all_types:
            name = get_element_name(tnt)
            if name == preferred_name:
                return tnt

    # Try project default
    default_id = doc.GetDefaultElementTypeId(DB.ElementTypeGroup.TextNoteType)
    if default_id and eid_int(default_id) != INVALID_ID_INT:
        default_type = doc.GetElement(default_id)
        if default_type:
            return default_type

    # Last resort: first available
    return all_types[0]


def get_existing_paint_notes(view):
    """
    Collect existing TextNotes in the view that were created by this tool.
    Uses a marker prefix in the text content for identification.
    """
    marker = u"[PAINT] "
    existing = set()
    notes = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfClass(DB.TextNote)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    for note in notes:
        try:
            text = note.Text or ""
            if text.startswith(marker):
                existing.add(text)
        except Exception:
            continue
    return existing


# ── Validation ────────────────────────────────────────────────────────────────
def validate():
    """Pre-condition checks."""
    if doc.IsReadOnly:
        forms.alert("Document is read-only. Cannot place annotations.", exitscript=True)

    active_view = uidoc.ActiveView
    if active_view.ViewType not in SUPPORTED_VIEW_TYPES:
        forms.alert(
            "This tool works in Floor Plan, Ceiling Plan, Area Plan, "
            "or Engineering Plan views.\n\n"
            "Current view type: {}".format(active_view.ViewType),
            exitscript=True,
        )
    return active_view


# ── Main Logic ────────────────────────────────────────────────────────────────
def main():
    active_view = validate()

    # ── Determine elements to scan ────────────────────────────────────────────
    selection = revit.get_selection()
    if selection.elements:
        elements = list(selection.elements)
        logger.info("Using {} selected elements.".format(len(elements)))
    else:
        # Collect all paintable elements visible in the active view
        elements = []
        for bic in PAINTABLE_CATEGORIES:
            try:
                cat_elements = (
                    DB.FilteredElementCollector(doc, active_view.Id)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToElements()
                )
                elements.extend(cat_elements)
            except Exception:
                continue
        logger.info("Collected {} elements from active view.".format(len(elements)))

    if not elements:
        forms.alert("No paintable elements found in the active view.", exitscript=True)

    if len(elements) > MAX_ELEMENTS:
        if not forms.alert(
            "Found {} elements — this may take a while.\nContinue?".format(len(elements)),
            yes=True, no=True,
        ):
            script.exit()

    # ── Detect painted surfaces ───────────────────────────────────────────────
    painted_data = []  # list of (element, centre_pt, {mat_name: mat_id})
    skipped = 0

    for el in elements:
        try:
            mats = get_painted_materials(el)
            if not mats:
                continue
            centre = get_element_centre(el, active_view)
            if centre is None:
                skipped += 1
                continue
            painted_data.append((el, centre, mats))
        except Exception as ex:
            logger.warning("Skipped element {}: {}".format(eid_int(el.Id), str(ex)))
            skipped += 1

    if not painted_data:
        forms.alert(
            "No painted surfaces detected on the {} elements scanned.\n\n"
            "Ensure walls have been painted using Revit's Paint tool "
            "(Modify tab > Paint).".format(len(elements)),
            exitscript=True,
        )

    # ── Let user pick which materials to annotate ─────────────────────────────
    all_mat_names = set()
    for _, _, mats in painted_data:
        all_mat_names.update(mats.keys())

    all_mat_names = sorted(all_mat_names)
    if len(all_mat_names) > 1:
        selected_mats = forms.SelectFromList.show(
            all_mat_names,
            title="Select Paint Materials to Annotate",
            multiselect=True,
        )
        if not selected_mats:
            script.exit()
    else:
        selected_mats = all_mat_names

    selected_mats_set = set(selected_mats)

    # ── Choose text note type ─────────────────────────────────────────────────
    text_type = get_text_note_type()
    if text_type is None:
        forms.alert("No TextNote types found in the project.", exitscript=True)

    # ── Check for duplicates (avoid re-annotating) ────────────────────────────
    existing_notes = get_existing_paint_notes(active_view)

    # ── Place annotations ─────────────────────────────────────────────────────
    created = 0
    duplicate_skip = 0
    failed = []

    with revit.Transaction("AUK: Tag Painted Surfaces"):
        for el, centre, mats in painted_data:
            y_offset = 0.0
            for mat_name in sorted(mats.keys()):
                if mat_name not in selected_mats_set:
                    continue

                note_text = u"[PAINT] {}".format(mat_name)

                # Skip if identical note already exists
                if note_text in existing_notes:
                    duplicate_skip += 1
                    continue

                try:
                    tag_pt = DB.XYZ(
                        centre.X,
                        centre.Y + TAG_OFFSET_Y + y_offset,
                        centre.Z,
                    )
                    DB.TextNote.Create(
                        doc,
                        active_view.Id,
                        tag_pt,
                        note_text,
                        text_type.Id,
                    )
                    created += 1
                    existing_notes.add(note_text)
                    y_offset += TAG_OFFSET_Y
                except Exception as ex:
                    logger.warning(
                        "Failed to annotate element {}: {}".format(
                            eid_int(el.Id), str(ex)
                        )
                    )
                    failed.append(eid_int(el.Id))

    # ── Report ────────────────────────────────────────────────────────────────
    output.print_md("## Paint Material Tagger — Results")
    output.print_md("- **Elements scanned:** {}".format(len(elements)))
    output.print_md("- **Painted elements found:** {}".format(len(painted_data)))
    output.print_md("- **Annotations created:** {}".format(created))
    if duplicate_skip:
        output.print_md("- **Duplicates skipped:** {}".format(duplicate_skip))
    if skipped:
        output.print_md("- **Elements skipped (no geometry):** {}".format(skipped))
    if failed:
        output.print_md("- **Failed annotations:** {}".format(len(failed)))
        if len(failed) <= 10:
            output.print_md("  Failed IDs: {}".format(
                ", ".join(str(i) for i in failed)
            ))

    output.print_md("---")
    output.print_md("**Materials annotated:** {}".format(", ".join(sorted(selected_mats_set))))
    output.print_md("")
    output.print_md(
        u"_Notes are prefixed with `[PAINT]` for identification. "
        u"Re-running the tool will skip elements already annotated._"
    )


# ── Entry Point ───────────────────────────────────────────────────────────────
try:
    main()
except Exception as ex:
    logger.error("Unexpected error: {}".format(str(ex)))
    forms.alert(
        "Unexpected error — check pyRevit log for details.\n\n{}".format(str(ex)),
        exitscript=True,
    )
