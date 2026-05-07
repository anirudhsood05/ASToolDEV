# -*- coding: utf-8 -*-
"""
Script:   Purge Unused Scope Boxes
Desc:     Removes scope boxes not referenced by any view or datum element.
          Supports purging all unused scope boxes, from current selection,
          or from a user-defined list selection.
Author:   Aukett Swanke Digital
Usage:    Optionally pre-select scope boxes, then run. Choose purge mode.
Result:   Unused scope boxes are deleted; a summary is printed.
"""
__title__ = "Purge Unused\nScope Boxes"
__doc__ = "Removes scope boxes not used on any views or datum elements. " \
          "Supports purging all unused, from current selection, or via list picker."

from pyrevit import revit, DB, forms, script
from pyrevit.framework import List

logger = script.get_logger()
output = script.get_output()
doc    = revit.doc
uidoc  = revit.uidoc

# ── Constants ─────────────────────────────────────────────────────────────────
COMPATIBLE_VIEW_TYPES = [
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.EngineeringPlan,
    DB.ViewType.AreaPlan,
    DB.ViewType.Section,
    DB.ViewType.Elevation,
    DB.ViewType.Detail,
    DB.ViewType.ThreeD,
]

MODE_ALL        = "Purge all unused scope boxes"
MODE_SELECTION  = "Purge from current Revit selection"
MODE_LIST       = "Choose from a list"

# ── Helpers ───────────────────────────────────────────────────────────────────
def get_all_scope_boxes():
    """Return list of all scope box elements in the document."""
    return list(
        DB.FilteredElementCollector(doc)
          .OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest)
          .WhereElementIsNotElementType()
          .ToElements()
    )


def get_used_scope_box_ids():
    """Return a set of ElementIds of scope boxes currently referenced by views or datums."""
    used = set()

    # Views that support scope boxes
    all_views = (
        DB.FilteredElementCollector(doc)
          .OfClass(DB.View)
          .WhereElementIsNotElementType()
          .ToElements()
    )
    for view in all_views:
        if view.ViewType not in COMPATIBLE_VIEW_TYPES:
            continue
        try:
            param = view.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
            if param is None:
                continue
            sb_id = param.AsElementId()
            if sb_id and sb_id != DB.ElementId.InvalidElementId:
                used.add(sb_id)
        except Exception as e:
            logger.debug("Skipping view {}: {}".format(view.Id, str(e)))

    # Levels and grids (datum elements)
    cat_list = List[DB.BuiltInCategory]([
        DB.BuiltInCategory.OST_Levels,
        DB.BuiltInCategory.OST_Grids,
    ])
    datums = (
        DB.FilteredElementCollector(doc)
          .WherePasses(DB.ElementMulticategoryFilter(cat_list))
          .WhereElementIsNotElementType()
          .ToElements()
    )
    for datum in datums:
        try:
            param = datum.get_Parameter(DB.BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
            if param is None:
                continue
            sb_id = param.AsElementId()
            if sb_id and sb_id != DB.ElementId.InvalidElementId:
                used.add(sb_id)
        except Exception as e:
            logger.debug("Skipping datum {}: {}".format(datum.Id, str(e)))

    return used


def safe_name(element):
    """Return element Name safely, falling back to Id string."""
    try:
        return element.Name
    except Exception:
        return str(element.Id.IntegerValue)


def get_selection_scope_boxes():
    """Return scope boxes from the current Revit selection (may be empty)."""
    try:
        selection = revit.get_selection()
        return [
            e for e in selection.elements
            if e.Category is not None
            and e.Category.Id == DB.Category.GetCategory(
                doc, DB.BuiltInCategory.OST_VolumeOfInterest
            ).Id
        ]
    except Exception as e:
        logger.warning("Could not read selection: {}".format(str(e)))
        return []


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    # 1. Collect all scope boxes
    all_scope_boxes = get_all_scope_boxes()
    if not all_scope_boxes:
        forms.alert("No scope boxes found in the model.", exitscript=True)

    # 2. Determine which are used
    used_ids = get_used_scope_box_ids()

    # 3. All unused
    unused_all = [sb for sb in all_scope_boxes if sb.Id not in used_ids]

    if not unused_all:
        forms.alert("All scope boxes are in use. Nothing to purge.", exitscript=True)

    # 4. Ask user for mode
    mode = forms.CommandSwitchWindow.show(
        [MODE_ALL, MODE_SELECTION, MODE_LIST],
        message="Select purge mode ({} unused scope box{} found):".format(
            len(unused_all), "es" if len(unused_all) != 1 else ""
        ),
    )
    if not mode:
        script.exit()

    # 5. Resolve candidates based on mode
    candidates = []

    if mode == MODE_ALL:
        candidates = unused_all

    elif mode == MODE_SELECTION:
        sel_sbs = get_selection_scope_boxes()
        if not sel_sbs:
            forms.alert(
                "No scope boxes found in the current selection.\n"
                "Please select scope boxes in Revit before running.",
                exitscript=True,
            )
        # Intersect selection with unused
        unused_sel_ids = {sb.Id for sb in unused_all}
        candidates = [sb for sb in sel_sbs if sb.Id in unused_sel_ids]
        in_use_count = len(sel_sbs) - len(candidates)
        if in_use_count:
            logger.info(
                "{} selected scope box(es) are in use and will be skipped.".format(in_use_count)
            )
        if not candidates:
            forms.alert(
                "All selected scope boxes are currently in use. Nothing to purge.",
                exitscript=True,
            )

    elif mode == MODE_LIST:
        # Present unused scope boxes for user selection
        name_map = {safe_name(sb): sb for sb in unused_all}
        chosen_names = forms.SelectFromList.show(
            sorted(name_map.keys()),
            title="Select Unused Scope Boxes to Purge",
            multiselect=True,
        )
        if not chosen_names:
            script.exit()
        candidates = [name_map[n] for n in chosen_names if n in name_map]

    if not candidates:
        forms.alert("No scope boxes selected for deletion.", exitscript=True)

    # 6. Confirm
    confirm_msg = (
        "You are about to delete {} scope box{}:\n\n{}\n\nThis cannot be undone. Continue?".format(
            len(candidates),
            "es" if len(candidates) != 1 else "",
            "\n".join("  \u2022 " + safe_name(sb) for sb in candidates),
        )
    )
    if not forms.alert(confirm_msg, ok=False, yes=True, no=True):
        script.exit()

    # 7. Delete
    deleted   = []
    failed    = []

    with revit.Transaction("Purge unused scope boxes"):
        for sb in candidates:
            try:
                name = safe_name(sb)
                doc.Delete(sb.Id)
                deleted.append(name)
            except Exception as e:
                logger.warning("Could not delete scope box {}: {}".format(safe_name(sb), str(e)))
                failed.append(safe_name(sb))

    # 8. Report
    output.print_md("## Purge Unused Scope Boxes \u2014 Results")
    output.print_md("**Deleted:** {}  |  **Failed:** {}".format(len(deleted), len(failed)))

    if deleted:
        output.print_md("### Deleted")
        for name in sorted(deleted):
            output.print_md("- {}".format(name))

    if failed:
        output.print_md("### Could Not Delete")
        for name in sorted(failed):
            output.print_md("- {}".format(name))


main()