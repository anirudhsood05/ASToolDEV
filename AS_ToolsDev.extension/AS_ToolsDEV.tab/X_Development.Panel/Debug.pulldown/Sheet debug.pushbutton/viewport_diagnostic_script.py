# -*- coding: utf-8 -*-
"""
Viewport Placement Diagnostic
==============================
Open a sheet with at least one viewport already placed at your desired position.
This script reads that position, then creates a test floor plan view,
places it on the same sheet, and attempts SetBoxCenter.
Logs every step so we can see exactly where it breaks.

Run from the sheet view.
"""
from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()
view = doc.ActiveView

# ── Step 1: Verify we're on a sheet ──────────────────────────────────────────
if view.ViewType != DB.ViewType.DrawingSheet:
    forms.alert("Open a sheet first.", exitscript=True)

output.print_md("# Viewport Placement Diagnostic")
output.print_md("**Sheet:** {} - {}".format(view.SheetNumber, view.Name))
output.print_md("")

# ── Step 2: Read existing viewport positions ────────────────────────────────
vp_ids = view.GetAllViewports()
output.print_md("## Existing Viewports ({})".format(len(vp_ids)))

target_x = None
target_y = None

for vp_id in vp_ids:
    vp = doc.GetElement(vp_id)
    if not vp:
        continue
    center = vp.GetBoxCenter()
    outline = vp.GetBoxOutline()

    output.print_md("**{}**".format(vp.Name))
    output.print_md("- GetBoxCenter: X={:.6f} ft ({:.1f} mm), Y={:.6f} ft ({:.1f} mm)".format(
        center.X, center.X * 304.8,
        center.Y, center.Y * 304.8))

    if outline:
        o_min = outline.MinimumPoint
        o_max = outline.MaximumPoint
        output.print_md("- GetBoxOutline Min: X={:.6f}, Y={:.6f}".format(o_min.X, o_min.Y))
        output.print_md("- GetBoxOutline Max: X={:.6f}, Y={:.6f}".format(o_max.X, o_max.Y))

    # Use first viewport as target
    if target_x is None:
        target_x = center.X
        target_y = center.Y

output.print_md("")

if target_x is None:
    forms.alert("No viewports found on this sheet. Place one manually first.", exitscript=True)

output.print_md("**Target position from first viewport:** X={:.6f}, Y={:.6f}".format(
    target_x, target_y))
output.print_md("")

# ── Step 3: Find a floor plan view not already on this sheet ─────────────────
all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()

# Get views already on this sheet
views_on_sheet = set()
for vp_id in vp_ids:
    vp = doc.GetElement(vp_id)
    if vp:
        try:
            views_on_sheet.add(vp.ViewId.IntegerValue)
        except AttributeError:
            views_on_sheet.add(vp.ViewId.Value)

test_view = None
for v in all_views:
    if v.IsTemplate:
        continue
    if v.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan]:
        continue
    try:
        vid = v.Id.IntegerValue
    except AttributeError:
        vid = v.Id.Value
    if vid in views_on_sheet:
        continue
    # Check it's not already placed on another sheet
    try:
        sheet_num_param = v.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER)
        if sheet_num_param and sheet_num_param.AsString():
            continue
    except Exception:
        pass
    test_view = v
    break

if not test_view:
    forms.alert("No unplaced floor plan view found to test with.\n\n"
                "Create a spare floor plan view first.", exitscript=True)

output.print_md("**Test view:** {}".format(test_view.Name))
output.print_md("")

# ── Step 4: Place viewport and test positioning ─────────────────────────────
output.print_md("## Placement Test")

target = DB.XYZ(target_x, target_y, 0)

with revit.Transaction("Diagnostic - Place Viewport"):
    # Create viewport
    new_vp = DB.Viewport.Create(doc, view.Id, test_view.Id, target)

    if not new_vp:
        output.print_md("**FAIL:** Viewport.Create returned None")
    else:
        # Read position immediately after creation (same transaction)
        center_after_create = new_vp.GetBoxCenter()
        output.print_md("**After Create (before SetBoxCenter):**")
        output.print_md("- Center: X={:.6f}, Y={:.6f}".format(
            center_after_create.X, center_after_create.Y))
        output.print_md("- Delta from target: dX={:.6f}, dY={:.6f}".format(
            target_x - center_after_create.X,
            target_y - center_after_create.Y))
        output.print_md("")

        # Try SetBoxCenter
        try:
            new_vp.SetBoxCenter(target)
            output.print_md("**SetBoxCenter called successfully**")
        except Exception as e:
            output.print_md("**SetBoxCenter FAILED:** {}".format(str(e)))

        # Read position after SetBoxCenter (same transaction)
        center_after_set = new_vp.GetBoxCenter()
        output.print_md("**After SetBoxCenter:**")
        output.print_md("- Center: X={:.6f}, Y={:.6f}".format(
            center_after_set.X, center_after_set.Y))
        output.print_md("- Delta from target: dX={:.6f}, dY={:.6f}".format(
            target_x - center_after_set.X,
            target_y - center_after_set.Y))
        output.print_md("")

        # Check if position actually changed
        moved = (abs(center_after_set.X - center_after_create.X) > 0.001 or
                 abs(center_after_set.Y - center_after_create.Y) > 0.001)
        if moved:
            output.print_md("**RESULT: SetBoxCenter WORKED** - viewport moved")
        else:
            output.print_md("**RESULT: SetBoxCenter HAD NO EFFECT** - position unchanged")
            output.print_md("")
            output.print_md("This means the view template or crop region is locking "
                            "the viewport position. Try with a view that has no "
                            "view template applied.")

output.print_md("")
output.print_md("---")
output.print_md("*Delete the test viewport manually after reviewing results.*")
