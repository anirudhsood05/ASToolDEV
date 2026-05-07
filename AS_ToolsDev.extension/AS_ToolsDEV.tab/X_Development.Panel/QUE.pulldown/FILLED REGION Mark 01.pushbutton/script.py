"""Filled Region Placeholders with Parameter Values
Creates filled regions with unique Mark or Comments parameter values.
"""

from pyrevit import revit, DB, forms, script
from pyrevit.framework import List
from Autodesk.Revit import Exceptions

__title__ = "Filled Region\nPlaceholders"
__author__ = "AUK Tools"

doc = revit.doc
view = revit.active_view

# Configuration (feet)
BOX_WIDTH = 2.5
BOX_HEIGHT = 1.5
COLUMN_SPACING = 3.5
ROW_SPACING = 2.0
COLUMNS = 4
TEXT_OFFSET_X = 0.2
TEXT_OFFSET_Y = 0.1


def get_filled_region_types():
    """Get all filled region types."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType)
    types = {}
    for fr_type in collector:
        name = fr_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        types[name] = fr_type
    return types


def get_parameter_values():
    """Get parameter values from user."""
    method = forms.CommandSwitchWindow.show(
        ["Enter manually (comma-separated)",
         "Generate alphabetic (A, B, C...)",
         "Generate numeric (1, 2, 3...)"],
        message="How to define parameter values?"
    )
    
    if not method:
        return None
    
    if "manually" in method:
        text = forms.ask_for_string(
            prompt="Enter values (comma-separated):",
            default="A, B, C, D, E"
        )
        return [v.strip() for v in text.split(",")] if text else None
    
    elif "alphabetic" in method:
        count = forms.ask_for_string(prompt="How many?", default="10")
        if count:
            n = min(int(count), 26)
            return [chr(65 + i) for i in range(n)]
    
    else:  # numeric
        count = forms.ask_for_string(prompt="How many?", default="10")
        if count:
            return [str(i + 1) for i in range(int(count))]
    
    return None


def create_rectangle(origin, width, height):
    """Create rectangle curve loop."""
    p1 = origin
    p2 = origin.Add(DB.XYZ(width, 0, 0))
    p3 = origin.Add(DB.XYZ(width, height, 0))
    p4 = origin.Add(DB.XYZ(0, height, 0))
    
    lines = [
        DB.Line.CreateBound(p1, p2),
        DB.Line.CreateBound(p2, p3),
        DB.Line.CreateBound(p3, p4),
        DB.Line.CreateBound(p4, p1)
    ]
    return lines


def scale_by_view(value):
    """Scale value by view scale."""
    return value * (float(view.Scale) / 100.0)


def main():
    # Validate view
    if view.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, 
                             DB.ViewType.Detail, DB.ViewType.DraftingView]:
        forms.alert("Use Floor Plan, Detail, or Drafting View.", exitscript=True)
    
    # Get filled region types
    fr_types = get_filled_region_types()
    if not fr_types:
        forms.alert("No filled region types found.", exitscript=True)
    
    # Select type
    selected_name = forms.SelectFromList.show(
        sorted(fr_types.keys()),
        title="Select Filled Region Type",
        button_name="Select"
    )
    if not selected_name:
        script.exit()
    
    selected_type = fr_types[selected_name]
    
    # Select parameter
    param_choice = forms.CommandSwitchWindow.show(
        ["Mark", "Comments"],
        message="Which parameter to use for values?"
    )
    if not param_choice:
        script.exit()
    
    # Get values
    values = get_parameter_values()
    if not values:
        script.exit()
    
    print("Creating {} filled regions with {} parameter".format(len(values), param_choice))
    
    # Get text style
    text_style_id = doc.GetDefaultElementTypeId(DB.ElementTypeGroup.TextNoteType)
    
    # Pick insertion point
    with forms.WarningBar(title="Pick insertion point"):
        try:
            insertion = revit.uidoc.Selection.PickPoint()
        except Exceptions.OperationCanceledException:
            script.exit()
    
    # Scale dimensions
    w = scale_by_view(BOX_WIDTH)
    h = scale_by_view(BOX_HEIGHT)
    col_space = scale_by_view(COLUMN_SPACING)
    row_space = scale_by_view(ROW_SPACING)
    txt_x = scale_by_view(TEXT_OFFSET_X)
    txt_y = scale_by_view(TEXT_OFFSET_Y)
    
    # Place filled regions
    placed = 0
    failed = 0
    row = 0
    col = 0
    
    with revit.Transaction("Place Filled Region Placeholders"):
        for value in values:
            # Calculate position
            x_off = col * col_space
            y_off = row * row_space
            origin = insertion.Add(DB.XYZ(x_off, y_off, 0))
            
            # Create rectangle
            curves = create_rectangle(origin, w, h)
            
            try:
                # Create filled region
                curve_loop = DB.CurveLoop.Create(List[DB.Curve](curves))
                region = DB.FilledRegion.Create(doc, selected_type.Id, view.Id, [curve_loop])
                
                # Set parameter
                if param_choice == "Mark":
                    param = region.LookupParameter("Mark")
                else:
                    param = region.LookupParameter("Comments")
                
                if param and not param.IsReadOnly:
                    param.Set(value)
                
                # Create label
                label_pos = origin.Add(DB.XYZ(txt_x, h - txt_y, 0))
                DB.TextNote.Create(doc, view.Id, label_pos, value, text_style_id)
                
                placed += 1
                print("[OK] {}".format(value))
                
            except Exception as e:
                failed += 1
                print("[FAILED] {}: {}".format(value, e))
            
            # Next position
            col += 1
            if col >= COLUMNS:
                col = 0
                row += 1
    
    # Summary
    print("\n{} placed, {} failed".format(placed, failed))
    forms.alert(
        "Completed!\n\nPlaced: {}\nFailed: {}".format(placed, failed),
        title="Done"
    )


if __name__ == '__main__':
    main()
