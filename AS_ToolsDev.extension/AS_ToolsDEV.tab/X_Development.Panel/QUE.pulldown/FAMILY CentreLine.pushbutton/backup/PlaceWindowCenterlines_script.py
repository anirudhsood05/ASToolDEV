# -*- coding: utf-8 -*-
"""Place Detail Lines at Window Centerlines
Creates detail lines at window family centerlines (vertical and horizontal) with 300mm extension.
"""

__title__ = "Window\nCenterlines"
__author__ = "AUK Digital"

from pyrevit import revit, DB, forms, script

# Constants
EXTENSION_MM = 300
EXTENSION_FEET = EXTENSION_MM / 304.8  # Convert mm to feet

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def get_line_style(style_name):
    """Get line style by name with fallback to default."""
    try:
        # Collect all graphic styles (includes line styles)
        collector = DB.FilteredElementCollector(doc)
        graphic_styles = collector.OfClass(DB.GraphicsStyle).ToElements()
        
        # Find the requested line style
        for style in graphic_styles:
            if style.Name == style_name and style.GraphicsStyleCategory.Name == "Lines":
                return style.Id
        
        # If not found, return InvalidElementId (uses default)
        output.print_md("**Line style '{}' not found, using default**".format(style_name))
        return DB.ElementId.InvalidElementId
        
    except Exception as e:
        output.print_md("**Warning**: Could not find line style: {}".format(str(e)))
        return DB.ElementId.InvalidElementId


def get_reference_line_geometry(window, ref_type):
    """Extract line geometry from family instance reference.
    
    Args:
        window: FamilyInstance element
        ref_type: FamilyInstanceReferenceType enum value
    
    Returns:
        Line object or None if not found
    """
    try:
        # Get references of the specified type (returns Python list)
        references = window.GetReferences(ref_type)
        
        if not references or len(references) == 0:
            return None
        
        # Get the first reference (should only be one for center references)
        reference = references[0]
        
        # Get the geometry object from the reference
        geom_obj = window.GetGeometryObjectFromReference(reference)
        
        if geom_obj and isinstance(geom_obj, DB.Line):
            return geom_obj
        
        return None
        
    except Exception as e:
        output.print_md("*Error getting reference geometry: {}*".format(str(e)))
        return None


def extend_line(line, extension):
    """Extend a line by specified distance on both ends.
    
    Args:
        line: DB.Line object
        extension: Distance to extend in feet
    
    Returns:
        New extended DB.Line
    """
    try:
        # Get line direction and normalize
        direction = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize()
        
        # Calculate new start and end points
        new_start = line.GetEndPoint(0) - (direction * extension)
        new_end = line.GetEndPoint(1) + (direction * extension)
        
        # Create new line
        return DB.Line.CreateBound(new_start, new_end)
        
    except Exception as e:
        output.print_md("*Error extending line: {}*".format(str(e)))
        return line  # Return original if extension fails


def create_detail_line(view, line, line_style_id):
    """Create a detail line in the specified view.
    
    Args:
        view: View to create the line in
        line: DB.Line geometry
        line_style_id: ElementId of line style (can be InvalidElementId for default)
    
    Returns:
        Created DetailLine element or None
    """
    try:
        # Create the detail curve
        detail_line = doc.Create.NewDetailCurve(view, line)
        
        # Apply line style if provided
        if line_style_id != DB.ElementId.InvalidElementId:
            try:
                detail_line.LineStyle = doc.GetElement(line_style_id)
            except:
                pass  # Use default if style assignment fails
        
        return detail_line
        
    except Exception as e:
        output.print_md("*Error creating detail line: {}*".format(str(e)))
        return None


def process_window(window, view, line_style_id, stats):
    """Process a single window to create centerlines.
    
    Args:
        window: FamilyInstance element (window)
        view: Active view
        line_style_id: ElementId of line style
        stats: Dictionary to track statistics
    """
    try:
        window_name = window.Symbol.FamilyName
        window_id = output.linkify(window.Id)
        
        # Track processing
        lines_created = 0
        
        # Try to get vertical centerline (Left/Right)
        vertical_line = get_reference_line_geometry(
            window, 
            DB.FamilyInstanceReferenceType.CenterLeftRight
        )
        
        if vertical_line:
            extended_line = extend_line(vertical_line, EXTENSION_FEET)
            detail_line = create_detail_line(view, extended_line, line_style_id)
            if detail_line:
                lines_created += 1
        
        # Try to get horizontal centerline (Front/Back)
        horizontal_line = get_reference_line_geometry(
            window,
            DB.FamilyInstanceReferenceType.CenterFrontBack
        )
        
        if horizontal_line:
            extended_line = extend_line(horizontal_line, EXTENSION_FEET)
            detail_line = create_detail_line(view, extended_line, line_style_id)
            if detail_line:
                lines_created += 1
        
        # Update statistics
        if lines_created > 0:
            stats['successful'] += 1
            stats['lines_created'] += lines_created
            output.print_md("✓ **{}** (ID: {}) - {} lines created".format(
                window_name, window_id, lines_created
            ))
        else:
            stats['failed'] += 1
            output.print_md("✗ **{}** (ID: {}) - No centerlines found".format(
                window_name, window_id
            ))
            
    except Exception as e:
        stats['failed'] += 1
        window_id = output.linkify(window.Id)
        output.print_md("✗ **Error** processing window ID {}: {}".format(
            window_id, str(e)
        ))


def main():
    """Main execution function."""
    
    # Get current selection
    selection = uidoc.Selection.GetElementIds()
    
    if not selection:
        forms.alert("Please select at least one window before running this tool.", 
                   title="No Selection", exitscript=True)
    
    # Filter selection to windows only
    selected_elements = [doc.GetElement(id) for id in selection]
    windows = [elem for elem in selected_elements 
               if isinstance(elem, DB.FamilyInstance) 
               and elem.Category 
               and elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Windows)]
    
    if not windows:
        forms.alert("No windows found in selection. Please select windows.", 
                   title="Invalid Selection", exitscript=True)
    
    # Get active view
    active_view = doc.ActiveView
    
    # Validate view type
    valid_view_types = [
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.Elevation,
        DB.ViewType.Section,
        DB.ViewType.Detail,
        DB.ViewType.DraftingView
    ]
    
    if active_view.ViewType not in valid_view_types:
        forms.alert(
            "This tool works in Plan, Elevation, Section, Detail, and Drafting views only.\n\n"
            "Current view type: {}".format(active_view.ViewType),
            title="Invalid View Type",
            exitscript=True
        )
    
    # Get line style
    line_style_id = get_line_style("Centerline")
    
    # Initialize statistics
    stats = {
        'successful': 0,
        'failed': 0,
        'lines_created': 0
    }
    
    # Print header
    output.print_md("# Window Centerlines Tool")
    output.print_md("---")
    output.print_md("**View**: {}".format(active_view.Name))
    output.print_md("**Windows Selected**: {}".format(len(windows)))
    output.print_md("**Extension**: {}mm each side".format(EXTENSION_MM))
    output.print_md("---\n")
    
    # Process windows in a transaction
    with revit.Transaction("Create Window Centerlines"):
        for window in windows:
            process_window(window, active_view, line_style_id, stats)
    
    # Print summary
    output.print_md("\n---")
    output.print_md("## Summary")
    output.print_md("**Successful**: {} windows".format(stats['successful']))
    output.print_md("**Failed**: {} windows".format(stats['failed']))
    output.print_md("**Total Detail Lines Created**: {}".format(stats['lines_created']))
    
    if stats['successful'] > 0:
        output.print_md("\n✓ **Operation completed successfully**")
    else:
        output.print_md("\n⚠ **Warning**: No centerlines were created. "
                       "Selected windows may not have centerline references defined in the family.")


if __name__ == "__main__":
    main()
