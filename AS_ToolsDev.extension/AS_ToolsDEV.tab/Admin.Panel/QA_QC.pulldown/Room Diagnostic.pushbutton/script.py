# -*- coding: utf-8 -*-
"""Room Boundary Diagnostic Tool - ASCII Only"""

from pyrevit import forms, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class RoomSelectionFilter(ISelectionFilter):
    """Filter to only allow room selection"""
    def AllowElement(self, element):
        return isinstance(element, Room)
    
    def AllowReference(self, reference, point):
        return False

def diagnose_room_multiple_methods():
    """Try multiple methods to get room for diagnosis"""
    output = script.get_output()
    room = None
    
    # Method 1: Check if room is already selected
    try:
        selection_ids = uidoc.Selection.GetElementIds()
        if selection_ids:
            for elem_id in selection_ids:
                element = doc.GetElement(elem_id)
                if isinstance(element, Room):
                    room = element
                    output.print_md("**Using pre-selected room**")
                    break
    except:
        pass
    
    # Method 2: Interactive selection with filter
    if not room:
        try:
            output.print_md("**Please select a room when prompted...**")
            room_filter = RoomSelectionFilter()
            selected_ref = uidoc.Selection.PickObject(
                ObjectType.Element, 
                room_filter,
                "Select the problematic room (only rooms can be selected)"
            )
            room = doc.GetElement(selected_ref.ElementId)
            output.print_md("**Room selected interactively**")
        except Exception as e:
            output.print_md("Interactive selection failed: {}".format(str(e)))
    
    # Method 3: Pick from list if selection failed
    if not room:
        try:
            # Get all rooms
            all_rooms = FilteredElementCollector(doc)\
                .OfCategory(BuiltInCategory.OST_Rooms)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            if not all_rooms:
                forms.alert("No rooms found in the project")
                return
            
            # Create room options for selection
            room_options = []
            room_dict = {}
            
            for r in all_rooms:
                try:
                    room_name = r.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or "Unnamed"
                    room_number = r.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() or "No Number"
                    room_area = r.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble()
                    
                    display_text = "Room {} - {} ({:.0f} sq ft) [ID: {}]".format(
                        room_number, room_name, room_area, r.Id.IntegerValue
                    )
                    room_options.append(display_text)
                    room_dict[display_text] = r
                except:
                    display_text = "Room [ID: {}]".format(r.Id.IntegerValue)
                    room_options.append(display_text)
                    room_dict[display_text] = r
            
            # Show selection dialog
            selected_room_text = forms.SelectFromList.show(
                room_options,
                title="Select Room to Diagnose",
                width=500,
                height=400,
                multiselect=False
            )
            
            if selected_room_text:
                room = room_dict[selected_room_text]
                output.print_md("**Room selected from list**")
            
        except Exception as e:
            output.print_md("List selection failed: {}".format(str(e)))
    
    # If we still don't have a room, give up
    if not room:
        forms.alert("Could not select a room for diagnosis", "Selection Failed")
        return
    
    # Proceed with diagnosis
    diagnose_room(room)

def diagnose_room(room):
    """Diagnose the specified room"""
    output = script.get_output()
    
    output.print_md("# Room Diagnostic Report")
    output.print_md("Room ID: **{}**".format(room.Id.IntegerValue))
    
    # Basic room info
    try:
        room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or "Unnamed"
        room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() or "No Number"
        room_area = room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble()
        room_perimeter = room.get_Parameter(BuiltInParameter.ROOM_PERIMETER).AsDouble()
        room_height = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble()
        room_volume = room.get_Parameter(BuiltInParameter.ROOM_VOLUME).AsDouble()
        
        output.print_md("## Basic Information")
        output.print_md("- **Name**: {}".format(room_name))
        output.print_md("- **Number**: {}".format(room_number))
        output.print_md("- **Area**: {:.2f} sq ft".format(room_area))
        output.print_md("- **Perimeter**: {:.2f} ft".format(room_perimeter))
        output.print_md("- **Height**: {:.2f} ft".format(room_height))
        output.print_md("- **Volume**: {:.2f} cu ft".format(room_volume))
        
        # Warnings for problematic values
        if room_area < 1.0:
            output.print_md("**WARNING**: Room area is very small - possible calculation error")
        
        if room_perimeter < 1.0:
            output.print_md("**WARNING**: Room perimeter is very small - possible calculation error")
        
        if room_height <= 0:
            output.print_md("**ERROR**: Room height is zero or negative")
            
    except Exception as e:
        output.print_md("ERROR getting basic room info: {}".format(str(e)))
    
    # Level information
    try:
        level = room.Level
        if level:
            output.print_md("## Level Information")
            output.print_md("- **Level Name**: {}".format(level.Name))
            output.print_md("- **Level Elevation**: {:.2f} ft".format(level.Elevation))
            
            # Check room offset
            try:
                room_offset = room.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).AsDouble()
                room_limit_offset = room.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).AsDouble()
                output.print_md("- **Room Lower Offset**: {:.2f} ft".format(room_offset))
                output.print_md("- **Room Upper Offset**: {:.2f} ft".format(room_limit_offset))
            except:
                pass
        else:
            output.print_md("**ERROR**: Room has no associated level")
    except Exception as e:
        output.print_md("ERROR getting level info: {}".format(str(e)))
    
    # Location information
    try:
        location = room.Location
        if location and hasattr(location, 'Point'):
            point = location.Point
            output.print_md("## Location")
            output.print_md("- **X**: {:.3f} ft".format(point.X))
            output.print_md("- **Y**: {:.3f} ft".format(point.Y))
            output.print_md("- **Z**: {:.3f} ft".format(point.Z))
        else:
            output.print_md("**ERROR**: Room has no location point")
    except Exception as e:
        output.print_md("ERROR getting location: {}".format(str(e)))
    
    # Boundary analysis - This is the key part for your issue
    try:
        output.print_md("## Boundary Analysis")
        room_boundary_options = SpatialElementBoundaryOptions()
        all_boundaries = room.GetBoundarySegments(room_boundary_options)
        
        output.print_md("- **Number of Boundaries**: {}".format(len(all_boundaries)))
        
        if len(all_boundaries) == 0:
            output.print_md("**CRITICAL ERROR**: Room has NO boundary segments")
            output.print_md("**Possible causes:**")
            output.print_md("  - Room is not properly enclosed")
            output.print_md("  - Room calculation failed")
            output.print_md("  - Model corruption")
            return
        
        for boundary_index, boundary in enumerate(all_boundaries):
            boundary_type = "Exterior" if boundary_index == 0 else "Interior"
            output.print_md("### Boundary {} ({})".format(boundary_index, boundary_type))
            output.print_md("- **Segments**: {}".format(len(boundary)))
            
            if len(boundary) == 0:
                output.print_md("**ERROR**: Boundary has no segments")
                continue
            
            # Analyze each segment
            total_length = 0
            curve_info = []
            gaps = []
            problem_segments = []
            
            for i, segment in enumerate(boundary):
                try:
                    curve = segment.GetCurve()
                    if curve:
                        curve_length = curve.Length
                        total_length += curve_length
                        curve_type = type(curve).__name__
                        
                        start_pt = curve.GetEndPoint(0)
                        end_pt = curve.GetEndPoint(1)
                        
                        curve_info.append({
                            'index': i,
                            'type': curve_type,
                            'length': curve_length,
                            'start': start_pt,
                            'end': end_pt,
                            'curve': curve
                        })
                        
                        # Check for zero-length curves
                        if curve_length < 0.001:  # Less than ~0.012 inches
                            problem_segments.append("Segment {}: Zero-length curve".format(i))
                        
                        # Check gap to next segment
                        if i < len(boundary) - 1:
                            next_segment = boundary[i + 1]
                            next_curve = next_segment.GetCurve()
                            if next_curve:
                                gap = end_pt.DistanceTo(next_curve.GetEndPoint(0))
                                gaps.append(gap)
                                
                                if gap > 0.01:  # More than ~0.12 inches
                                    problem_segments.append("Gap after segment {}: {:.3f} inches".format(i, gap * 12))
                        else:
                            # Gap back to first segment (closing the loop)
                            if len(curve_info) > 0:
                                first_start = curve_info[0]['start']
                                gap = end_pt.DistanceTo(first_start)
                                gaps.append(gap)
                                
                                if gap > 0.01:
                                    problem_segments.append("Closing gap: {:.3f} inches".format(gap * 12))
                    else:
                        problem_segments.append("Segment {}: No curve found".format(i))
                        
                except Exception as seg_error:
                    problem_segments.append("Segment {}: {}".format(i, str(seg_error)))
            
            output.print_md("- **Total Length**: {:.3f} ft".format(total_length))
            
            # Report problems
            if problem_segments:
                output.print_md("#### **PROBLEMS DETECTED**:")
                for problem in problem_segments:
                    output.print_md("  - {}".format(problem))
            
            # Gap analysis
            if gaps:
                max_gap = max(gaps)
                total_gaps = sum(gaps)
                output.print_md("- **Maximum Gap**: {:.6f} ft ({:.3f} inches)".format(max_gap, max_gap * 12))
                output.print_md("- **Total Gaps**: {:.6f} ft ({:.3f} inches)".format(total_gaps, total_gaps * 12))
                
                if max_gap > 0.01:
                    output.print_md("**ERROR**: Large gaps detected - boundary not properly closed")
            
            # Show curve details
            if curve_info:
                output.print_md("#### Curve Details")
                for curve in curve_info:
                    output.print_md("- **Segment {}**: {} - {:.3f} ft".format(
                        curve['index'], curve['type'], curve['length']
                    ))
            
            # Test CurveLoop creation - THIS IS THE KEY TEST
            try:
                curves = [info['curve'] for info in curve_info if 'curve' in info]
                
                if curves:
                    test_loop = CurveLoop.Create(curves)
                    output.print_md("**CurveLoop Test**: SUCCESS")
                else:
                    output.print_md("**CurveLoop Test**: FAILED - No curves extracted")
                    
            except Exception as loop_error:
                output.print_md("**CurveLoop Test**: FAILED - {}".format(str(loop_error)))
                output.print_md("**This is why ceiling creation fails!**")
    
    except Exception as e:
        output.print_md("ERROR in boundary analysis: {}".format(str(e)))
    
    # Bounding box info
    try:
        bbox = room.get_BoundingBox(None)
        if bbox:
            output.print_md("## Bounding Box")
            output.print_md("- **Min**: ({:.3f}, {:.3f}, {:.3f})".format(bbox.Min.X, bbox.Min.Y, bbox.Min.Z))
            output.print_md("- **Max**: ({:.3f}, {:.3f}, {:.3f})".format(bbox.Max.X, bbox.Max.Y, bbox.Max.Z))
            
            width = bbox.Max.X - bbox.Min.X
            depth = bbox.Max.Y - bbox.Min.Y
            height = bbox.Max.Z - bbox.Min.Z
            
            output.print_md("- **Dimensions**: {:.3f} x {:.3f} x {:.3f} ft".format(width, depth, height))
            
            if width <= 0 or depth <= 0:
                output.print_md("**ERROR**: Invalid bounding box dimensions")
        else:
            output.print_md("**ERROR**: Room has no bounding box")
    except Exception as e:
        output.print_md("ERROR getting bounding box: {}".format(str(e)))
    
    # Recommendations
    output.print_md("## Recommended Fixes")
    output.print_md("Based on the analysis above:")
    output.print_md("1. **If CurveLoop test failed**: Use Room Separation Lines to redraw room boundary")
    output.print_md("2. **If large gaps detected**: Check wall connections and joins")
    output.print_md("3. **If zero-length curves**: Delete and recreate the room")
    output.print_md("4. **If no boundaries**: Check room is properly enclosed by walls/boundaries")
    output.print_md("5. **As last resort**: Use rectangular ceiling fallback in the main script")

if __name__ == '__main__':
    diagnose_room_multiple_methods()