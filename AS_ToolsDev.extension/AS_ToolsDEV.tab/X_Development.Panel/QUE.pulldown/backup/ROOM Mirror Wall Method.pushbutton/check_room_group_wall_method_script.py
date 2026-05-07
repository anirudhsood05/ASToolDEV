# -*- coding: utf-8 -*-
"""Room Mirror Check - Wall-Based Method"""

__title__ = "ROOM Check Mirror in Group"
__doc__ = "Detects mirror status using wall elements in groups"

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()


class WallBasedAnalyzer:
    def __init__(self, document):
        self.doc = document
        self.group_cache = {}
    
    def get_wall_in_group(self, group_instance):
        """Find a wall in the group"""
        member_ids = group_instance.GetMemberIds()
        
        for member_id in member_ids:
            member = self.doc.GetElement(member_id)
            if isinstance(member, DB.Wall):
                return member
        return None
    
    def get_wall_direction_vector(self, wall):
        """Get wall direction from location curve"""
        try:
            location = wall.Location
            if isinstance(location, DB.LocationCurve):
                curve = location.Curve
                # Get direction at start
                start = curve.GetEndPoint(0)
                end = curve.GetEndPoint(1)
                direction = end.Subtract(start).Normalize()
                return direction
        except:
            pass
        return None
    
    def analyze_group_mirror(self, group_instance):
        """Analyze group using wall direction"""
        group_id = group_instance.Id.IntegerValue
        
        if group_id in self.group_cache:
            return self.group_cache[group_id]
        
        try:
            # Find wall in instance
            inst_wall = self.get_wall_in_group(group_instance)
            if not inst_wall:
                output.print_md("  - No wall found in group")
                return None
            
            inst_direction = self.get_wall_direction_vector(inst_wall)
            if not inst_direction:
                output.print_md("  - Could not get wall direction")
                return None
            
            output.print_md("  - Found wall: {}".format(inst_wall.Name))
            
            # Get group type
            group_type = self.doc.GetElement(group_instance.GetTypeId())
            
            # Create temp group
            temp_location = DB.XYZ(0, 0, 0)
            temp_group = self.doc.Create.PlaceGroup(temp_location, group_type)
            self.doc.Regenerate()
            
            # Find matching wall in temp
            temp_wall = None
            temp_member_ids = temp_group.GetMemberIds()
            
            for temp_id in temp_member_ids:
                temp_member = self.doc.GetElement(temp_id)
                if isinstance(temp_member, DB.Wall) and temp_member.Name == inst_wall.Name:
                    temp_wall = temp_member
                    break
            
            if not temp_wall:
                self.doc.Delete(temp_group.Id)
                output.print_md("  - No matching wall in type")
                return None
            
            type_direction = self.get_wall_direction_vector(temp_wall)
            self.doc.Delete(temp_group.Id)
            
            if not type_direction:
                output.print_md("  - Could not get type wall direction")
                return None
            
            # Calculate cross product with Z-axis to detect mirroring
            z_axis = DB.XYZ(0, 0, 1)
            
            inst_cross = inst_direction.CrossProduct(z_axis)
            type_cross = type_direction.CrossProduct(z_axis)
            
            # Dot product of perpendiculars
            dot = inst_cross.DotProduct(type_cross)
            
            output.print_md("  - Direction dot product: {:.6f}".format(dot))
            
            # Also check the direct dot product
            direct_dot = inst_direction.DotProduct(type_direction)
            output.print_md("  - Direct dot product: {:.6f}".format(direct_dot))
            
            # If perpendiculars point opposite directions, it's mirrored
            is_mirrored = dot < -0.5
            
            self.group_cache[group_id] = is_mirrored
            return is_mirrored
            
        except Exception as e:
            output.print_md("  - Error: {}".format(str(e)))
            return None


def main():
    output.print_md("# Room Mirror Check - Wall Method")
    output.print_md("---\n")
    
    analyzer = WallBasedAnalyzer(doc)
    
    # Get rooms in groups
    rooms = []
    collector = DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_Rooms
    ).WhereElementIsNotElementType()
    
    for room in collector:
        if room.GroupId != DB.ElementId.InvalidElementId:
            rooms.append(room)
    
    if not rooms:
        forms.alert("No rooms in groups", exitscript=True)
    
    output.print_md("Found {} rooms in groups\n".format(len(rooms)))
    
    # Get parameters
    text_params = []
    for param in rooms[0].Parameters:
        if param.StorageType == DB.StorageType.String and not param.IsReadOnly:
            text_params.append(param.Definition.Name)
    
    if not text_params:
        forms.alert("No writable text parameters", exitscript=True)
    
    selected_param = forms.SelectFromList.show(
        sorted(text_params),
        title="Select Parameter",
        button_name="Select"
    )
    
    if not selected_param:
        forms.alert("No parameter selected", exitscript=True)
    
    output.print_md("Writing to: **{}**\n".format(selected_param))
    output.print_md("## Processing\n---\n")
    
    results = {'processed': 0, 'mirrored': 0, 'not_mirrored': 0, 'failed': 0}
    
    with revit.Transaction("Write Mirror Status"):
        for room in rooms:
            results['processed'] += 1
            
            room_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
            if not room_name:
                room_name = "Room"
            
            group = doc.GetElement(room.GroupId)
            
            output.print_md("**{}** (ID: {})".format(room_name, room.Id.IntegerValue))
            
            is_mirrored = analyzer.analyze_group_mirror(group)
            
            if is_mirrored is None:
                results['failed'] += 1
                output.print_md("  - **FAILED**\n")
                continue
            
            mirror_text = "Mirrored" if is_mirrored else "Not Mirrored"
            
            param = room.LookupParameter(selected_param)
            if param and not param.IsReadOnly:
                param.Set(mirror_text)
                if is_mirrored:
                    results['mirrored'] += 1
                else:
                    results['not_mirrored'] += 1
                output.print_md("  - Result: **{}**\n".format(mirror_text))
            else:
                results['failed'] += 1
                output.print_md("  - **FAILED** - Parameter issue\n")
    
    output.print_md("\n## Summary")
    output.print_md("---")
    output.print_md("- Total: {}".format(results['processed']))
    output.print_md("- Mirrored: {}".format(results['mirrored']))
    output.print_md("- Not Mirrored: {}".format(results['not_mirrored']))
    output.print_md("- Failed: {}".format(results['failed']))


if __name__ == '__main__':
    main()
