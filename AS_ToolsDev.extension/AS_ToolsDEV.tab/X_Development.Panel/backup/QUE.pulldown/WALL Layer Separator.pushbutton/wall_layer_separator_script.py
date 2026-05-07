# -*- coding: utf-8 -*-
"""
Wall Layer Separation Tool - Ultra Simple Version
Just separates walls into layers, user handles hosted elements manually
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import Selection, ObjectType
from pyrevit import script, forms
from System.Collections.Generic import List

# Variables
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

class SimpleWallSeparator(object):
    def __init__(self):
        self.output = script.get_output()
        self.created_walls = 0
        self.processed_walls = 0
        self.errors = []
        self.skipped_layers = 0
    
    def get_element_name_safely(self, element):
        """Safely get element name."""
        try:
            return element.Name
        except:
            return "Unknown_Element_{}".format(element.Id.IntegerValue)
    
    def get_wall_type_by_name(self, wall_type_name):
        """Get existing WallType by name to avoid duplicates."""
        try:
            pvp = ParameterValueProvider(ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME))
            condition = FilterStringEquals()
            
            rvt_year = int(app.VersionNumber)
            if rvt_year < 2022:
                fRule = FilterStringRule(pvp, condition, wall_type_name, True)
            else:
                fRule = FilterStringRule(pvp, condition, wall_type_name)
            
            my_filter = ElementParameterFilter(fRule)
            
            return FilteredElementCollector(doc).OfClass(WallType)\
                .WherePasses(my_filter)\
                .FirstElement()
        except:
            return None
    
    def create_wall_type_for_layer(self, base_wall_type, layer_material, layer_thickness, layer_index):
        """Create new wall type for a specific material layer."""
        try:
            thickness_mm = layer_thickness * 304.8
            if thickness_mm < 0.1:
                self.output.print_md("Skipping layer {} - thickness too small ({:.3f}mm)".format(
                    layer_index + 1, thickness_mm))
                self.skipped_layers += 1
                return None
            
            base_type_name = self.get_element_name_safely(base_wall_type)
            
            if layer_material:
                material_name = self.get_element_name_safely(layer_material)
            else:
                material_name = "NoMaterial"
            
            # Clean material name
            material_name = material_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
            material_name = material_name.replace("(", "").replace(")", "").replace("-", "_")
            
            thickness_mm_int = int(thickness_mm)
            new_type_name = "{}_Layer{}_{}_{}mm".format(
                base_type_name, layer_index + 1, material_name, thickness_mm_int)
            
            # Limit name length
            if len(new_type_name) > 100:
                new_type_name = "Layer{}_{}_{}mm".format(layer_index + 1, material_name[:20], thickness_mm_int)
            
            # Check if exists
            existing_type = self.get_wall_type_by_name(new_type_name)
            if existing_type:
                return existing_type
            
            # Create new type
            new_wall_type = base_wall_type.Duplicate(new_type_name)
            
            # Create single-layer structure
            new_structure = CompoundStructure.CreateSingleLayerCompoundStructure(
                MaterialFunctionAssignment.Structure,
                layer_thickness,
                layer_material.Id if layer_material else ElementId.InvalidElementId)
            
            new_wall_type.SetCompoundStructure(new_structure)
            
            self.output.print_md("Created wall type: {}".format(new_type_name))
            return new_wall_type
            
        except Exception as e:
            error_msg = "Failed to create wall type for layer {}: {}".format(layer_index, str(e))
            self.errors.append(error_msg)
            self.output.print_md("**Error**: {}".format(error_msg))
            return None
    
    def calculate_layer_positions(self, wall, layers):
        """Calculate position offsets for each layer."""
        try:
            wall_location = wall.Location
            if not isinstance(wall_location, LocationCurve):
                return [XYZ(0, 0, 0)] * len(layers)
            
            wall_curve = wall_location.Curve
            wall_direction = (wall_curve.GetEndPoint(1) - wall_curve.GetEndPoint(0)).Normalize()
            wall_normal = XYZ(-wall_direction.Y, wall_direction.X, 0).Normalize()
            
            positions = []
            total_thickness = sum(layer.Width for layer in layers)
            current_offset = -total_thickness / 2.0
            
            for layer in layers:
                layer_center_offset = current_offset + layer.Width / 2.0
                position_offset = wall_normal * layer_center_offset
                positions.append(position_offset)
                current_offset += layer.Width
            
            return positions
            
        except Exception as e:
            self.output.print_md("**Warning**: Error calculating positions: {}".format(str(e)))
            return [XYZ(0, 0, 0)] * len(layers)
    
    def separate_wall_layers(self, wall):
        """Separate wall into layers - simple version."""
        try:
            wall_type = wall.WallType
            wall_type_name = self.get_element_name_safely(wall_type)
            
            self.output.print_md("## Processing Wall ID: {} ({})".format(
                wall.Id.IntegerValue, wall_type_name))
            
            # Get compound structure
            structure = wall_type.GetCompoundStructure()
            if not structure:
                self.output.print_md("**Error**: Wall has no compound structure")
                return False
            
            layers = structure.GetLayers()
            if not layers or len(layers) <= 1:
                self.output.print_md("**Info**: Wall has {} layers - no separation needed".format(len(layers)))
                return False
            
            self.output.print_md("Found {} layers to process".format(len(layers)))
            
            # Calculate positions
            layer_positions = self.calculate_layer_positions(wall, layers)
            
            # Get wall properties
            wall_location = wall.Location
            if not isinstance(wall_location, LocationCurve):
                self.output.print_md("**Error**: Wall has no location curve")
                return False
            
            wall_curve = wall_location.Curve
            wall_height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble()
            
            created_count = 0
            
            # Create walls for each layer
            for i, layer in enumerate(layers):
                self.output.print_md("Processing layer {} of {}".format(i + 1, len(layers)))
                
                # Get layer material
                mat_id = layer.MaterialId
                material = None
                material_name = "NoMaterial"
                
                if mat_id != ElementId(-1):
                    try:
                        material = doc.GetElement(mat_id)
                        if material:
                            material_name = self.get_element_name_safely(material)
                    except:
                        pass
                
                # Create wall type
                layer_wall_type = self.create_wall_type_for_layer(wall_type, material, layer.Width, i)
                if not layer_wall_type:
                    continue
                
                # Calculate wall position
                position_offset = layer_positions[i]
                
                # Create offset curve
                if position_offset.GetLength() > 0.001:
                    try:
                        offset_distance = position_offset.GetLength()
                        wall_direction = (wall_curve.GetEndPoint(1) - wall_curve.GetEndPoint(0)).Normalize()
                        wall_normal = XYZ(-wall_direction.Y, wall_direction.X, 0).Normalize()
                        
                        if position_offset.DotProduct(wall_normal) < 0:
                            offset_distance = -offset_distance
                        
                        layer_curve = wall_curve.CreateOffset(offset_distance, XYZ.BasisZ)
                    except:
                        layer_curve = wall_curve
                else:
                    layer_curve = wall_curve
                
                # Create new wall
                try:
                    new_wall = Wall.Create(doc, layer_curve, layer_wall_type.Id, wall.LevelId, wall_height, 0, False, False)
                    
                    if new_wall:
                        created_count += 1
                        self.created_walls += 1
                        
                        layer_thickness_mm = layer.Width * 304.8
                        self.output.print_md("SUCCESS: Layer {} - {} ({:.1f}mm) - Wall ID: {}".format(
                            i + 1, material_name, layer_thickness_mm, new_wall.Id.IntegerValue))
                    else:
                        self.output.print_md("**Error**: Failed to create wall for layer {}".format(i + 1))
                        
                except Exception as wall_error:
                    self.output.print_md("**Error**: Failed to create wall for layer {}: {}".format(
                        i + 1, str(wall_error)))
            
            if created_count > 0:
                self.processed_walls += 1
                return True
            else:
                return False
                
        except Exception as e:
            error_msg = "Failed to process wall {}: {}".format(wall.Id.IntegerValue, str(e))
            self.errors.append(error_msg)
            self.output.print_md("**Error**: {}".format(error_msg))
            return False
    
    def run(self):
        """Main execution function."""
        self.output.print_md("# Wall Layer Separation Tool - Ultra Simple")
        self.output.print_md("Creates individual walls for each material layer")
        self.output.print_md("**Note**: Hosted elements (doors/windows) must be moved manually")
        self.output.print_md("---")
        
        # Get wall selection
        selection = uidoc.Selection
        selected_el_ids = selection.GetElementIds()
        selected_el = [doc.GetElement(e_id) for e_id in selected_el_ids]
        
        if not selected_el:
            try:
                selected_el_refs = selection.PickObjects(ObjectType.Element, "Select walls to separate")
                selected_el = [doc.GetElement(ref.ElementId) for ref in selected_el_refs]
            except:
                forms.alert('No walls selected. Operation cancelled.', exitscript=True)
        
        # Filter walls
        selected_walls = [el for el in selected_el if isinstance(el, Wall)]
        
        if not selected_walls:
            forms.alert('No walls selected. Please select walls and try again.', exitscript=True)
        
        self.output.print_md("Selected {} walls for processing".format(len(selected_walls)))
        
        # Ask user about deleting original walls
        delete_originals = forms.alert(
            "Create separate walls for each material layer?\n\n"
            "This will create new walls at calculated positions.\n"
            "Original walls will remain (you can delete them manually).\n\n"
            "Continue?",
            title="Confirm Wall Separation",
            yes=True, no=True
        )
        
        if not delete_originals:
            self.output.print_md("Operation cancelled by user.")
            return
        
        # Process walls in single transaction
        t = Transaction(doc, "Create Wall Layer Separation")
        t.Start()
        
        try:
            # Process each wall
            for wall in selected_walls:
                self.separate_wall_layers(wall)
            
            # Commit transaction
            t.Commit()
            self.output.print_md("**Transaction committed successfully**")
            
        except Exception as e:
            t.RollBack()
            self.output.print_md("**Error**: Transaction failed: {}".format(str(e)))
            forms.alert("Operation failed: {}".format(str(e)))
            return
        
        # Final report
        self.output.print_md("---")
        self.output.print_md("## Summary")
        self.output.print_md("- **Processed walls**: {}".format(self.processed_walls))
        self.output.print_md("- **Created walls**: {}".format(self.created_walls))
        self.output.print_md("- **Skipped layers**: {} (zero thickness)".format(self.skipped_layers))
        self.output.print_md("- **Errors**: {}".format(len(self.errors)))
        
        if self.errors:
            self.output.print_md("### Errors:")
            for error in self.errors:
                self.output.print_md("- {}".format(error))
        
        # Instructions for user
        if self.created_walls > 0:
            self.output.print_md("---")
            self.output.print_md("## Next Steps")
            self.output.print_md("1. **Check the new walls** - they should be positioned correctly")
            self.output.print_md("2. **Handle hosted elements manually**:")
            self.output.print_md("   - Select doors/windows on original walls")
            self.output.print_md("   - Copy and paste to appropriate layer walls")
            self.output.print_md("   - Use 'Move' command to position precisely")
            self.output.print_md("3. **Delete original walls** when satisfied with results")
            
            forms.alert(
                "Wall separation completed!\n\n"
                "Created: {} new walls\n"
                "Processed: {} original walls\n\n"
                "You can now manually move hosted elements\n"
                "and delete original walls when ready.",
                title="Operation Complete"
            )

# Run the tool
if __name__ == '__main__':
    separator = SimpleWallSeparator()
    separator.run()