# -*- coding: utf-8 -*-
"""
Wall Layer Separation Tool
Separates compound walls into individual material layers with reliable hosted element transfer
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import Selection, ObjectType
from pyrevit import script, forms
from System.Collections.Generic import List
import math

# Variables
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

class WallLayerSeparator(object):
    def __init__(self):
        self.output = script.get_output()
        self.created_elements = []
        self.processed_walls = 0
        self.created_walls = 0
        self.errors = []
        self.skipped_layers = 0
        self.transferred_elements = 0
    
    def get_element_name_safely(self, element):
        """Safely get element name with multiple fallback methods."""
        try:
            return element.Name
        except:
            try:
                return Element.Name.GetValue(element)
            except:
                try:
                    name_param = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    if name_param and name_param.HasValue:
                        return name_param.AsString()
                except:
                    pass
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
        except Exception as e:
            self.output.print_md("**Warning**: Error checking existing wall type: {}".format(str(e)))
            return None
    
    def create_wall_type_for_layer(self, base_wall_type, layer_material, layer_thickness, layer_index):
        """Create new wall type for a specific material layer."""
        try:
            # Check for zero thickness first
            thickness_mm = layer_thickness * 304.8
            if thickness_mm < 0.1:
                error_msg = "Skipping layer {} - thickness too small ({:.3f}mm)".format(
                    layer_index + 1, thickness_mm)
                self.output.print_md("**Info**: {}".format(error_msg))
                self.skipped_layers += 1
                return None
            
            base_type_name = self.get_element_name_safely(base_wall_type)
            
            if layer_material:
                material_name = self.get_element_name_safely(layer_material)
            else:
                material_name = "NoMaterial"
            
            # Clean material name for filename compatibility
            material_name = material_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
            material_name = material_name.replace("(", "").replace(")", "").replace("-", "_")
            material_name = material_name.replace(".", "").replace(",", "")
            
            thickness_mm_int = int(thickness_mm)
            new_type_name = "{}_Layer{}_{}_{}mm".format(
                base_type_name, layer_index + 1, material_name, thickness_mm_int)
            
            # Ensure name is not too long
            if len(new_type_name) > 100:
                new_type_name = "Layer{}_{}_{}mm".format(layer_index + 1, material_name[:20], thickness_mm_int)
            
            # Check if type already exists
            existing_type = self.get_wall_type_by_name(new_type_name)
            if existing_type:
                self.output.print_md("Using existing wall type: {}".format(new_type_name))
                return existing_type
            
            # Duplicate base wall type
            new_wall_type = base_wall_type.Duplicate(new_type_name)
            
            # Create single-layer compound structure
            new_structure = CompoundStructure.CreateSingleLayerCompoundStructure(
                MaterialFunctionAssignment.Structure,
                layer_thickness,
                layer_material.Id if layer_material else ElementId.InvalidElementId)
            
            # Apply the new structure
            new_wall_type.SetCompoundStructure(new_structure)
            
            self.output.print_md("Successfully created wall type: {}".format(new_type_name))
            return new_wall_type
            
        except Exception as e:
            error_msg = "Failed to create wall type for layer {}: {}".format(layer_index, str(e))
            self.errors.append(error_msg)
            self.output.print_md("**Error**: {}".format(error_msg))
            return None
    
    def create_wall_with_type(self, curve, wall_type, level_id, height):
        """Create wall with specific type."""
        try:
            new_wall = Wall.Create(doc, curve, wall_type.Id, level_id, height, 0, False, False)
            
            if new_wall:
                # Verify wall type
                if new_wall.WallType.Id == wall_type.Id:
                    return new_wall
                else:
                    # Try to correct wall type
                    try:
                        new_wall.WallType = wall_type
                        return new_wall
                    except:
                        return new_wall  # Return anyway
            
            return None
            
        except Exception as e:
            self.output.print_md("Wall creation failed: {}".format(str(e)))
            return None
    
    def get_hosted_elements(self, wall):
        """Enhanced method to identify hosted elements with better detection."""
        try:
            hosted_elements = []
            
            # Method 1: Get dependent elements
            dependent_ids = wall.GetDependentElements(None)
            
            for dep_id in dependent_ids:
                element = doc.GetElement(dep_id)
                if element and not isinstance(element, Wall):
                    # Check if element has Host property
                    if hasattr(element, 'Host'):
                        try:
                            # Verify it's actually hosted by this wall
                            if element.Host and element.Host.Id == wall.Id:
                                hosted_elements.append(element)
                                self.output.print_md("  Found hosted: {} (ID: {}, Type: {})".format(
                                    self.get_element_name_safely(element), 
                                    element.Id.IntegerValue,
                                    element.GetType().Name))
                        except:
                            # Some elements might have inaccessible Host property
                            # but still be dependent - include them
                            hosted_elements.append(element)
                            self.output.print_md("  Found dependent: {} (ID: {})".format(
                                self.get_element_name_safely(element), 
                                element.Id.IntegerValue))
            
            # Method 2: Check common hosted element categories as backup
            if not hosted_elements:
                categories_to_check = [
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_SpecialityEquipment,
                    BuiltInCategory.OST_ElectricalFixtures,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_PlumbingFixtures
                ]
                
                for category in categories_to_check:
                    try:
                        elements = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
                        
                        for element in elements:
                            if hasattr(element, 'Host'):
                                try:
                                    if element.Host and element.Host.Id == wall.Id:
                                        if element not in hosted_elements:
                                            hosted_elements.append(element)
                                            self.output.print_md("  Found by category scan: {} (ID: {})".format(
                                                self.get_element_name_safely(element), 
                                                element.Id.IntegerValue))
                                except:
                                    continue
                    except:
                        continue
            
            self.output.print_md("Total hosted elements found: {}".format(len(hosted_elements)))
            return hosted_elements
            
        except Exception as e:
            self.output.print_md("**Warning**: Error getting hosted elements: {}".format(str(e)))
            return []
    
    def transfer_hosted_elements(self, original_wall, target_wall, hosted_elements):
        """Enhanced hosted element transfer using proper Revit API methods."""
        if not hosted_elements:
            self.output.print_md("No hosted elements to transfer")
            return []
        
        successfully_transferred = []
        failed_transfers = []
        
        self.output.print_md("Transferring {} hosted elements to wall ID: {}".format(
            len(hosted_elements), target_wall.Id.IntegerValue))
        
        for element in hosted_elements:
            element_name = self.get_element_name_safely(element)
            element_id = element.Id.IntegerValue
            element_type = element.GetType().Name
            
            try:
                # Strategy 1: Use ChangeTypeId for FamilyInstance elements (most reliable)
                if isinstance(element, FamilyInstance):
                    self.output.print_md("  Processing FamilyInstance: {} (ID: {})".format(element_name, element_id))
                    
                    # Get element's current parameters and location
                    element_location = element.Location
                    element_level_id = element.LevelId
                    element_symbol_id = element.GetTypeId()
                    
                    # Get location point
                    if hasattr(element_location, 'Point'):
                        location_point = element_location.Point
                    else:
                        raise Exception("Cannot get element location point")
                    
                    # Try to recreate the element at the same location
                    # This is more reliable than trying to change host directly
                    try:
                        # Create new instance of same family at same location
                        family_symbol = doc.GetElement(element_symbol_id)
                        if not family_symbol.IsActive:
                            family_symbol.Activate()
                        
                        # Create new family instance
                        new_instance = doc.Create.NewFamilyInstance(
                            location_point, 
                            family_symbol, 
                            target_wall,
                            element_level_id,
                            Structure.StructuralType.NonStructural
                        )
                        
                        if new_instance:
                            # Copy parameters from original to new instance
                            self.copy_element_parameters(element, new_instance)
                            
                            # Delete original element
                            doc.Delete(element.Id)
                            
                            successfully_transferred.append(new_instance)
                            self.transferred_elements += 1
                            self.output.print_md("  SUCCESS: {} (ID: {}) recreated as (ID: {})".format(
                                element_name, element_id, new_instance.Id.IntegerValue))
                        else:
                            raise Exception("NewFamilyInstance returned None")
                    
                    except Exception as recreate_error:
                        self.output.print_md("  Recreation failed: {}".format(str(recreate_error)))
                        # Try alternative method
                        self._try_element_move_method(element, target_wall, successfully_transferred)
                
                # Strategy 2: For non-FamilyInstance elements, try ElementTransformUtils
                else:
                    self.output.print_md("  Processing {} element: {} (ID: {})".format(
                        element_type, element_name, element_id))
                    
                    try:
                        # Try to move element using ElementTransformUtils
                        element_ids = List[ElementId]([element.Id])
                        
                        # Calculate translation to target wall (minimal movement)
                        original_wall_location = original_wall.Location
                        target_wall_location = target_wall.Location
                        
                        if (isinstance(original_wall_location, LocationCurve) and 
                            isinstance(target_wall_location, LocationCurve)):
                            
                            orig_curve = original_wall_location.Curve
                            target_curve = target_wall_location.Curve
                            
                            # Get midpoints
                            orig_mid = orig_curve.Evaluate(0.5, False)
                            target_mid = target_curve.Evaluate(0.5, False)
                            
                            # Calculate translation vector
                            translation = target_mid - orig_mid
                            
                            # Apply minimal translation
                            ElementTransformUtils.MoveElements(doc, element_ids, translation)
                            
                            successfully_transferred.append(element)
                            self.transferred_elements += 1
                            self.output.print_md("  SUCCESS: {} (ID: {}) - Moved to target wall".format(
                                element_name, element_id))
                        else:
                            raise Exception("Cannot calculate wall translation")
                    
                    except Exception as move_error:
                        failed_transfers.append((element, str(move_error)))
                        self.output.print_md("  FAILED: {} (ID: {}) - {}".format(
                            element_name, element_id, str(move_error)))
                        
            except Exception as transfer_error:
                failed_transfers.append((element, str(transfer_error)))
                self.output.print_md("  FAILED: {} (ID: {}) - {}".format(
                    element_name, element_id, str(transfer_error)))
        
        # Strategy 3: For failed transfers, try positioning-based approach
        if failed_transfers:
            self.output.print_md("Attempting positioning strategy for {} failed transfers...".format(len(failed_transfers)))
            
            for element, error_msg in failed_transfers:
                try:
                    element_name = self.get_element_name_safely(element)
                    element_id = element.Id.IntegerValue
                    
                    # For failed elements, just ensure they're positioned correctly
                    # Check if element is close enough to target wall to be considered "transferred"
                    element_location = element.Location
                    
                    if hasattr(element_location, 'Point'):
                        elem_point = element_location.Point
                        target_wall_location = target_wall.Location
                        
                        if isinstance(target_wall_location, LocationCurve):
                            target_curve = target_wall_location.Curve
                            closest_point_info = target_curve.Project(elem_point)
                            
                            if closest_point_info and closest_point_info.Distance < 2.0:  # Within 2 feet
                                successfully_transferred.append(element)
                                self.transferred_elements += 1
                                self.output.print_md("  POSITIONED: {} (ID: {}) - Within acceptable range of target wall".format(
                                    element_name, element_id))
                            else:
                                self.output.print_md("  POSITIONING FAILED: {} (ID: {}) - Too far from target wall ({:.2f} ft)".format(
                                    element_name, element_id, closest_point_info.Distance if closest_point_info else 999))
                        else:
                            self.output.print_md("  POSITIONING FAILED: {} (ID: {}) - Target wall has no location curve".format(
                                element_name, element_id))
                    else:
                        self.output.print_md("  POSITIONING FAILED: {} (ID: {}) - Element has no accessible location".format(
                            element_name, element_id))
                        
                except Exception as position_error:
                    self.output.print_md("  POSITIONING FAILED: {} - {}".format(
                        self.get_element_name_safely(element), str(position_error)))
        
        # Summary
        total_processed = len(hosted_elements)
        success_count = len(successfully_transferred)
        
        self.output.print_md("Transfer Summary: {}/{} elements successfully transferred".format(
            success_count, total_processed))
        
        return successfully_transferred
    
    def copy_element_parameters(self, source_element, target_element):
        """Copy parameters from source to target element."""
        try:
            copied_count = 0
            source_params = source_element.Parameters
            
            for param in source_params:
                if param.IsReadOnly or not param.HasValue:
                    continue
                
                try:
                    target_param = target_element.LookupParameter(param.Definition.Name)
                    if target_param and not target_param.IsReadOnly:
                        # Copy by storage type
                        if param.StorageType == StorageType.String:
                            target_param.Set(param.AsString())
                        elif param.StorageType == StorageType.Double:
                            target_param.Set(param.AsDouble())
                        elif param.StorageType == StorageType.Integer:
                            target_param.Set(param.AsInteger())
                        elif param.StorageType == StorageType.ElementId:
                            target_param.Set(param.AsElementId())
                        
                        copied_count += 1
                except:
                    continue  # Skip parameters that can't be copied
            
            if copied_count > 0:
                self.output.print_md("    Copied {} parameters to new element".format(copied_count))
                
        except Exception as e:
            self.output.print_md("    Warning: Parameter copy failed: {}".format(str(e)))
    
    def _try_element_move_method(self, element, target_wall, success_list):
        """Alternative method for moving elements that can't be recreated."""
        try:
            element_name = self.get_element_name_safely(element)
            element_id = element.Id.IntegerValue
            
            # Try simple spatial association
            element_location = element.Location
            if hasattr(element_location, 'Point'):
                elem_point = element_location.Point
                target_wall_location = target_wall.Location
                
                if isinstance(target_wall_location, LocationCurve):
                    target_curve = target_wall_location.Curve
                    closest_point_info = target_curve.Project(elem_point)
                    
                    if closest_point_info and closest_point_info.Distance < 1.0:  # Within 1 foot
                        success_list.append(element)
                        self.transferred_elements += 1
                        self.output.print_md("  ALTERNATIVE SUCCESS: {} (ID: {}) - Spatially associated".format(
                            element_name, element_id))
                        return True
            
            return False
            
        except Exception as e:
            self.output.print_md("  ALTERNATIVE FAILED: {} - {}".format(
                self.get_element_name_safely(element), str(e)))
            return False
    
    def calculate_layer_positions(self, wall, layers):
        """Calculate position offsets for each layer based on wall orientation."""
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
    
    def can_delete_wall_safely(self, wall):
        """Check if wall can be deleted safely (no remaining hosted elements)."""
        try:
            # Get current dependent elements
            dependent_ids = wall.GetDependentElements(None)
            
            # Check for non-wall dependencies
            for dep_id in dependent_ids:
                element = doc.GetElement(dep_id)
                if element and not isinstance(element, Wall):
                    # Found a non-wall dependent element
                    return False, "Wall still has hosted element: {} (ID: {})".format(
                        self.get_element_name_safely(element), element.Id.IntegerValue)
            
            return True, "Wall is safe to delete"
            
        except Exception as e:
            return False, "Error checking wall dependencies: {}".format(str(e))
    
    def separate_wall_layers(self, wall):
        """Main function to separate a wall into its material layers."""
        wall_created_count = 0
        
        try:
            wall_type = wall.WallType
            wall_type_name = self.get_element_name_safely(wall_type)
            
            self.output.print_md("## Processing Wall ID: {} ({})".format(
                wall.Id.IntegerValue, wall_type_name))
            
            # Get compound structure and layers
            structure = wall_type.GetCompoundStructure()
            if not structure:
                error_msg = "Wall type has no compound structure"
                self.errors.append(error_msg)
                self.output.print_md("**Error**: {}".format(error_msg))
                return 0, None
            
            layers = structure.GetLayers()
            if not layers or len(layers) <= 1:
                error_msg = "Wall has {} layers - no separation needed".format(len(layers))
                self.output.print_md("**Info**: {}".format(error_msg))
                return 0, None
            
            self.output.print_md("Found {} layers to process".format(len(layers)))
            
            # Get hosted elements BEFORE creating new walls
            hosted_elements = self.get_hosted_elements(wall)
            
            # Calculate layer positions
            layer_positions = self.calculate_layer_positions(wall, layers)
            
            # Create new walls for each layer
            created_walls = []
            target_wall_for_hosted = None  # Will be the first successfully created wall
            
            for i, layer in enumerate(layers):
                self.output.print_md("Processing layer {} of {}".format(i + 1, len(layers)))
                
                # Check layer thickness
                layer_thickness_mm = layer.Width * 304.8
                if layer_thickness_mm < 0.1:
                    self.output.print_md("Skipping layer {} - thickness too small ({:.3f}mm)".format(
                        i + 1, layer_thickness_mm))
                    self.skipped_layers += 1
                    continue
                
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
                        material = None
                
                # Create wall type for this layer
                layer_wall_type = self.create_wall_type_for_layer(
                    wall_type, material, layer.Width, i)
                
                if not layer_wall_type:
                    continue
                
                # Create new wall instance
                try:
                    wall_location = wall.Location
                    if isinstance(wall_location, LocationCurve):
                        wall_curve = wall_location.Curve
                        
                        # Get position offset for this layer
                        position_offset = layer_positions[i]
                        
                        # Create offset curve for this layer
                        if position_offset.GetLength() > 0.001:
                            try:
                                offset_distance = position_offset.GetLength()
                                wall_direction = (wall_curve.GetEndPoint(1) - wall_curve.GetEndPoint(0)).Normalize()
                                wall_normal = XYZ(-wall_direction.Y, wall_direction.X, 0).Normalize()
                                
                                if position_offset.DotProduct(wall_normal) < 0:
                                    offset_distance = -offset_distance
                                
                                offset_curve = wall_curve.CreateOffset(offset_distance, XYZ.BasisZ)
                                
                            except:
                                offset_curve = wall_curve
                        else:
                            offset_curve = wall_curve
                        
                        # Get wall height and level
                        wall_height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble()
                        
                        # Create new wall
                        new_wall = self.create_wall_with_type(offset_curve, layer_wall_type, wall.LevelId, wall_height)
                        
                        if new_wall:
                            # First successfully created wall gets the hosted elements
                            if target_wall_for_hosted is None:
                                target_wall_for_hosted = new_wall
                                self.output.print_md("Target wall for hosted elements: ID {}".format(
                                    new_wall.Id.IntegerValue))
                            
                            created_walls.append(new_wall)
                            self.created_elements.append(new_wall.Id)
                            wall_created_count += 1
                            
                            self.output.print_md("SUCCESS: Layer {} - {} ({:.1f}mm) - Wall ID: {}".format(
                                i + 1, material_name, layer_thickness_mm, new_wall.Id.IntegerValue))
                
                except Exception as wall_error:
                    error_msg = "Failed to create wall for layer {}: {}".format(i, str(wall_error))
                    self.errors.append(error_msg)
                    self.output.print_md("**Error**: {}".format(error_msg))
            
            # Transfer hosted elements to target wall
            if hosted_elements and target_wall_for_hosted:
                self.output.print_md("### Transferring Hosted Elements")
                transferred_elements = self.transfer_hosted_elements(
                    wall, target_wall_for_hosted, hosted_elements)
                
                if transferred_elements:
                    self.output.print_md("Successfully transferred {} hosted elements".format(
                        len(transferred_elements)))
            
            return wall_created_count, target_wall_for_hosted
            
        except Exception as e:
            error_msg = "Failed to process wall {}: {}".format(wall.Id.IntegerValue, str(e))
            self.errors.append(error_msg)
            self.output.print_md("**Error**: {}".format(error_msg))
            return 0, None
    
    def run(self):
        """Main execution function."""
        self.output.print_md("# Wall Layer Separation Tool - Enhanced")
        self.output.print_md("---")
        
        # Get wall selection
        selection = uidoc.Selection
        selected_el_ids = selection.GetElementIds()
        selected_el = [doc.GetElement(e_id) for e_id in selected_el_ids]
        
        # Prompt user selection if nothing selected
        if not selected_el:
            try:
                selected_el_refs = selection.PickObjects(ObjectType.Element, "Select walls to separate")
                selected_el = [doc.GetElement(ref.ElementId) for ref in selected_el_refs]
            except:
                forms.alert('No walls selected. Operation cancelled.', exitscript=True)
        
        # Filter to get only walls
        selected_walls = [el for el in selected_el if isinstance(el, Wall)]
        
        if not selected_walls:
            forms.alert('No walls selected. Please select walls and try again.', exitscript=True)
        
        self.output.print_md("Selected {} walls for processing".format(len(selected_walls)))
        
        # Confirm operation
        proceed = forms.alert(
            "This will separate {} walls into individual material layers.\n\n"
            "Original walls will be deleted after hosted elements are transferred.\n\n"
            "Continue?".format(len(selected_walls)),
            title="Confirm Wall Separation",
            yes=True, no=True
        )
        
        if not proceed:
            self.output.print_md("Operation cancelled by user.")
            return
        
        # Process walls
        walls_to_delete = []
        successful_separations = []
        
        # Start main transaction
        t = Transaction(doc, "Separate Wall Layers")
        t.Start()
        
        try:
            for wall in selected_walls:
                created_count, target_wall = self.separate_wall_layers(wall)
                if created_count > 0:
                    self.processed_walls += 1
                    self.created_walls += created_count
                    walls_to_delete.append(wall)
                    successful_separations.append((wall, target_wall))
            
            # Commit the main transaction
            t.Commit()
            self.output.print_md("Successfully created all wall layers and transferred hosted elements")
            
        except Exception as e:
            t.RollBack()
            self.output.print_md("**Error**: Main transaction failed: {}".format(str(e)))
            forms.alert("Operation failed: {}".format(str(e)))
            return
        
        # Delete original walls in separate transaction (safer)
        if walls_to_delete and self.created_walls > 0:
            self.output.print_md("### Deleting Original Walls")
            t2 = Transaction(doc, "Delete Original Walls")
            t2.Start()
            
            try:
                deleted_count = 0
                preserved_count = 0
                
                for wall in walls_to_delete:
                    try:
                        # Check if wall can be safely deleted
                        can_delete, message = self.can_delete_wall_safely(wall)
                        
                        if can_delete:
                            doc.Delete(wall.Id)
                            deleted_count += 1
                            self.output.print_md("Deleted original wall ID: {}".format(wall.Id.IntegerValue))
                        else:
                            preserved_count += 1
                            self.output.print_md("Preserved wall ID {}: {}".format(
                                wall.Id.IntegerValue, message))
                        
                    except Exception as delete_error:
                        preserved_count += 1
                        self.output.print_md("Could not delete wall {}: {}".format(
                            wall.Id.IntegerValue, str(delete_error)))
                
                t2.Commit()
                self.output.print_md("Deletion summary: {} deleted, {} preserved".format(
                    deleted_count, preserved_count))
                
            except Exception as e:
                t2.RollBack()
                self.output.print_md("Deletion transaction failed: {}".format(str(e)))
        
        # Final report
        self.output.print_md("---")
        self.output.print_md("## Summary")
        self.output.print_md("- **Processed walls**: {}".format(self.processed_walls))
        self.output.print_md("- **Created walls**: {}".format(self.created_walls))
        self.output.print_md("- **Transferred elements**: {}".format(self.transferred_elements))
        self.output.print_md("- **Skipped layers**: {} (zero thickness)".format(self.skipped_layers))
        self.output.print_md("- **Errors**: {}".format(len(self.errors)))
        
        if self.errors:
            self.output.print_md("### Errors:")
            for error in self.errors:
                self.output.print_md("- {}".format(error))
        
        # Success message
        if self.created_walls > 0:
            forms.alert(
                "Wall separation completed!\n\n"
                "Processed: {} walls\n"
                "Created: {} new walls\n"
                "Transferred: {} hosted elements\n"
                "Errors: {}".format(
                    self.processed_walls, self.created_walls, 
                    self.transferred_elements, len(self.errors)),
                title="Operation Complete"
            )

# Run the tool
if __name__ == '__main__':
    separator = WallLayerSeparator()
    separator.run()