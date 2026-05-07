# -*- coding: utf-8 -*-
"""Setup Script for CAD Block to Family Placement Tool
Run this script to configure and test the tool installation."""

from pyrevit import script, forms, revit
import os
import json

class ToolSetup:
    def __init__(self):
        self.script_dir = script.get_script_path()
        self.config_file = os.path.join(self.script_dir, 'config.json')
        self.test_results = []
    
    def run_setup(self):
        """Run the complete setup process"""
        forms.alert("Starting CAD Block to Family Placement Tool Setup", 
                   title="Setup Wizard")
        
        # Step 1: Create default configuration
        if self.create_default_config():
            self.test_results.append("✓ Configuration file created")
        else:
            self.test_results.append("✗ Failed to create configuration")
        
        # Step 2: Test Revit API access
        if self.test_revit_api():
            self.test_results.append("✓ Revit API access confirmed")
        else:
            self.test_results.append("✗ Revit API access failed")
        
        # Step 3: Check for CAD links
        cad_count = self.check_cad_links()
        if cad_count >= 0:
            self.test_results.append("✓ Found {} CAD link(s) in current view".format(cad_count))
        else:
            self.test_results.append("✗ Error checking CAD links")
        
        # Step 4: Verify family access
        family_count = self.check_loaded_families()
        if family_count >= 0:
            self.test_results.append("✓ Found {} loaded family/families".format(family_count))
        else:
            self.test_results.append("✗ Error checking families")
        
        # Step 5: Test file permissions
        if self.test_file_permissions():
            self.test_results.append("✓ File permissions verified")
        else:
            self.test_results.append("✗ File permission issues detected")
        
        # Show results
        self.show_setup_results()
    
    def create_default_config(self):
        """Create default configuration file"""
        try:
            default_config = {
                "version": "1.0.0",
                "settings": {
                    "default_search_radius": 5.0,
                    "max_placement_distance": 10.0,
                    "prefer_horizontal_faces": True,
                    "auto_activate_family_types": True,
                    "enable_logging": True,
                    "log_level": "INFO"
                },
                "ui_preferences": {
                    "remember_last_selections": True,
                    "show_preview_by_default": False,
                    "confirm_batch_operations": True
                },
                "layer_mappings": {
                    "LIGHTING": {
                        "family": "Light Fixture",
                        "type": "Default",
                        "description": "Standard lighting fixtures"
                    },
                    "EQUIPMENT": {
                        "family": "Equipment",
                        "type": "Default", 
                        "description": "Mechanical/Electrical equipment"
                    }
                },
                "parameter_mappings": {
                    "CAD_LAYER": "Source_Layer",
                    "CAD_BLOCK_NAME": "Block_Name",
                    "CAD_ID": "CAD_Reference_ID",
                    "PLACEMENT_DATE": "Installation_Date"
                },
                "family_search_paths": [
                    "C:\\ProgramData\\Autodesk\\RVT 2023\\Families",
                    "C:\\Company\\Revit\\Families"
                ]
            }
            
            with open(self.config_file, 'w') as f:
                json.dump(default_config, f, indent=4)
            
            return True
            
        except Exception as e:
            script.get_logger().error("Error creating config: {}".format(str(e)))
            return False
    
    def test_revit_api(self):
        """Test basic Revit API functionality"""
        try:
            doc = revit.doc
            if not doc:
                return False
            
            # Test basic document operations
            app = doc.Application
            version = app.VersionNumber
            
            # Test element collection
            elements = revit.query.get_all_elements()
            
            return True
            
        except Exception as e:
            script.get_logger().error("Revit API test failed: {}".format(str(e)))
            return False
    
    def check_cad_links(self):
        """Check for CAD links in current view"""
        try:
            doc = revit.doc
            active_view = doc.ActiveView
            
            cad_links = revit.query.get_elements_by_class(
                revit.DB.ImportInstance, 
                view_id=active_view.Id
            )
            
            return len(cad_links)
            
        except Exception as e:
            script.get_logger().error("CAD link check failed: {}".format(str(e)))
            return -1
    
    def check_loaded_families(self):
        """Check for loaded families in the project"""
        try:
            families = revit.query.get_elements_by_class(revit.DB.Family)
            
            # Filter for face-hosted families
            hosted_families = []
            for family in families:
                if not family.IsInPlace:
                    symbols = family.GetFamilySymbolIds()
                    if symbols:
                        hosted_families.append(family)
            
            return len(hosted_families)
            
        except Exception as e:
            script.get_logger().error("Family check failed: {}".format(str(e)))
            return -1
    
    def test_file_permissions(self):
        """Test file read/write permissions"""
        try:
            # Test write permission
            test_file = os.path.join(self.script_dir, 'test_permissions.tmp')
            
            with open(test_file, 'w') as f:
                f.write("Permission test")
            
            # Test read permission
            with open(test_file, 'r') as f:
                content = f.read()
            
            # Clean up
            os.remove(test_file)
            
            return content == "Permission test"
            
        except Exception as e:
            script.get_logger().error("Permission test failed: {}".format(str(e)))
            return False
    
    def show_setup_results(self):
        """Display setup results to user"""
        result_text = "Setup Results:\n\n"
        result_text += "\n".join(self.test_results)
        
        # Determine overall status
        failed_tests = [r for r in self.test_results if r.startswith("✗")]
        
        if not failed_tests:
            result_text += "\n\n✓ Setup completed successfully!"
            result_text += "\nThe tool is ready to use."
        else:
            result_text += "\n\n⚠ Setup completed with issues."
            result_text += "\nSome features may not work correctly."
            result_text += "\nPlease address the failed tests above."
        
        # Add configuration location
        result_text += "\n\nConfiguration file location:"
        result_text += "\n{}".format(self.config_file)
        
        forms.alert(result_text, title="Setup Complete")
    
    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            return None
        except Exception as e:
            script.get_logger().error("Error loading config: {}".format(str(e)))
            return None
    
    def save_config(self, config):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            return True
        except Exception as e:
            script.get_logger().error("Error saving config: {}".format(str(e)))
            return False

# Configuration manager class for use in main tool
class ConfigManager:
    def __init__(self, config_path=None):
        if config_path is None:
            self.config_path = os.path.join(script.get_script_path(), 'config.json')
        else:
            self.config_path = config_path
        
        self.config = self.load_config()
    
    def load_config(self):
        """Load configuration with defaults"""
        default_config = {
            "settings": {
                "default_search_radius": 5.0,
                "max_placement_distance": 10.0,
                "prefer_horizontal_faces": True,
                "auto_activate_family_types": True
            },
            "parameter_mappings": {},
            "layer_mappings": {}
        }
        
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    loaded_config = json.load(f)
                    # Merge with defaults
                    default_config.update(loaded_config)
            return default_config
        except Exception as e:
            script.get_logger().error("Error loading config: {}".format(str(e)))
            return default_config
    
    def get_setting(self, key, default=None):
        """Get a setting value"""
        return self.config.get('settings', {}).get(key, default)
    
    def get_layer_mapping(self, layer_name):
        """Get family mapping for a layer"""
        return self.config.get('layer_mappings', {}).get(layer_name)
    
    def get_parameter_mappings(self):
        """Get parameter mappings"""
        return self.config.get('parameter_mappings', {})
    
    def update_setting(self, key, value):
        """Update a setting"""
        if 'settings' not in self.config:
            self.config['settings'] = {}
        self.config['settings'][key] = value
        self.save_config()
    
    def save_config(self):
        """Save current configuration"""
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.config, f, indent=4)
            return True
        except Exception as e:
            script.get_logger().error("Error saving config: {}".format(str(e)))
            return False

# Sample usage integration in main tool
class EnhancedCADBlockPlacementWindow(forms.WPFWindow):
    def __init__(self):
        # Load configuration
        self.config_manager = ConfigManager()
        
        # Apply configuration
        self.search_radius = self.config_manager.get_setting('default_search_radius', 5.0)
        self.auto_activate = self.config_manager.get_setting('auto_activate_family_types', True)
        
        # Continue with normal initialization
        super(EnhancedCADBlockPlacementWindow, self).__init__()
    
    def apply_configuration_settings(self):
        """Apply configuration settings to tool behavior"""
        # Example: Set UI preferences
        if self.config_manager.get_setting('remember_last_selections', True):
            self.restore_last_selections()
        
        # Example: Load parameter mappings
        self.parameter_mappings = self.config_manager.get_parameter_mappings()
        
        # Example: Load layer mappings for auto-selection
        self.layer_mappings = self.config_manager.config.get('layer_mappings', {})

# Run setup if executed directly
if __name__ == '__main__':
    setup = ToolSetup()
    setup.run_setup()