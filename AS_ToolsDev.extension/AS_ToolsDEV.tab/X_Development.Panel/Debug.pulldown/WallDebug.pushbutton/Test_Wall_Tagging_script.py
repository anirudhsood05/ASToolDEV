# -*- coding: utf-8 -*-
"""
TEST: Wall Tagging API
======================
Quick test to verify which tagging API works
"""

__title__ = "Test Wall\nTagging"
__author__ = "Anirudh Sood"

from pyrevit import revit, DB, forms, script

output = script.get_output()
doc = revit.doc

output.print_md("# Wall Tagging API Test")
output.print_md("---")

# Check if we have a selection
selection = revit.get_selection()
if not selection or len(selection) == 0:
    forms.alert("Please select at least one wall first", exitscript=True)

walls = [e for e in selection if e.Category.Name == "Walls"]
if len(walls) == 0:
    forms.alert("Please select at least one wall", exitscript=True)

wall = walls[0]
output.print_md("Testing with wall: {}".format(wall.Id))

# Get active view
view = doc.ActiveView
output.print_md("Active view: {}".format(view.Name))

# Get wall center
try:
    bbox = wall.get_BoundingBox(None)
    if bbox:
        center = (bbox.Min + bbox.Max) / 2.0
        tag_location = DB.XYZ(center.X + 0.5, center.Y + 0.5, center.Z)
        output.print_md("Tag location: ({}, {}, {})".format(tag_location.X, tag_location.Y, tag_location.Z))
    else:
        forms.alert("Cannot get wall center", exitscript=True)
except Exception as e:
    forms.alert("Error getting center: {}".format(str(e)), exitscript=True)

# Try tagging
output.print_md("")
output.print_md("## Attempting to create tag...")
output.print_md("")

with revit.Transaction("Test Tag"):
    try:
        # Method 1: doc.Create.NewTag
        output.print_md("**Method 1:** doc.Create.NewTag")
        tag = doc.Create.NewTag(
            view,
            wall,
            True,  # Has leader
            DB.TagMode.TM_ADDBY_CATEGORY,
            DB.TagOrientation.Horizontal,
            tag_location
        )
        
        if tag:
            output.print_md("✓ SUCCESS! Tag created with ID: {}".format(tag.Id))
            output.print_md("Tag is visible in current view.")
        else:
            output.print_md("✗ FAILED: NewTag returned None")
    
    except Exception as e:
        output.print_md("✗ FAILED: {}".format(str(e)))
        output.print_md("")
        
        # Try Method 2
        try:
            output.print_md("**Method 2:** IndependentTag.Create")
            tag = DB.IndependentTag.Create(
                doc,
                view.Id,
                DB.Reference(wall),
                True,
                DB.TagMode.TM_ADDBY_CATEGORY,
                DB.TagOrientation.Horizontal,
                tag_location
            )
            
            if tag:
                output.print_md("✓ SUCCESS! Tag created with ID: {}".format(tag.Id))
            else:
                output.print_md("✗ FAILED: Create returned None")
        
        except Exception as e2:
            output.print_md("✗ FAILED: {}".format(str(e2)))

output.print_md("")
output.print_md("---")
output.print_md("Test complete. Check your active view for the tag.")
