# -*- coding: utf-8 -*-
"""
DIAGNOSTIC: Show Door Type Names
=================================
"""

__title__ = "Show Door\nType Names"
__author__ = "Anirudh Sood"

from pyrevit import revit, DB, script
from pyrevit.revit.db import query
from pychilizer import database

output = script.get_output()
doc = revit.doc

output.print_md("# Door Type Names in Project")
output.print_md("---")

# Get all doors
doors = DB.FilteredElementCollector(doc)\
    .OfCategory(DB.BuiltInCategory.OST_Doors)\
    .WhereElementIsNotElementType()\
    .ToElements()

output.print_md("Found {} doors in project".format(len(doors)))
output.print_md("")

# Get unique type names
type_names = set()
for door in doors:
    try:
        door_type = query.get_type(door)
        if door_type:
            type_name = database.get_name(door_type)
            type_names.add(type_name)
    except:
        pass

output.print_md("## Unique Door Type Names ({} types):".format(len(type_names)))
output.print_md("")

# Sort and display
for type_name in sorted(type_names):
    # Check if matches filter
    matches_auk_door = "AUK_Door_" in type_name
    matches_auk = "AUK_Door" in type_name
    
    marker = ""
    if matches_auk_door:
        marker = " ✓ MATCHES 'AUK_Door_'"
    elif matches_auk:
        marker = " ⚠ MATCHES 'AUK_Door' (without _)"
    else:
        marker = " ✗ NO MATCH"
    
    output.print_md("- `{}`{}".format(type_name, marker))

output.print_md("")
output.print_md("---")
output.print_md("**Test Pattern Matching:**")
output.print_md("")

# Test a few examples
test_names = [
    "AUK_Door_Int-Sgl-Timber-01",
    "AUK_Door-Int-Sgl-Timber-01",  # hyphen instead of underscore?
    "AUK Door Int-Sgl-Timber-01",   # space instead?
]

for test in test_names:
    if test in type_names or True:  # Show anyway
        result = "AUK_Door_" in test
        output.print_md("- `'AUK_Door_' in '{}'` = **{}**".format(test, result))
