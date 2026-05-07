# -*- coding: utf-8 -*-
"""Convert Elevation View to Section View
Recreates an elevation view as a section view by calculating matching geometry.
"""

__title__ = 'Elevation\nto Section'
__author__ = 'Your Name'

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

def get_section_type():
    """Get first available section view type."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    for vt in collector:
        if vt.ViewFamily == DB.ViewFamily.Section:
            return vt
    return None

def get_unique_view_name(base_name):
    """Generate unique view name by appending number if needed."""
    existing_views = DB.FilteredElementCollector(doc).OfClass(DB.View)
    existing_names = set(v.Name for v in existing_views)
    
    # Try base name first
    if base_name not in existing_names:
        return base_name
    
    # Append numbers until unique
    counter = 1
    while True:
        new_name = "{} ({})".format(base_name, counter)
        if new_name not in existing_names:
            return new_name
        counter += 1

def create_section_from_elevation(elev_view):
    """Create section view matching elevation view geometry."""
    
    # Get elevation properties
    origin = elev_view.Origin
    view_dir = elev_view.ViewDirection
    up_dir = elev_view.UpDirection
    right_dir = elev_view.RightDirection
    
    # Get crop box dimensions
    crop_manager = elev_view.GetCropRegionShapeManager()
    crop_curves = crop_manager.GetCropShape()[0]
    
    # Calculate crop box extents in elevation's coordinate system
    min_x = min_y = float('inf')
    max_x = max_y = float('-inf')
    
    for curve in crop_curves:
        for i in range(2):
            pt = curve.GetEndPoint(i)
            # Transform to elevation view coordinates
            local_pt = pt - origin
            x = local_pt.DotProduct(right_dir)
            y = local_pt.DotProduct(up_dir)
            min_x = min(min_x, x)
            max_x = max(max_x, x)
            min_y = min(min_y, y)
            max_y = max(max_y, y)
    
    # Create section bounding box
    depth = 10.0  # Default depth (adjustable)
    
    transform = DB.Transform.Identity
    transform.Origin = origin
    transform.BasisX = right_dir
    transform.BasisY = up_dir
    transform.BasisZ = -view_dir  # Section looks opposite direction
    
    section_box = DB.BoundingBoxXYZ()
    section_box.Transform = transform
    section_box.Min = DB.XYZ(min_x, min_y, -1.0)  # Near clip
    section_box.Max = DB.XYZ(max_x, max_y, depth)  # Far clip
    
    return section_box

# Get active view
active_view = uidoc.ActiveView

# Validate view type
if active_view.ViewType != DB.ViewType.Elevation:
    forms.alert('Please run this tool in an Elevation view.', exitscript=True)

if active_view.IsTemplate:
    forms.alert('Cannot convert view templates.', exitscript=True)

# Get section view type
section_type = get_section_type()
if not section_type:
    forms.alert('No section view type found in project.', exitscript=True)

# Store the new section ID
new_section_id = None

# Create new section
with revit.Transaction('Convert Elevation to Section'):
    try:
        # Calculate section geometry
        section_box = create_section_from_elevation(active_view)
        
        # Create section view
        new_section = DB.ViewSection.CreateSection(doc, section_type.Id, section_box)
        new_section_id = new_section.Id
        
        # Generate unique name
        base_name = active_view.Name + ' (Section)'
        unique_name = get_unique_view_name(base_name)
        new_section.Name = unique_name
        
        # Copy basic properties
        new_section.Scale = active_view.Scale
        
        # Copy view template if applied
        if active_view.ViewTemplateId != DB.ElementId.InvalidElementId:
            new_section.ViewTemplateId = active_view.ViewTemplateId
        
        # Try to copy detail level (skip if template controls it)
        try:
            new_section.DetailLevel = active_view.DetailLevel
        except:
            pass
        
        # Copy display style
        try:
            new_section.DisplayStyle = active_view.DisplayStyle
        except:
            pass
        
        output.print_md('**Success!** Created section view: **{}**'.format(new_section.Name))
        output.print_md('Source elevation: {}'.format(active_view.Name))
        
    except Exception as e:
        forms.alert('Error creating section:\n{}'.format(str(e)))
        import sys
        sys.exit()

# Open new section view AFTER transaction closes
if new_section_id:
    new_section = doc.GetElement(new_section_id)
    uidoc.ActiveView = new_section
    output.print_md('\n*Adjust section depth via crop region if needed*')