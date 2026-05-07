# -*- coding: utf-8 -*-
"""
Copy Sheets Between Documents - WITH ELEVATION TO SECTION CONVERSION
=====================================================================
Converts elevation views to sections, copies to destination, places on sheets.

WORKFLOW:
1. Convert all elevations to sections in SOURCE document
2. Copy sections to DESTINATION (sections copy reliably)
3. Create sheets and place section viewports
4. Copy other view types as normal

REQUIREMENTS:
- Models must be linked by shared coordinates
- Source and destination documents must be open
"""

import sys
from pyrevit.framework import List
from pyrevit import forms, revit, DB, UI, script
from pyrevit.revit import query

logger = script.get_logger()
output = script.get_output()

VIEW_TOS_PARAM = DB.BuiltInParameter.VIEW_DESCRIPTION


# =============================================================================
# OPTION CLASSES
# =============================================================================

class Option(forms.TemplateListItem):
    def __init__(self, op_name, default_state=False):
        super(Option, self).__init__(op_name)
        self.state = default_state


class OptionSet:
    def __init__(self):
        self.op_convert_elevations = Option(
            'Convert Elevations to Sections (recommended for coordinate issues)', 
            True
        )
        self.op_copy_vports = Option('Copy Viewports', True)
        self.op_copy_schedules = Option('Copy Schedules', True)
        self.op_copy_titleblock = Option('Copy Sheet Titleblock', True)
        self.op_copy_revisions = Option('Copy and Set Sheet Revisions', False)
        self.op_copy_placeholders_as_sheets = Option('Copy Placeholders as Sheets', True)
        self.op_copy_guides = Option('Copy Guide Grids', True)
        self.op_update_exist_view_contents = Option('Update Existing View Contents')


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    """Handle copy and paste errors."""
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


# =============================================================================
# COORDINATE TRANSFORMATION
# =============================================================================

def copy_section_via_reference_lines(source_doc, source_section, dest_doc):
    """Copy section by storing origin point and dimensions, then recreating."""
    try:
        # Get section crop box
        crop_box = source_section.CropBox
        if not crop_box:
            logger.error('Could not get crop box')
            return None
        
        # Get section geometry
        transform = crop_box.Transform
        origin = transform.Origin
        basis_x = transform.BasisX  # Section width direction
        basis_y = transform.BasisY  # Section height direction  
        basis_z = transform.BasisZ  # Section view direction
        
        min_pt = crop_box.Min
        max_pt = crop_box.Max
        
        # Store section properties as detail lines that will copy correctly
        # Create 3 points: origin, origin+X, origin+Y to define orientation
        p1 = origin
        p2 = origin + (basis_x * 100)  # 100mm reference length
        p3 = origin + (basis_y * 100)
        
        # Create temporary model lines in SOURCE
        temp_line_ids = []
        with revit.Transaction('Create Reference Geometry', doc=source_doc):
            # Create on horizontal plane at origin elevation
            plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, origin)
            sketch_plane = DB.SketchPlane.Create(source_doc, plane)
            
            # Create two lines to preserve orientation
            line1 = DB.Line.CreateBound(p1, p2)  # X direction
            line2 = DB.Line.CreateBound(p1, p3)  # Y direction
            
            model_line1 = source_doc.Create.NewModelCurve(line1, sketch_plane)
            model_line2 = source_doc.Create.NewModelCurve(line2, sketch_plane)
            
            temp_line_ids.append(model_line1.Id)
            temp_line_ids.append(model_line2.Id)
        
        # Copy to destination
        cp_options = DB.CopyPasteOptions()
        cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())
        
        with revit.Transaction('Copy Reference Geometry', doc=dest_doc):
            copied_line_ids = DB.ElementTransformUtils.CopyElements(
                source_doc,
                List[DB.ElementId](temp_line_ids),
                dest_doc,
                None,
                cp_options
            )
        
        if not copied_line_ids or copied_line_ids.Count < 2:
            with revit.Transaction('Delete Temp Lines', doc=source_doc):
                for lid in temp_line_ids:
                    source_doc.Delete(lid)
            return None
        
        # Get copied geometry in destination
        line1_copied = dest_doc.GetElement(copied_line_ids[0])
        line2_copied = dest_doc.GetElement(copied_line_ids[1])
        
        curve1 = line1_copied.GeometryCurve
        curve2 = line2_copied.GeometryCurve
        
        # Extract orientation from copied lines
        new_origin = curve1.GetEndPoint(0)
        new_basis_x = (curve1.GetEndPoint(1) - new_origin).Normalize()
        new_basis_y = (curve2.GetEndPoint(1) - new_origin).Normalize()
        new_basis_z = new_basis_x.CrossProduct(new_basis_y).Normalize()
        
        # Get section type
        section_type = get_section_type(dest_doc)
        if not section_type:
            return None
        
        # Create section in destination
        new_section = None
        with revit.Transaction('Create Section', doc=dest_doc):
            # Create transform matching original
            sec_transform = DB.Transform.Identity
            sec_transform.Origin = new_origin
            sec_transform.BasisX = new_basis_x
            sec_transform.BasisY = new_basis_y
            sec_transform.BasisZ = new_basis_z
            
            # Use original dimensions
            new_box = DB.BoundingBoxXYZ()
            new_box.Transform = sec_transform
            new_box.Min = min_pt
            new_box.Max = max_pt
            
            # Create section
            new_section = DB.ViewSection.CreateSection(dest_doc, section_type.Id, new_box)
            new_section.Name = get_unique_view_name(dest_doc, source_section.Name)
            new_section.Scale = source_section.Scale
            
            try:
                new_section.DetailLevel = source_section.DetailLevel
            except:
                pass
            
            # Delete temp lines
            for lid in copied_line_ids:
                dest_doc.Delete(lid)
        
        # Delete source temp lines
        with revit.Transaction('Delete Temp Lines', doc=source_doc):
            for lid in temp_line_ids:
                source_doc.Delete(lid)
        
        logger.debug('Section recreated at: {}'.format(new_origin))
        return new_section
        
    except Exception as e:
        logger.error('Error copying section: {}'.format(str(e)))
        return None
    """Get Project Base Point position in Internal Origin coordinates."""
    try:
        # Find Project Base Point
        pbp = DB.FilteredElementCollector(doc)\
                .OfCategory(DB.BuiltInCategory.OST_ProjectBasePoint)\
                .FirstElement()
        
        if not pbp:
            logger.warning('No Project Base Point found')
            return None
        
        # Get position directly
        position = pbp.Position
        
        logger.debug('Project Base Point: E={:.2f}, N={:.2f}, Elev={:.2f}'.format(
            position.X, position.Y, position.Z
        ))
        
        return position
        
    except Exception as e:
        logger.error('Error getting Project Base Point: {}'.format(str(e)))
        return None


def get_coordinate_transform(source_doc, dest_doc):
    """Calculate transformation using Project Base Point offset."""
    try:
        # Get Project Base Point positions
        source_pbp = get_project_base_point_offset(source_doc)
        dest_pbp = get_project_base_point_offset(dest_doc)
        
        if not source_pbp or not dest_pbp:
            logger.warning('Could not get Project Base Points')
            return None
        
        # Calculate offset: how much to shift from source to dest
        # Dest_Internal = Source_Internal + (Source_PBP - Dest_PBP)
        translation_vector = source_pbp - dest_pbp
        
        # Check if offset is significant
        if translation_vector.GetLength() < 0.003:  # ~1mm
            logger.debug('Project Base Points aligned')
            return None
        
        # Create transform
        transform = DB.Transform.CreateTranslation(translation_vector)
        
        logger.debug('PBP offset: E={:.2f}, N={:.2f}'.format(
            translation_vector.X, translation_vector.Y
        ))
        
        output.print_md('*Project Base Point offset: E={:.1f}mm, N={:.1f}mm*'.format(
            translation_vector.X, translation_vector.Y
        ))
        
        return transform
        
    except Exception as e:
        logger.error('Error calculating transformation: {}'.format(str(e)))
        return None


# =============================================================================
# ELEVATION TO SECTION CONVERSION
# =============================================================================

def get_section_type(doc):
    """Get first available section view type."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    for vt in collector:
        if vt.ViewFamily == DB.ViewFamily.Section:
            return vt
    return None


def get_unique_view_name(doc, base_name):
    """Generate unique view name by appending number if needed."""
    existing_views = DB.FilteredElementCollector(doc).OfClass(DB.View)
    existing_names = set(v.Name for v in existing_views)
    
    if base_name not in existing_names:
        return base_name
    
    counter = 1
    while True:
        new_name = "{} ({})".format(base_name, counter)
        if new_name not in existing_names:
            return new_name
        counter += 1


def create_section_from_elevation(doc, elev_view):
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


def convert_elevation_to_section(doc, elev_view):
    """Convert elevation view to section, return new section view."""
    
    section_type = get_section_type(doc)
    if not section_type:
        logger.error('No section view type found')
        return None
    
    try:
        # Calculate section geometry
        section_box = create_section_from_elevation(doc, elev_view)
        
        # Create section view
        new_section = DB.ViewSection.CreateSection(doc, section_type.Id, section_box)
        
        # Generate unique name
        base_name = elev_view.Name + ' (Section)'
        unique_name = get_unique_view_name(doc, base_name)
        new_section.Name = unique_name
        
        # Copy basic properties
        new_section.Scale = elev_view.Scale
        
        # Copy view template if applied
        if elev_view.ViewTemplateId != DB.ElementId.InvalidElementId:
            new_section.ViewTemplateId = elev_view.ViewTemplateId
        
        # Try to copy detail level
        try:
            new_section.DetailLevel = elev_view.DetailLevel
        except:
            pass
        
        # Copy display style
        try:
            new_section.DisplayStyle = elev_view.DisplayStyle
        except:
            pass
        
        logger.debug('Created section from elevation: {}'.format(new_section.Name))
        return new_section
        
    except Exception as e:
        logger.error('Failed to convert elevation {}: {}'.format(
            elev_view.Name, str(e)
        ))
        return None


# =============================================================================
# SHEET COPY FUNCTIONS (MODIFIED FOR SECTION CONVERSION)
# =============================================================================

def get_default_type(source_doc, type_group):
    return source_doc.GetDefaultElementTypeId(type_group)


def find_matching_view(dest_doc, source_view):
    for v in DB.FilteredElementCollector(dest_doc).OfClass(DB.View):
        if v.ViewType == source_view.ViewType \
                and query.get_name(v) == query.get_name(source_view):
            if source_view.ViewType == DB.ViewType.DrawingSheet:
                if v.SheetNumber == source_view.SheetNumber:
                    return v
            else:
                return v


def find_guide(guide_name, source_doc):
    guide_elements = \
        DB.FilteredElementCollector(source_doc)\
            .OfCategory(DB.BuiltInCategory.OST_GuideGrid)\
            .WhereElementIsNotElementType()\
            .ToElements()
    
    for guide in guide_elements:
        if str(guide.Name).lower() == guide_name.lower():
            return guide


def can_copy_view_contents(view):
    """Check if a view type supports content copying."""
    
    unsupported_types = [
        DB.ViewType.Schedule,
        DB.ViewType.CostReport,
        DB.ViewType.LoadsReport,
        DB.ViewType.PresureLossReport,
        DB.ViewType.ColumnSchedule,
        DB.ViewType.PanelSchedule,
        DB.ViewType.Walkthrough,
        DB.ViewType.Rendering,
        DB.ViewType.SystemsAnalysisReport,
        DB.ViewType.ProjectBrowser,
        DB.ViewType.Internal,
        DB.ViewType.Undefined,
        DB.ViewType.ThreeD
    ]
    
    if view.ViewType in unsupported_types:
        return False
    
    if view.ViewType in [DB.ViewType.Elevation, DB.ViewType.Section]:
        return False
    
    try:
        if view.IsTemplate:
            return False
        
        primary_view_id = view.GetPrimaryViewId()
        if primary_view_id != DB.ElementId.InvalidElementId:
            return False
            
    except Exception:
        return False
    
    return True


def get_view_contents(dest_doc, source_view):
    view_elements = DB.FilteredElementCollector(dest_doc, source_view.Id)\
                      .WhereElementIsNotElementType()\
                      .ToElements()

    elements_ids = []
    for element in view_elements:
        if (element.Category and element.Category.Name == 'Title Blocks') \
                and not OPTION_SET.op_copy_titleblock:
            continue
        elif isinstance(element, DB.ScheduleSheetInstance) \
                and not OPTION_SET.op_copy_schedules:
            continue
        elif isinstance(element, DB.Viewport) \
                or 'ExtentElem' in query.get_name(element):
            continue
        elif isinstance(element, DB.Element) \
                and element.Category \
                and 'guide' in str(element.Category.Name).lower():
            continue
        elif isinstance(element, DB.Element) \
                and element.Category \
                and 'views' == str(element.Category.Name).lower():
            continue
        else:
            elements_ids.append(element.Id)
    return elements_ids


def ensure_dest_revision(src_rev, all_dest_revs, dest_doc):
    for rev in all_dest_revs:
        if query.compare_revisions(rev, src_rev):
            return rev

    logger.warning('Revision could not be found in destination model.')
    return revit.create.create_revision(
        description=src_rev.Description,
        by=src_rev.IssuedBy,
        to=src_rev.IssuedTo,
        date=src_rev.RevisionDate,
        doc=dest_doc
    )


def clear_view_contents(dest_doc, dest_view):
    logger.debug('Removing view contents: {}'.format(dest_view.Name))
    elements_ids = get_view_contents(dest_doc, dest_view)

    with revit.Transaction('Delete View Contents', doc=dest_doc):
        for el_id in elements_ids:
            try:
                dest_doc.Delete(el_id)
            except Exception:
                continue

    return True


def copy_view_contents(activedoc, source_view, dest_doc, dest_view,
                       clear_contents=False):
    """Copy view contents with robust error handling."""
    
    logger.debug('Copying view contents: {}'.format(source_view.Name))

    if not can_copy_view_contents(source_view):
        logger.debug('View type does not support content copying: {}'.format(
            source_view.ViewType
        ))
        return True
    
    if not can_copy_view_contents(dest_view):
        return True

    elements_ids = get_view_contents(activedoc, source_view)

    if clear_contents:
        if not clear_view_contents(dest_doc, dest_view):
            return False

    if not elements_ids:
        return True

    cp_options = DB.CopyPasteOptions()
    cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

    try:
        with revit.Transaction('Copy View Contents',
                               doc=dest_doc,
                               swallow_errors=True):
            DB.ElementTransformUtils.CopyElements(
                source_view,
                List[DB.ElementId](elements_ids),
                dest_view, None, cp_options
            )
        return True
        
    except Exception as copy_err:
        logger.error('Failed to copy view contents: {}'.format(str(copy_err)))
        return True


def copy_view_props(source_view, dest_view):
    """Copy view properties with error handling."""
    try:
        dest_view.Scale = source_view.Scale
    except Exception:
        pass
    
    try:
        dest_view.Parameter[VIEW_TOS_PARAM].Set(
            source_view.Parameter[VIEW_TOS_PARAM].AsString()
        )
    except Exception:
        pass


def copy_view(activedoc, source_view, dest_doc):
    """Copy view to destination document."""
    
    matching_view = find_matching_view(dest_doc, source_view)
    if matching_view:
        print('\t\t\tView/Sheet already exists in document.')
        if OPTION_SET.op_update_exist_view_contents:
            copy_view_contents(activedoc, source_view, dest_doc, matching_view,
                             clear_contents=True)
        return matching_view

    logger.debug('Copying view: {}'.format(source_view.Name))
    new_view = None

    if source_view.ViewType == DB.ViewType.DrawingSheet:
        try:
            with revit.Transaction('Create Sheet', doc=dest_doc):
                if not source_view.IsPlaceholder \
                        or (source_view.IsPlaceholder
                                and OPTION_SET.op_copy_placeholders_as_sheets):
                    new_view = DB.ViewSheet.Create(
                        dest_doc,
                        DB.ElementId.InvalidElementId
                    )
                else:
                    new_view = DB.ViewSheet.CreatePlaceholder(dest_doc)

                revit.update.set_name(new_view, query.get_name(source_view))
                new_view.SheetNumber = source_view.SheetNumber
        except Exception as sheet_err:
            logger.error('Error creating sheet: {}'.format(sheet_err))
            
    elif source_view.ViewType == DB.ViewType.DraftingView:
        try:
            with revit.Transaction('Create Drafting View', doc=dest_doc):
                new_view = DB.ViewDrafting.Create(
                    dest_doc,
                    get_default_type(dest_doc, DB.ElementTypeGroup.ViewTypeDrafting)
                )
                revit.update.set_name(new_view, query.get_name(source_view))
                copy_view_props(source_view, new_view)
        except Exception as sheet_err:
            logger.error('Error creating drafting view: {}'.format(sheet_err))
            
    elif source_view.ViewType == DB.ViewType.Legend:
        try:
            first_legend = query.find_first_legend(dest_doc)
            if first_legend:
                with revit.Transaction('Create Legend View', doc=dest_doc):
                    new_view = dest_doc.GetElement(
                        first_legend.Duplicate(DB.ViewDuplicateOption.Duplicate)
                    )
                    revit.update.set_name(new_view, query.get_name(source_view))
                    copy_view_props(source_view, new_view)
            else:
                logger.error('Destination must have at least one Legend view')
        except Exception as sheet_err:
            logger.error('Error creating legend view: {}'.format(sheet_err))
    
    elif source_view.ViewType in [DB.ViewType.Section, DB.ViewType.Detail]:
        try:
            with revit.Transaction('Copy Section/Detail View', doc=dest_doc):
                cp_options = DB.CopyPasteOptions()
                cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())
                
                copied_ids = DB.ElementTransformUtils.CopyElements(
                    activedoc,
                    List[DB.ElementId]([source_view.Id]),
                    dest_doc,
                    None,  # No transform - may need manual repositioning
                    cp_options
                )
                
                if copied_ids and copied_ids.Count > 0:
                    new_view = dest_doc.GetElement(copied_ids[0])
                    print('\t\t\tSection copied (may need manual repositioning).')
                else:
                    logger.error('Copy operation returned no elements')
                    
        except Exception as section_err:
            logger.error('Error copying section/detail: {}'.format(section_err))
    
    else:
        try:
            with revit.Transaction('Copy View', doc=dest_doc):
                cp_options = DB.CopyPasteOptions()
                cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())
                
                copied_ids = DB.ElementTransformUtils.CopyElements(
                    activedoc,
                    List[DB.ElementId]([source_view.Id]),
                    dest_doc,
                    None,
                    cp_options
                )
                
                if copied_ids and copied_ids.Count > 0:
                    new_view = dest_doc.GetElement(copied_ids[0])
                else:
                    logger.error('Copy operation returned no elements')
                    
        except Exception as view_err:
            logger.error('Error copying view: {}'.format(view_err))

    if new_view:
        copy_view_contents(activedoc, source_view, dest_doc, new_view)

    return new_view


def copy_viewport_types(activedoc, vport_type, vport_typename,
                        dest_doc, newvport):
    dest_vport_typenames = [
        query.get_name(dest_doc.GetElement(x))
        for x in newvport.GetValidTypes()
    ]

    cp_options = DB.CopyPasteOptions()
    cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

    if vport_typename not in dest_vport_typenames:
        with revit.Transaction('Copy Viewport Types',
                               doc=dest_doc,
                               swallow_errors=True):
            DB.ElementTransformUtils.CopyElements(
                activedoc,
                List[DB.ElementId]([vport_type.Id]),
                dest_doc,
                None,
                cp_options,
            )


def apply_viewport_type(activedoc, vport_id, dest_doc, newvport_id):
    try:
        with revit.Transaction('Apply Viewport Type', doc=dest_doc):
            vport = activedoc.GetElement(vport_id)
            vport_type = activedoc.GetElement(vport.GetTypeId())
            vport_typename = query.get_name(vport_type)

            newvport = dest_doc.GetElement(newvport_id)

            copy_viewport_types(activedoc, vport_type, vport_typename,
                                dest_doc, newvport)

            for vtype_id in newvport.GetValidTypes():
                vtype = dest_doc.GetElement(vtype_id)
                if query.get_name(vtype) == vport_typename:
                    newvport.ChangeTypeId(vtype_id)
                    break
    except Exception as err:
        logger.debug('Could not apply viewport type: {}'.format(str(err)))


def copy_sheet_viewports(activedoc, source_sheet, dest_doc, dest_sheet, conversion_map):
    """Copy viewports - uses conversion map for elevations converted to sections."""
    
    existing_views = [dest_doc.GetElement(x).ViewId
                      for x in dest_sheet.GetAllViewports()]

    for vport_id in source_sheet.GetAllViewports():
        try:
            vport = activedoc.GetElement(vport_id)
            vport_view = activedoc.GetElement(vport.ViewId)

            print('\t\tProcessing view: {}'.format(query.get_name(vport_view)))
            
            # Check if this elevation was converted to section
            if vport_view.Id in conversion_map:
                print('\t\t\t(Using converted section instead of elevation)')
                source_view_to_copy = conversion_map[vport_view.Id]['section_view']
            else:
                source_view_to_copy = vport_view
            
            try:
                new_view = copy_view(activedoc, source_view_to_copy, dest_doc)

                if new_view:
                    if new_view.Id in existing_views:
                        print('\t\t\tView already on this sheet.')
                        continue
                    
                    # Check if on another sheet
                    ref_info = revit.query.get_view_sheetrefinfo(new_view)
                    if ref_info and ref_info.sheet_num:
                        print('\t\t\tRemoving from temporary sheet...')
                        all_sheets = DB.FilteredElementCollector(dest_doc)\
                                       .OfClass(DB.ViewSheet)\
                                       .ToElements()
                        
                        for temp_sheet in all_sheets:
                            if temp_sheet.SheetNumber == ref_info.sheet_num:
                                for vp_id in temp_sheet.GetAllViewports():
                                    vp = dest_doc.GetElement(vp_id)
                                    if vp and vp.ViewId == new_view.Id:
                                        with revit.Transaction('Remove Viewport',
                                                             doc=dest_doc):
                                            dest_doc.Delete(vp_id)
                                        break
                                break
                    
                    print('\t\t\tPlacing view on sheet.')
                    with revit.Transaction('Place View on Sheet', doc=dest_doc):
                        nvport = DB.Viewport.Create(
                            dest_doc,
                            dest_sheet.Id,
                            new_view.Id,
                            vport.GetBoxCenter()
                        )
                    
                    if nvport:
                        apply_viewport_type(activedoc, vport_id, dest_doc, nvport.Id)
                else:
                    print('\t\t\tWARNING: Could not copy view.')
                    
            except Exception as vport_err:
                logger.error('Failed to process viewport: {}'.format(str(vport_err)))
                print('\t\t\tWARNING: Skipping problematic viewport.')
                continue
                
        except Exception as outer_err:
            logger.error('Unexpected error: {}'.format(str(outer_err)))
            continue


def copy_sheet_revisions(activedoc, source_sheet, dest_doc, dest_sheet):
    all_src_revs = query.get_revisions(doc=activedoc)
    all_dest_revs = query.get_revisions(doc=dest_doc)
    revisions_to_set = []

    with revit.Transaction('Copy and Set Revisions', doc=dest_doc):
        for src_revid in source_sheet.GetAdditionalRevisionIds():
            set_rev = ensure_dest_revision(
                activedoc.GetElement(src_revid),
                all_dest_revs,
                dest_doc
            )
            revisions_to_set.append(set_rev)

        if revisions_to_set:
            revit.update.update_sheet_revisions(
                revisions_to_set,
                [dest_sheet],
                state=True,
                doc=dest_doc
            )


def copy_sheet_guides(activedoc, source_sheet, dest_doc, dest_sheet):
    source_sheet_guide_param = \
        source_sheet.Parameter[DB.BuiltInParameter.SHEET_GUIDE_GRID]
    source_sheet_guide_element = \
        activedoc.GetElement(source_sheet_guide_param.AsElementId())
    
    if source_sheet_guide_element:
        if not find_guide(source_sheet_guide_element.Name, dest_doc):
            cp_options = DB.CopyPasteOptions()
            cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

            with revit.Transaction('Copy Sheet Guide', doc=dest_doc):
                DB.ElementTransformUtils.CopyElements(
                    activedoc,
                    List[DB.ElementId]([source_sheet_guide_element.Id]),
                    dest_doc, None, cp_options
                )

        dest_guide = find_guide(source_sheet_guide_element.Name, dest_doc)
        if dest_guide:
            with revit.Transaction('Set Sheet Guide', doc=dest_doc):
                dest_sheet_guide_param = \
                    dest_sheet.Parameter[DB.BuiltInParameter.SHEET_GUIDE_GRID]
                dest_sheet_guide_param.Set(dest_guide.Id)


def copy_sheet(activedoc, source_sheet, dest_doc, conversion_map):
    """Copy sheet with converted sections."""
    
    logger.debug('Copying sheet {} to document {}'.format(
        source_sheet.Name, dest_doc.Title
    ))
    print('\tCopying Sheet: {}'.format(source_sheet.Name))
    
    sheet_success = False
    
    try:
        with revit.TransactionGroup('Import Sheet', doc=dest_doc):
            new_sheet = copy_view(activedoc, source_sheet, dest_doc)

            if new_sheet:
                if not new_sheet.IsPlaceholder:
                    if OPTION_SET.op_copy_vports:
                        copy_sheet_viewports(
                            activedoc,
                            source_sheet,
                            dest_doc,
                            new_sheet,
                            conversion_map
                        )
                    else:
                        print('Skipping viewports...')

                    if OPTION_SET.op_copy_guides:
                        try:
                            copy_sheet_guides(activedoc, source_sheet,
                                            dest_doc, new_sheet)
                        except Exception:
                            pass
                    else:
                        print('Skipping sheet guides...')

                if OPTION_SET.op_copy_revisions:
                    try:
                        copy_sheet_revisions(activedoc, source_sheet,
                                           dest_doc, new_sheet)
                    except Exception:
                        pass
                else:
                    print('Skipping revisions...')
                
                sheet_success = True
            else:
                logger.error('Failed copying sheet: {}'.format(source_sheet.Name))
                
    except Exception as sheet_err:
        logger.error('Critical error copying sheet: {}'.format(str(sheet_err)))
    
    return sheet_success


# =============================================================================
# UI FUNCTIONS
# =============================================================================

def get_user_options():
    op_set = OptionSet()
    return_options = forms.SelectFromList.show(
        [getattr(op_set, x) for x in dir(op_set) if x.startswith('op_')],
        title='Select Copy Options',
        button_name='Copy Now',
        multiselect=True
    )

    if not return_options:
        sys.exit(0)

    return op_set


def get_dest_docs():
    selected_dest_docs = forms.select_open_docs(
        title='Select Destination Documents',
        filterfunc=lambda d: not d.IsFamilyDocument
    )
    if not selected_dest_docs:
        sys.exit(0)
    return selected_dest_docs


def get_source_sheets():
    sheet_elements = forms.select_sheets(
        button_name='Copy Sheets',
        use_selection=True
    )

    if not sheet_elements:
        sys.exit(0)

    return sheet_elements


# =============================================================================
# MAIN EXECUTION
# =============================================================================

output.print_md('---')
output.print_md('# SHEET COPY WITH ELEVATION-TO-SECTION CONVERSION')
output.print_md('---')

# Get user selections
dest_docs = get_dest_docs()
doc_count = len(dest_docs)

source_sheets = get_source_sheets()
sheet_count = len(source_sheets)

OPTION_SET = get_user_options()

# Statistics
total_work = doc_count * sheet_count
work_counter = 0
sheets_succeeded = 0
sheets_failed = 0
elevations_converted = 0
sections_copied = 0

# =============================================================================
# STEP 1: CONVERT ELEVATIONS TO SECTIONS IN SOURCE DOCUMENT
# =============================================================================

conversion_map = {}  # Maps elevation view ID to conversion info

if OPTION_SET.op_convert_elevations:
    output.print_md('---')
    output.print_md('**STEP 1: CONVERTING ELEVATIONS TO SECTIONS**')
    output.print_md('*This preserves geometry but annotations are not transferred*')
    output.print_md('')
    
    # Collect all elevation views from selected sheets
    elevation_views = []
    for sheet in source_sheets:
        for vport_id in sheet.GetAllViewports():
            vport = revit.doc.GetElement(vport_id)
            view = revit.doc.GetElement(vport.ViewId)
            
            if view.ViewType == DB.ViewType.Elevation:
                if view.Id not in [ev.Id for ev in elevation_views]:
                    elevation_views.append(view)
    
    if elevation_views:
        output.print_md('Found {} elevation view(s) to convert'.format(
            len(elevation_views)
        ))
        
        with revit.Transaction('Convert Elevations to Sections', doc=revit.doc):
            for elev_view in elevation_views:
                print('Converting: {}'.format(elev_view.Name))
                
                section_view = convert_elevation_to_section(revit.doc, elev_view)
                
                if section_view:
                    conversion_map[elev_view.Id] = {
                        'original_elevation': elev_view,
                        'section_view': section_view,
                        'section_name': section_view.Name
                    }
                    elevations_converted += 1
                    print('\t-> Created: {}'.format(section_view.Name))
                else:
                    print('\tWARNING: Conversion failed')
        
        output.print_md('---')
        output.print_md('**Conversion Summary:**')
        output.print_md('- Elevations converted: {}'.format(elevations_converted))
        output.print_md('- Sections created: {}'.format(elevations_converted))
    else:
        output.print_md('No elevation views found on selected sheets.')

# =============================================================================
# STEP 2: COPY SHEETS AND VIEWPORTS TO DESTINATION
# =============================================================================

output.print_md('---')
output.print_md('**STEP 2: COPYING SHEETS TO DESTINATION DOCUMENTS**')
output.print_md('*Note: Sections will copy but may need manual repositioning*')
output.print_md('')

for dest_doc in dest_docs:
    output.print_md('**Target Document:** {}'.format(dest_doc.Title))

    for source_sheet in source_sheets:
        print('Copying Sheet: {} - {}'.format(
            source_sheet.SheetNumber,
            source_sheet.Name
        ))
        
        if copy_sheet(revit.doc, source_sheet, dest_doc, conversion_map):
            sheets_succeeded += 1
        else:
            sheets_failed += 1
            
        work_counter += 1
        output.update_progress(work_counter, total_work)

# =============================================================================
# FINAL SUMMARY
# =============================================================================

output.print_md('---')
output.print_md('# COPY COMPLETE')
output.print_md('---')

output.print_md('## Conversion Results')
if OPTION_SET.op_convert_elevations:
    output.print_md('- **Elevations converted to sections:** {}'.format(
        elevations_converted
    ))
    output.print_md('  - *Geometry preserved, annotations not transferred*')
else:
    output.print_md('- Elevation conversion was disabled')

output.print_md('')
output.print_md('## Sheet Copy Results')
output.print_md('- **Successfully copied:** {} sheets'.format(sheets_succeeded))
if sheets_failed > 0:
    output.print_md('- **Failed to copy:** {} sheets (see warnings above)'.format(
        sheets_failed
    ))
output.print_md('- **Target documents:** {}'.format(doc_count))

output.print_md('')
output.print_md('## Notes')
output.print_md('- **Sections may need manual repositioning** in destination model')
output.print_md('- Use aligned views from source as reference for correct positioning')
output.print_md('- Section geometry and properties are preserved')
output.print_md('- **Annotations are not transferred** during elevation→section conversion')
output.print_md('- Manually copy annotations if needed (Ctrl+C/Ctrl+V)')
