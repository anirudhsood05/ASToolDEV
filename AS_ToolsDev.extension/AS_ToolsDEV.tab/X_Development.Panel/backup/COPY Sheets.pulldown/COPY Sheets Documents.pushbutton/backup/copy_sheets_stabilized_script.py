import sys

from pyrevit.framework import List
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.revit import query
from pyrevit import script
from Autodesk.Revit.DB import Element as DBElement


logger = script.get_logger()
output = script.get_output()

selection = revit.get_selection()


VIEW_TOS_PARAM = DB.BuiltInParameter.VIEW_DESCRIPTION


class Option(forms.TemplateListItem):
    def __init__(self, op_name, default_state=False):
        super(Option, self).__init__(op_name)
        self.state = default_state


class OptionSet:
    def __init__(self):
        self.op_copy_elevation_markers = Option('Copy Elevation Markers First '
                                                '(creates elevation views)', True)
        self.op_copy_vports = Option('Copy Viewports', True)
        self.op_copy_schedules = Option('Copy Schedules', True)
        self.op_copy_titleblock = Option('Copy Sheet Titleblock', True)
        self.op_copy_revisions = Option('Copy and Set Sheet Revisions', False)
        self.op_copy_placeholders_as_sheets = \
            Option('Copy Placeholders as Sheets', True)
        self.op_copy_guides = Option('Copy Guide Grids', True)
        self.op_update_exist_view_contents = \
            Option('Update Existing View Contents')


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    """Handle copy and paste errors."""

    def OnDuplicateTypeNamesFound(self, args):
        """Use destination model types if duplicate."""
        return DB.DuplicateTypeAction.UseDestinationTypes


def get_user_options():
    op_set = OptionSet()
    return_options = \
        forms.SelectFromList.show(
            [getattr(op_set, x) for x in dir(op_set) if x.startswith('op_')],
            title='Select Copy Options',
            button_name='Copy Now',
            multiselect=True
            )

    if not return_options:
        sys.exit(0)

    return op_set


def get_dest_docs():
    # find open documents other than the active doc
    selected_dest_docs = \
        forms.select_open_docs(title='Select Destination Documents',
                               filterfunc=lambda d: not d.IsFamilyDocument)
    if not selected_dest_docs:
        sys.exit(0)
    else:
        return selected_dest_docs


def get_source_sheets():
    sheet_elements = forms.select_sheets(button_name='Copy Sheets',
                                         use_selection=True)

    if not sheet_elements:
        sys.exit(0)

    return sheet_elements


def collect_elevation_markers_from_plans(doc):
    """Collect all elevation markers from plan views in the document."""
    elevation_markers = []
    
    # Get all plan views
    plan_views = DB.FilteredElementCollector(doc)\
                   .OfClass(DB.ViewPlan)\
                   .ToElements()
    
    for plan_view in plan_views:
        # Skip templates
        if plan_view.IsTemplate:
            continue
            
        # Collect elevation markers in this plan view
        markers_in_view = DB.FilteredElementCollector(doc, plan_view.Id)\
                            .OfCategory(DB.BuiltInCategory.OST_Elev)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
        
        for marker in markers_in_view:
            if marker.Id not in [m.Id for m in elevation_markers]:
                elevation_markers.append(marker)
    
    return elevation_markers

# =============================================================================
# COORDINATE TRANSFORMATION FIX FOR ELEVATION MARKERS
# =============================================================================
# Add these functions to your script BEFORE the copy_elevation_markers function

def get_coordinate_transform_between_documents(source_doc, dest_doc):
    """Calculate transformation to convert coordinates from source to destination document.
    
    This handles projects with different internal origins or project base points.
    Uses the shared coordinate system as the common reference.
    
    Returns:
        Transform or None - Transform to apply, or None if documents at same location
    """
    try:
        # Get source document's transform (Internal -> Shared)
        source_location = source_doc.ActiveProjectLocation
        source_transform = source_location.GetTotalTransform()
        
        # Get destination document's transform (Shared -> Internal)  
        dest_location = dest_doc.ActiveProjectLocation
        dest_transform = dest_location.GetTotalTransform()
        dest_inverse = dest_transform.Inverse
        
        # Check if transformation is actually needed
        # If both transforms are essentially identity, no transformation required
        source_is_identity = is_transform_identity(source_transform)
        dest_is_identity = is_transform_identity(dest_transform)
        
        if source_is_identity and dest_is_identity:
            logger.debug('Both documents at same coordinates - no transform needed')
            return None
        
        # Combined transform: Source Internal -> Shared -> Dest Internal
        # This converts from source internal coordinates to destination internal coordinates
        combined_transform = dest_inverse.Multiply(source_transform)
        
        logger.debug('Coordinate transformation required:')
        logger.debug('  Source origin: ({:.2f}, {:.2f}, {:.2f})'.format(
            source_transform.Origin.X, 
            source_transform.Origin.Y, 
            source_transform.Origin.Z
        ))
        logger.debug('  Dest origin: ({:.2f}, {:.2f}, {:.2f})'.format(
            dest_transform.Origin.X,
            dest_transform.Origin.Y,
            dest_transform.Origin.Z
        ))
        
        return combined_transform
        
    except Exception as transform_err:
        logger.error('Could not calculate coordinate transform: {}'.format(str(transform_err)))
        logger.debug('Will copy markers without transformation - they may not align correctly')
        return None


def is_transform_identity(transform):
    """Check if transform is essentially an identity transform (no translation/rotation).
    
    Args:
        transform - Revit Transform object
        
    Returns:
        bool - True if transform is identity (no real transformation)
    """
    try:
        tolerance = 0.001  # ~0.3mm tolerance
        
        # Check translation - origin should be at (0,0,0)
        origin = transform.Origin
        has_translation = (abs(origin.X) > tolerance or 
                          abs(origin.Y) > tolerance or 
                          abs(origin.Z) > tolerance)
        
        if has_translation:
            return False
        
        # Check rotation - basis vectors should align with world axes
        basis_x = transform.BasisX
        basis_y = transform.BasisY
        basis_z = transform.BasisZ
        
        # BasisX should be (1, 0, 0)
        if not (abs(basis_x.X - 1.0) < tolerance and 
                abs(basis_x.Y) < tolerance and 
                abs(basis_x.Z) < tolerance):
            return False
            
        # BasisY should be (0, 1, 0)
        if not (abs(basis_y.Y - 1.0) < tolerance and 
                abs(basis_y.X) < tolerance and 
                abs(basis_y.Z) < tolerance):
            return False
            
        # BasisZ should be (0, 0, 1)
        if not (abs(basis_z.Z - 1.0) < tolerance and 
                abs(basis_z.X) < tolerance and 
                abs(basis_z.Y) < tolerance):
            return False
        
        return True
        
    except Exception as check_err:
        logger.debug('Error checking transform identity: {}'.format(str(check_err)))
        return True  # If we can't check, assume identity (safer)

# =============================================================================
# UPDATED copy_elevation_markers FUNCTION
# Replace your existing copy_elevation_markers function with this version:
# =============================================================================

def copy_elevation_markers(activedoc, dest_doc):
    """Copy all elevation markers from source to destination document.
    
    This will automatically create the elevation views in the destination.
    Handles coordinate transformation for models at different locations.
    
    Args:
        activedoc - Source Revit document
        dest_doc - Destination Revit document
        
    Returns:
        int - Number of markers successfully copied
    """
    logger.debug('Collecting elevation markers from source document...')
    
    elevation_markers = collect_elevation_markers_from_plans(activedoc)
    
    if not elevation_markers:
        logger.debug('No elevation markers found in source document.')
        return 0
    
    logger.debug('Found {} elevation marker(s) to copy'.format(len(elevation_markers)))
    print('Copying {} elevation marker(s) (this will create elevation views)...'.format(
        len(elevation_markers)
    ))
    
    # Calculate coordinate transformation between documents
    coord_transform = get_coordinate_transform_between_documents(activedoc, dest_doc)
    
    if coord_transform:
        print('  NOTE: Applying coordinate transformation for different project locations')
        print('        (Models are linked by shared coordinates)')
    else:
        print('  NOTE: Both models at same location - no coordinate adjustment needed')
    
    marker_ids = [marker.Id for marker in elevation_markers]
    
    try:
        with revit.Transaction('Copy Elevation Markers', doc=dest_doc):
            cp_options = DB.CopyPasteOptions()
            cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())
            
            # Copy markers with coordinate transformation if needed
            copied_ids = DB.ElementTransformUtils.CopyElements(
                activedoc,
                List[DB.ElementId](marker_ids),
                dest_doc,
                coord_transform,  # Apply transform (or None if not needed)
                cp_options
            )
            
            if copied_ids:
                print('Successfully copied {} elevation marker(s)'.format(copied_ids.Count))
                print('Elevation views have been automatically created in destination model.')
                
                if coord_transform:
                    print('Markers have been transformed to match destination coordinate system.')
                    
                return copied_ids.Count
            else:
                logger.warning('No elevation markers were copied')
                return 0
                
    except Exception as marker_err:
        logger.error('Error copying elevation markers: {}'.format(str(marker_err)))
        print('WARNING: Could not copy elevation markers.')
        print('         Elevation views will be skipped.')
        print('         Error: {}'.format(str(marker_err)))
        return 0


def get_default_type(source_doc, type_group):
    return source_doc.GetDefaultElementTypeId(type_group)


def find_matching_elevation_view(dest_doc, source_elev_view, source_doc):
    """
    Find matching elevation view in destination by comparing view properties.
    Elevation views created from copied markers may have different names,
    so we need to match by orientation and location instead.
    """
    # Get source view properties for comparison
    try:
        source_origin = source_elev_view.Origin
        source_direction = source_elev_view.ViewDirection
        source_up = source_elev_view.UpDirection
    except:
        return None
    
    # Collect all elevation views in destination
    dest_elevations = DB.FilteredElementCollector(dest_doc)\
                        .OfClass(DB.View)\
                        .ToElements()
    
    # Find elevation with matching properties
    for dest_view in dest_elevations:
        if dest_view.ViewType != DB.ViewType.Elevation:
            continue
        
        try:
            dest_origin = dest_view.Origin
            dest_direction = dest_view.ViewDirection
            dest_up = dest_view.UpDirection
            
            # Compare view properties (with tolerance for floating point)
            origin_match = source_origin.IsAlmostEqualTo(dest_origin, 0.01)
            direction_match = source_direction.IsAlmostEqualTo(dest_direction, 0.01)
            up_match = source_up.IsAlmostEqualTo(dest_up, 0.01)
            
            if origin_match and direction_match and up_match:
                logger.debug('Found matching elevation by geometry: {} -> {}'.format(
                    source_elev_view.Name,
                    dest_view.Name
                ))
                return dest_view
        except:
            continue
    
    return None


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
    # collect guides in dest_doc
    guide_elements = \
        DB.FilteredElementCollector(source_doc)\
            .OfCategory(DB.BuiltInCategory.OST_GuideGrid)\
            .WhereElementIsNotElementType()\
            .ToElements()
    
    # find guide with same name
    for guide in guide_elements:
        if str(guide.Name).lower() == guide_name.lower():
            return guide


def can_copy_view_contents(view):
    """Check if a view type supports content copying via ElementTransformUtils."""
    
    # Views that CANNOT be used as source/destination for CopyElements
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
        DB.ViewType.ThreeD  # 3D views cannot be used for content copying
    ]
    
    if view.ViewType in unsupported_types:
        return False
    
    # Elevation and Section views are model-based, not annotation-based
    # They typically should not have content "copied" in the traditional sense
    # Content copying should be skipped for these view types
    if view.ViewType in [DB.ViewType.Elevation, 
                         DB.ViewType.Section]:
        return False
    
    # Additional checks for problematic views
    try:
        # Check if view is a template (cannot copy contents)
        if view.IsTemplate:
            return False
        
        # Check if view is a dependent view
        view_id = view.Id
        primary_view_id = view.GetPrimaryViewId()
        if primary_view_id != DB.ElementId.InvalidElementId:
            # This is a dependent view - skip content copy
            return False
            
    except Exception:
        # If any check fails, assume unsafe
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
    # check to see if revision exists
    for rev in all_dest_revs:
        if query.compare_revisions(rev, src_rev):
            return rev

    # if no matching revisions found, create a new revision and return
    logger.warning('Revision could not be found in destination model.\n'
                   'Revision Date: {}\n'
                   'Revision Description: {}\n'
                   'Creating a new revision. Please review revisions '
                   'after copying process is finished.'
                   .format(src_rev.RevisionDate, src_rev.Description))
    return revit.create.create_revision(description=src_rev.Description,
                                        by=src_rev.IssuedBy,
                                        to=src_rev.IssuedTo,
                                        date=src_rev.RevisionDate,
                                        doc=dest_doc)


def clear_view_contents(dest_doc, dest_view):
    logger.debug('Removing view contents: {}'.format(dest_view.Name))
    elements_ids = get_view_contents(dest_doc, dest_view)

    with revit.Transaction('Delete View Contents', doc=dest_doc):
        for el_id in elements_ids:
            try:
                dest_doc.Delete(el_id)
            except Exception as err:
                continue

    return True


def copy_view_contents(activedoc, source_view, dest_doc, dest_view,
                       clear_contents=False):
    """Copy view contents with robust error handling for unsupported view types."""
    
    logger.debug('Copying view contents: {} : {}'
                 .format(source_view.Name, source_view.ViewType))

    # Check if source view supports content copying
    if not can_copy_view_contents(source_view):
        logger.warning('Skipping content copy for view type: {} ({})'
                       .format(source_view.Name, source_view.ViewType))
        print('\t\t\tView type does not support content copying - skipping.')
        return True
    
    # Check if destination view supports content copying
    if not can_copy_view_contents(dest_view):
        logger.warning('Destination view cannot receive copied content: {} ({})'
                       .format(dest_view.Name, dest_view.ViewType))
        return True

    elements_ids = get_view_contents(activedoc, source_view)

    if clear_contents:
        if not clear_view_contents(dest_doc, dest_view):
            return False

    if not elements_ids:
        logger.debug('No elements to copy in view: {}'.format(source_view.Name))
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
        logger.error('Failed to copy view contents for {}: {}'
                     .format(source_view.Name, str(copy_err)))
        print('\t\t\tWARNING: Could not copy view contents - view will be empty.')
        # Return True to continue processing other views
        return True


def copy_view_props(source_view, dest_view):
    """Copy view properties with error handling."""
    try:
        dest_view.Scale = source_view.Scale
    except Exception as scale_err:
        logger.debug('Could not copy scale: {}'.format(str(scale_err)))
    
    try:
        dest_view.Parameter[VIEW_TOS_PARAM].Set(
            source_view.Parameter[VIEW_TOS_PARAM].AsString()
        )
    except Exception as param_err:
        logger.debug('Could not copy view description: {}'.format(str(param_err)))


def copy_view(activedoc, source_view, dest_doc):
    matching_view = find_matching_view(dest_doc, source_view)
    if matching_view:
        print('\t\t\tView/Sheet already exists in document.')
        if OPTION_SET.op_update_exist_view_contents:
            if not copy_view_contents(activedoc,
                                      source_view,
                                      dest_doc,
                                      matching_view,
                                      clear_contents=True):
                logger.error('Could not copy view contents: {}'
                             .format(source_view.Name))

        return matching_view

    logger.debug('Copying view: {} (Type: {})'.format(
        source_view.Name, 
        source_view.ViewType
    ))
    new_view = None

    if source_view.ViewType == DB.ViewType.DrawingSheet:
        try:
            logger.debug('Source view is a sheet. '
                         'Creating destination sheet.')

            with revit.Transaction('Create Sheet', doc=dest_doc):
                if not source_view.IsPlaceholder \
                        or (source_view.IsPlaceholder
                                and OPTION_SET.op_copy_placeholders_as_sheets):
                    new_view = \
                        DB.ViewSheet.Create(
                            dest_doc,
                            DB.ElementId.InvalidElementId
                            )
                else:
                    new_view = DB.ViewSheet.CreatePlaceholder(dest_doc)

                revit.update.set_name(new_view,
                                      revit.query.get_name(source_view))
                new_view.SheetNumber = source_view.SheetNumber
        except Exception as sheet_err:
            logger.error('Error creating sheet. | {}'.format(sheet_err))
            
    elif source_view.ViewType == DB.ViewType.DraftingView:
        try:
            logger.debug('Source view is a drafting. '
                         'Creating destination drafting view.')

            with revit.Transaction('Create Drafting View', doc=dest_doc):
                new_view = DB.ViewDrafting.Create(
                    dest_doc,
                    get_default_type(dest_doc,
                                     DB.ElementTypeGroup.ViewTypeDrafting)
                )
                revit.update.set_name(new_view,
                                      revit.query.get_name(source_view))
                copy_view_props(source_view, new_view)
        except Exception as sheet_err:
            logger.error('Error creating drafting view. | {}'
                         .format(sheet_err))
            
    elif source_view.ViewType == DB.ViewType.Legend:
        try:
            logger.debug('Source view is a legend. '
                         'Creating destination legend view.')

            first_legend = query.find_first_legend(dest_doc)
            if first_legend:
                with revit.Transaction('Create Legend View', doc=dest_doc):
                    new_view = \
                        dest_doc.GetElement(
                            first_legend.Duplicate(
                                DB.ViewDuplicateOption.Duplicate
                                )
                            )
                    revit.update.set_name(new_view,
                                        revit.query.get_name(source_view))
                    copy_view_props(source_view, new_view)
            else:
                logger.error('Destination document must have at least one '
                             'Legend view. Skipping legend.')
        except Exception as sheet_err:
            logger.error('Error creating drafting view. | {}'
                         .format(sheet_err))
    
    elif source_view.ViewType == DB.ViewType.Elevation:
        # Elevation views are created by copying elevation markers in plan views
        # Names may differ, so we match by geometry (origin, direction, up vector)
        logger.debug('Looking for elevation view by geometry: {}'.format(source_view.Name))
        
        # First try name match
        matching_view = find_matching_view(dest_doc, source_view)
        
        # If name match fails, try geometric match
        if not matching_view:
            matching_view = find_matching_elevation_view(dest_doc, source_view, activedoc)
        
        if matching_view:
            print('\t\t\tFound matching elevation view: {}'.format(matching_view.Name))
            new_view = matching_view
            
            # Copy the name from source to make it consistent
            try:
                with revit.Transaction('Rename Elevation View', doc=dest_doc):
                    revit.update.set_name(matching_view, query.get_name(source_view))
                    print('\t\t\tRenamed to: {}'.format(query.get_name(source_view)))
            except Exception as rename_err:
                logger.debug('Could not rename elevation view: {}'.format(rename_err))
        else:
            logger.warning('Elevation view not found in destination. '
                          'Elevation marker may not have been copied. '
                          'Skipping view: {}'.format(source_view.Name))
            print('\t\t\tSkipping - No matching elevation found in destination.')
            return None
    
    elif source_view.ViewType in [DB.ViewType.Section, DB.ViewType.Detail]:
        try:
            logger.debug('Source view is a section/detail. '
                         'Using direct copy method.')
            
            # Sections can be copied directly to the destination document
            with revit.Transaction('Copy Section/Detail View', doc=dest_doc):
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
                    
                    # Try to apply matching view template
                    source_template_id = source_view.ViewTemplateId
                    if source_template_id != DB.ElementId.InvalidElementId:
                        source_template = activedoc.GetElement(source_template_id)
                        if source_template:
                            # Look for matching template by name in dest
                            for template in DB.FilteredElementCollector(dest_doc)\
                                              .OfClass(DB.View)\
                                              .ToElements():
                                if template.IsTemplate \
                                        and query.get_name(template) == query.get_name(source_template):
                                    try:
                                        new_view.ViewTemplateId = template.Id
                                    except:
                                        pass
                                    break
                    
                    print('\t\t\tSection/Detail view copied successfully.')
                else:
                    logger.error('Copy operation returned no elements')
                    
        except Exception as section_err:
            logger.error('Error copying section/detail view: {}'
                         .format(section_err))
            print('\t\t\tWARNING: Could not copy section/detail view.')
    
    else:
        # Handle any other view types (FloorPlan, CeilingPlan, ThreeD, etc.)
        try:
            logger.debug('Source view is type: {}. Attempting direct copy.'.format(
                source_view.ViewType
            ))
            
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
                    print('\t\t\t{} view copied successfully.'.format(
                        source_view.ViewType
                    ))
                else:
                    logger.error('Copy operation returned no elements')
                    
        except Exception as view_err:
            logger.error('Error copying view type {}: {}'.format(
                source_view.ViewType,
                view_err
            ))
            print('\t\t\tWARNING: Could not copy view type: {}'.format(
                source_view.ViewType
            ))

    if new_view:
        copy_view_contents(activedoc, source_view, dest_doc, new_view)

    return new_view


def copy_viewport_types(activedoc, vport_type, vport_typename,
                        dest_doc, newvport):
    dest_vport_typenames = [DBElement.Name.GetValue(dest_doc.GetElement(x))
                            for x in newvport.GetValidTypes()]

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
            vport_typename = DBElement.Name.GetValue(vport_type)

            newvport = dest_doc.GetElement(newvport_id)

            copy_viewport_types(activedoc, vport_type, vport_typename,
                                dest_doc, newvport)

            for vtype_id in newvport.GetValidTypes():
                vtype = dest_doc.GetElement(vtype_id)
                if DBElement.Name.GetValue(vtype) == vport_typename:
                    newvport.ChangeTypeId(vtype_id)
                    break
    except Exception as type_err:
        logger.debug('Could not apply viewport type: {}'.format(str(type_err)))


def copy_sheet_viewports(activedoc, source_sheet, dest_doc, dest_sheet):
    """Copy viewports with robust error handling and proper elevation placement."""
    
    global elevations_skipped
    
    existing_views = [dest_doc.GetElement(x).ViewId
                      for x in dest_sheet.GetAllViewports()]

    viewport_errors = []

    for vport_id in source_sheet.GetAllViewports():
        try:
            vport = activedoc.GetElement(vport_id)
            vport_view = activedoc.GetElement(vport.ViewId)

            print('\t\tCopying/updating view: {}'
                  .format(revit.query.get_name(vport_view)))
            
            is_elevation = vport_view.ViewType == DB.ViewType.Elevation
            
            # Get or create the view
            try:
                new_view = copy_view(activedoc, vport_view, dest_doc)

                if new_view:
                    # Check if view is already on THIS sheet
                    if new_view.Id in existing_views:
                        print('\t\t\tView already exists on this sheet.')
                        # Still apply viewport type to match source
                        try:
                            existing_vport_ids = [x for x in dest_sheet.GetAllViewports() 
                                                 if dest_doc.GetElement(x).ViewId == new_view.Id]
                            if existing_vport_ids:
                                apply_viewport_type(activedoc, vport_id,
                                                  dest_doc, existing_vport_ids[0])
                        except:
                            pass
                        continue
                    
                    # Check if view is placed on a DIFFERENT sheet (common after ElementTransformUtils.CopyElements)
                    ref_info = revit.query.get_view_sheetrefinfo(new_view)
                    
                    if ref_info and ref_info.sheet_num:
                        # View is on another sheet - need to remove it first
                        if ref_info.sheet_num == dest_sheet.SheetNumber:
                            # Already on target sheet (shouldn't happen but safety check)
                            print('\t\t\tView already on target sheet.')
                            continue
                        else:
                            # View is on a different sheet (likely temporary "00100" from copy operation)
                            # Remove it from that sheet first
                            print('\t\t\tRemoving view from temporary sheet {}...'.format(
                                ref_info.sheet_num))
                            try:
                                # Find all sheets in destination
                                all_sheets = DB.FilteredElementCollector(dest_doc)\
                                               .OfClass(DB.ViewSheet)\
                                               .ToElements()
                                
                                # Find the sheet with matching number
                                for temp_sheet in all_sheets:
                                    if temp_sheet.SheetNumber == ref_info.sheet_num:
                                        # Find viewport on this sheet with our view
                                        for vp_id in temp_sheet.GetAllViewports():
                                            vp = dest_doc.GetElement(vp_id)
                                            if vp and vp.ViewId == new_view.Id:
                                                with revit.Transaction('Remove Viewport', doc=dest_doc):
                                                    dest_doc.Delete(vp_id)
                                                logger.debug('Removed viewport from sheet {}'.format(
                                                    ref_info.sheet_num))
                                                break
                                        break
                            except Exception as remove_err:
                                logger.error('Could not remove view from temp sheet: {}'.format(
                                    str(remove_err)))
                                print('\t\t\tWARNING: Could not remove from temp sheet - will try to place anyway.')
                    
                    # Now place on target sheet
                    print('\t\t\tPlacing view on target sheet.')
                    with revit.Transaction('Place View on Sheet', doc=dest_doc):
                        nvport = DB.Viewport.Create(dest_doc,
                                                    dest_sheet.Id,
                                                    new_view.Id,
                                                    vport.GetBoxCenter())
                    if nvport:
                        apply_viewport_type(activedoc, vport_id,
                                          dest_doc, nvport.Id)
                        
                else:
                    # View creation returned None
                    if is_elevation:
                        elevations_skipped += 1
                    else:
                        viewport_errors.append(
                            'Could not copy view: {}'.format(
                                revit.query.get_name(vport_view)
                            )
                        )
                    
            except Exception as vport_err:
                error_msg = 'Failed to process viewport for view {}: {}'.format(
                    revit.query.get_name(vport_view),
                    str(vport_err)
                )
                logger.error(error_msg)
                viewport_errors.append(error_msg)
                print('\t\t\tWARNING: Skipping problematic viewport.')
                continue
                
        except Exception as outer_err:
            logger.error('Unexpected error in viewport processing: {}'.format(
                str(outer_err)
            ))
            continue
    
    # Report summary of viewport errors if any occurred
    if viewport_errors:
        print('\t\tWARNING: {} viewport(s) could not be copied:'.format(
            len(viewport_errors)
        ))
        for err in viewport_errors[:3]:  # Show first 3 errors
            print('\t\t  - {}'.format(err))
        if len(viewport_errors) > 3:
            print('\t\t  ... and {} more'.format(len(viewport_errors) - 3))


def copy_sheet_revisions(activedoc, source_sheet, dest_doc, dest_sheet):
    all_src_revs = query.get_revisions(doc=activedoc)
    all_dest_revs = query.get_revisions(doc=dest_doc)
    revisions_to_set = []

    with revit.Transaction('Copy and Set Revisions', doc=dest_doc):
        for src_revid in source_sheet.GetAdditionalRevisionIds():
            set_rev = ensure_dest_revision(activedoc.GetElement(src_revid),
                                           all_dest_revs,
                                           dest_doc)
            revisions_to_set.append(set_rev)

        if revisions_to_set:
            revit.update.update_sheet_revisions(revisions_to_set,
                                                [dest_sheet],
                                                state=True,
                                                doc=dest_doc)


def copy_sheet_guides(activedoc, source_sheet, dest_doc, dest_sheet):
    # sheet guide
    source_sheet_guide_param = \
        source_sheet.Parameter[DB.BuiltInParameter.SHEET_GUIDE_GRID]
    source_sheet_guide_element = \
        activedoc.GetElement(source_sheet_guide_param.AsElementId())
    
    if source_sheet_guide_element:
        if not find_guide(source_sheet_guide_element.Name, dest_doc):
            # copy guides to dest_doc
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
            # set the guide
            with revit.Transaction('Set Sheet Guide', doc=dest_doc):
                dest_sheet_guide_param = \
                    dest_sheet.Parameter[DB.BuiltInParameter.SHEET_GUIDE_GRID]
                dest_sheet_guide_param.Set(dest_guide.Id)
        else:
            logger.error('Error copying and setting sheet guide for sheet {}'
                         .format(source_sheet.Name))


def copy_sheet(activedoc, source_sheet, dest_doc):
    """Copy sheet with comprehensive error handling and recovery."""
    
    logger.debug('Copying sheet {} to document {}'
                 .format(source_sheet.Name, dest_doc.Title))
    print('\tCopying/updating Sheet: {}'.format(source_sheet.Name))
    
    sheet_success = False
    
    try:
        with revit.TransactionGroup('Import Sheet', doc=dest_doc):
            logger.debug('Creating destination sheet...')
            new_sheet = copy_view(activedoc, source_sheet, dest_doc)

            if new_sheet:
                if not new_sheet.IsPlaceholder:
                    if OPTION_SET.op_copy_vports:
                        logger.debug('Copying sheet viewports...')
                        try:
                            copy_sheet_viewports(activedoc, source_sheet,
                                                 dest_doc, new_sheet)
                        except Exception as vport_err:
                            logger.error('Error copying viewports: {}'.format(
                                str(vport_err)
                            ))
                            print('\t\tWARNING: Some viewports may not have copied.')
                    else:
                        print('Skipping viewports...')

                    if OPTION_SET.op_copy_guides:
                        logger.debug('Copying sheet guide grids...')
                        try:
                            copy_sheet_guides(activedoc, source_sheet,
                                            dest_doc, new_sheet)
                        except Exception as guide_err:
                            logger.error('Error copying guides: {}'.format(
                                str(guide_err)
                            ))
                            print('\t\tWARNING: Could not copy guide grids.')
                    else:
                        print('Skipping sheet guides...')

                if OPTION_SET.op_copy_revisions:
                    logger.debug('Copying sheet revisions...')
                    try:
                        copy_sheet_revisions(activedoc, source_sheet,
                                           dest_doc, new_sheet)
                    except Exception as rev_err:
                        logger.error('Error copying revisions: {}'.format(
                            str(rev_err)
                        ))
                        print('\t\tWARNING: Could not copy revisions.')
                else:
                    print('Skipping revisions...')
                
                sheet_success = True

            else:
                logger.error('Failed copying sheet: {}'.format(source_sheet.Name))
                
    except Exception as sheet_err:
        logger.error('Critical error copying sheet {}: {}'.format(
            source_sheet.Name,
            str(sheet_err)
        ))
        print('\t\tERROR: Sheet copy failed - {}'.format(str(sheet_err)))
    
    return sheet_success


# ============================================================================
# MAIN EXECUTION
# ============================================================================

dest_docs = get_dest_docs()
doc_count = len(dest_docs)

source_sheets = get_source_sheets()
sheet_count = len(source_sheets)

OPTION_SET = get_user_options()

total_work = doc_count * sheet_count
work_counter = 0

# Track success/failure
sheets_succeeded = 0
sheets_failed = 0
elevations_skipped = 0

# Copy elevation markers first if option is enabled
# This creates elevation views automatically in destination documents
if OPTION_SET.op_copy_elevation_markers:
    output.print_md('---')
    output.print_md('**STEP 1: COPYING ELEVATION MARKERS**')
    output.print_md('*Elevation markers will create elevation views in destination models*')
    output.print_md('')
    
    for dest_doc in dest_docs:
        output.print_md('Copying elevation markers to: **{}**'.format(dest_doc.Title))
        markers_copied = copy_elevation_markers(revit.doc, dest_doc)
    
    output.print_md('---')
    output.print_md('**STEP 2: COPYING SHEETS AND VIEWPORTS**')
    output.print_md('')

for dest_doc in dest_docs:
    output.print_md('**Copying Sheet(s) to Document:** {0}'
                    .format(dest_doc.Title))

    for source_sheet in source_sheets:
        print('Copying Sheet: {0} - {1}'.format(source_sheet.SheetNumber,
                                                source_sheet.Name))
        
        if copy_sheet(revit.doc, source_sheet, dest_doc):
            sheets_succeeded += 1
        else:
            sheets_failed += 1
            
        work_counter += 1
        output.update_progress(work_counter, total_work)

output.print_md('---')
output.print_md('**COPY COMPLETE**')
output.print_md('- Successfully copied: {} sheets'.format(sheets_succeeded))
if sheets_failed > 0:
    output.print_md('- Failed to copy: {} sheets (see warnings above)'.format(
        sheets_failed
    ))
if elevations_skipped > 0:
    output.print_md('- **Note:** {} elevation view(s) were skipped'.format(
        elevations_skipped
    ))
    output.print_md('  - *Elevation views require plan view markers and cannot be copied between documents*')
    output.print_md('  - *To include elevations: manually create elevation markers in the destination model*')
output.print_md('- Target documents: {}'.format(doc_count))
