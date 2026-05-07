# -*- coding: utf-8 -*-
"""Sums up the values of selected numerical parameter on selected elements.
Sum is calculated for every Type and overall.

NEW: Count of elements per type is also available.
"""

from collections import namedtuple

from pyrevit import revit, DB, HOST_APP
from pyrevit import forms
from pyrevit import script
from pyrevit.coreutils import pyutils

__context__ = 'selection'
__doc__ = '''Sums up the values of selected numerical parameter on selected elements.
Sum is calculated for every Type and overall.

NEW:
Count of elements per type is also available.'''
__title__ = 'Sum\nParameters'
__helpurl__ = 'https://youtu.be/aWg-rj8k0Ts'
__highlight__ = 'updated'


# Get selection and output
selection = revit.get_selection()
logger = script.get_logger()
output = script.get_output()

# FIXED: Removed ct_icon() call that was causing the error
# The custom icon functionality has been removed to prevent path errors

ParamDef = namedtuple('ParamDef', ['name', 'type', 'readableTypeName'])


def is_calculable_param(param):
    """Check if parameter can be summed (numeric types only)"""
    if param.StorageType == DB.StorageType.Double:
        return True

    if param.StorageType == DB.StorageType.Integer:
        val_str = param.AsValueString()
        if val_str and unicode(val_str).lower().isdigit():
            return True

    return False


def get_definition_type(definition):
    """Get parameter type (compatible with different Revit versions)"""
    if HOST_APP.is_newer_than(2022):
        return definition.GetDataType()
    else:
        return definition.ParameterType


def get_definition_readableTypeName(definition):
    """Get human-readable parameter type name"""
    if HOST_APP.is_newer_than(2022):
        return DB.LabelUtils.GetLabelForSpec(definition.GetDataType())
    else:
        return definition.ParameterType


def calc_param_total(element_list, param_name):
    """Calculate total value of parameter across all elements"""
    sum_total = 0.0

    def _add_total(total, param):
        if param.StorageType == DB.StorageType.Double:
            total += param.AsDouble()
        elif param.StorageType == DB.StorageType.Integer:
            total += param.AsInteger()
        return total

    for el in element_list:
        param = el.LookupParameter(param_name)
        if not param:
            # Try to get parameter from element type
            el_type = revit.doc.GetElement(el.GetTypeId())
            type_param = el_type.LookupParameter(param_name) if el_type else None
            if not type_param:
                logger.error('Element with ID: {} '
                           'does not have parameter: {}.'.format(el.Id,
                                                                param_name))
            else:
                sum_total = _add_total(sum_total, type_param)
        else:
            sum_total = _add_total(sum_total, param)

    return sum_total


def format_length(total):
    """Convert from feet to meters with 3 decimal places"""
    return "{:.3f} m".format(total / 3.28084)


def format_area(total):
    """Convert from square feet to square meters with 3 decimal places"""
    return "{:.3f} m2".format(total / 10.7639)


def format_volume(total):
    """Convert from cubic feet to cubic meters with 3 decimal places"""
    return "{:.3f} m3".format(total / 35.3147)


# Setup formatter functions based on Revit version
if HOST_APP.is_newer_than(2022):
    formatter_funcs = {
        DB.SpecTypeId.Length: format_length,
        DB.SpecTypeId.Area: format_area,
        DB.SpecTypeId.Volume: format_volume
    }
else:
    formatter_funcs = {
        DB.ParameterType.Length: format_length,
        DB.ParameterType.Area: format_area,
        DB.ParameterType.Volume: format_volume
    }


def process_options(element_list):
    """Find all calculable parameters shared by selected elements"""
    param_sets = []

    for el in element_list:
        shared_params = set()
        
        # Find element instance parameters
        for param in el.ParametersMap:
            if is_calculable_param(param):
                pdef = param.Definition
                shared_params.add(ParamDef(
                    pdef.Name,
                    get_definition_type(pdef),
                    get_definition_readableTypeName(pdef)
                ))

        # Find element type parameters
        el_type = revit.doc.GetElement(el.GetTypeId())
        if el_type and el_type.Id != DB.ElementId.InvalidElementId:
            for type_param in el_type.ParametersMap:
                if is_calculable_param(type_param):
                    pdef = type_param.Definition
                    shared_params.add(ParamDef(
                        pdef.Name,
                        get_definition_type(pdef),
                        get_definition_readableTypeName(pdef)
                    ))

        param_sets.append(shared_params)

    # Find parameters common to all selected elements
    if param_sets:
        all_shared_params = param_sets[0]
        for param_set in param_sets[1:]:
            all_shared_params = all_shared_params.intersection(param_set)

        return {'{} <{}>'.format(x.name, x.readableTypeName): x
                for x in all_shared_params}

    return None


def process_sets(element_list):
    """Group elements by their type"""
    el_sets = pyutils.DefaultOrderedDict(list)

    # Separate elements into sets based on their type
    for el in element_list:
        if hasattr(el, 'LineStyle'):
            el_sets[el.LineStyle.Name].append(el)
        else:
            eltype = revit.doc.GetElement(el.GetTypeId())
            if eltype:
                el_sets[revit.query.get_name(eltype)].append(el)
    
    # Add all elements as last set for totals
    el_sets['SPOLU'].extend(element_list)

    return el_sets


def get_document_name():
    """Get document name safely"""
    try:
        doc = revit.doc
        if doc.IsWorkshared:
            return doc.Title
        else:
            return doc.Title if doc.Title else "Untitled"
    except:
        return "Document"


# MAIN EXECUTION -------------------------------------------------------------

# Check if elements are selected
if not selection.elements:
    forms.alert('No elements selected. Please select elements and try again.',
                title='No Selection')
    script.exit()

# Get available parameters from selection
options = process_options(selection.elements)

# Add "Count" option which counts elements rather than summing parameters
if options:
    options["Count"] = "count"
else:
    forms.alert('Selected elements do not have any calculable parameters.',
                title='No Calculable Parameters')
    script.exit()

# Ask user to select parameter to sum
selected_switch = forms.CommandSwitchWindow.show(
    sorted(options),
    message='Select parameter to sum:'
)

if not selected_switch:
    script.exit()

selected_option = options[selected_switch]

# Calculate and display results
if selected_option == "count":
    # COUNT MODE: Count elements per type
    output.print_md("## Count")
    output.print_md("### {}".format(get_document_name()))
    md_schedule = "| Family Type | Count |\n| ----------- | ----------- |"
    
    for type_name, element_set in process_sets(selection.elements).items():
        # Escape special characters for markdown
        type_name = type_name.replace('<', '&lt;').replace('>', '&gt;')
        total_value = len(element_set)
        
        # Highlight total row
        if type_name == "SPOLU":
            strong_tag1 = "<span style='color:darkorange'>**"
            strong_tag2 = "**</span>"
        else:
            strong_tag1 = strong_tag2 = ""
        
        new_schedule_line = "\n| {} {} {} | {} {} {} |".format(
            strong_tag1, type_name, strong_tag2,
            strong_tag1, str(total_value), strong_tag2
        )
        md_schedule += new_schedule_line

else:
    # SUM MODE: Sum parameter values per type
    output.print_md("## {}".format(selected_option.name))
    output.print_md("### {}".format(get_document_name()))
    md_schedule = "| Family Type | Parameter Value |\n| ----------- | ----------- |"
    
    for type_name, element_set in process_sets(selection.elements).items():
        # Escape special characters for markdown
        type_name = type_name.replace('<', '&lt;').replace('>', '&gt;')
        total_value = calc_param_total(element_set, selected_option.name)
        
        # Highlight total row
        if type_name == "SPOLU":
            strong_tag1 = "<span style='color:darkorange'>**"
            strong_tag2 = "**</span>"
        else:
            strong_tag1 = strong_tag2 = ""
        
        # Format value based on parameter type
        if selected_option.type in formatter_funcs.keys():
            formatted_value = formatter_funcs[selected_option.type](total_value)
        else:
            formatted_value = str(total_value)
        
        new_schedule_line = "\n| {} {} {} | {} {} {} |".format(
            strong_tag1, type_name, strong_tag2,
            strong_tag1, formatted_value, strong_tag2
        )
        md_schedule += new_schedule_line

# Print the results table
output.print_md(md_schedule)