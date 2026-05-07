# -*- coding: utf-8 -*-
"""
LEGEND GENERATOR MODULE
========================
Reusable legend generation from view filters
Extracted for use by Drawing Package Automation

Author: Anirudh Sood
Compatible: Revit 2023-2025, Python 2.7
"""

from Autodesk.Revit.DB import *
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# ================================================================================================
# CONFIGURATION
# ================================================================================================

class LegendConfig:
    """Default legend dimensions (in millimeters)"""
    REGION_WIDTH_MM = 500.0
    REGION_HEIGHT_MM = 500.0
    SPACING_MM = 100.0
    LINE_WIDTH_MM = 500.0


# ================================================================================================
# HELPER FUNCTIONS
# ================================================================================================

def convert_mm_to_feet(mm):
    """Convert millimeters to feet."""
    return mm / 304.8


def create_text_note(doc, view, x, y, text, text_type):
    """Create a text note at specified coordinates."""
    point = XYZ(x, y, 0)
    text_note = TextNote.Create(doc, view.Id, point, text, text_type.Id)
    return text_note


def create_region(doc, view, x, y, width, height):
    """Create a filled region."""
    half_width = width / 2.0
    half_height = height / 2.0
    
    profile_loop = CurveLoop()
    p1 = XYZ(x - half_width, y - half_height, 0)
    p2 = XYZ(x + half_width, y - half_height, 0)
    p3 = XYZ(x + half_width, y + half_height, 0)
    p4 = XYZ(x - half_width, y + half_height, 0)
    
    profile_loop.Append(Line.CreateBound(p1, p2))
    profile_loop.Append(Line.CreateBound(p2, p3))
    profile_loop.Append(Line.CreateBound(p3, p4))
    profile_loop.Append(Line.CreateBound(p4, p1))
    
    profile_loops = List[CurveLoop]()
    profile_loops.Add(profile_loop)
    
    filled_region_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.FilledRegionType)
    filled_region = FilledRegion.Create(doc, filled_region_type_id, view.Id, profile_loops)
    
    return filled_region


def create_horizontal_line(doc, view, x, y, width):
    """Create a horizontal detail line."""
    start_point = XYZ(x, y, 0)
    end_point = XYZ(x + width, y, 0)
    line = Line.CreateBound(start_point, end_point)
    detail_line = doc.Create.NewDetailCurve(view, line)
    return detail_line


def override_graphics_region(doc, view, region, fg_pattern_id=None, fg_color=None, 
                              bg_pattern_id=None, bg_color=None, line_color=None, 
                              line_pattern_id=None, lineweight=None):
    """Override graphics settings for a filled region."""
    settings = OverrideGraphicSettings()
    
    if fg_pattern_id and fg_pattern_id != ElementId.InvalidElementId:
        settings.SetSurfaceForegroundPatternId(fg_pattern_id)
    if fg_color:
        settings.SetSurfaceForegroundPatternColor(fg_color)
    if bg_pattern_id and bg_pattern_id != ElementId.InvalidElementId:
        settings.SetSurfaceBackgroundPatternId(bg_pattern_id)
    if bg_color:
        settings.SetSurfaceBackgroundPatternColor(bg_color)
    if line_color:
        settings.SetProjectionLineColor(line_color)
    if line_pattern_id and line_pattern_id != ElementId.InvalidElementId:
        settings.SetProjectionLinePatternId(line_pattern_id)
    if lineweight and lineweight != -1:
        settings.SetProjectionLineWeight(lineweight)
    
    view.SetElementOverrides(region.Id, settings)


def override_graphics_line(doc, view, line, line_color=None, line_pattern_id=None, lineweight=None):
    """Override graphics settings for a detail line."""
    settings = OverrideGraphicSettings()
    
    if line_color:
        settings.SetProjectionLineColor(line_color)
    if line_pattern_id and line_pattern_id != ElementId.InvalidElementId:
        settings.SetProjectionLinePatternId(line_pattern_id)
    if lineweight and lineweight != -1:
        settings.SetProjectionLineWeight(lineweight)
    
    view.SetElementOverrides(line.Id, settings)


def get_workset_names(doc):
    """Get workset names dictionary"""
    try:
        all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
        return {int(str(w.Id)): w.Name for w in all_worksets}
    except:
        return {}


# ================================================================================================
# FILTER RULE PARSING
# ================================================================================================

class FilterRule:
    """Parse and store filter rule information."""
    
    dict_BIPs = {str(i.value__): i for i in BuiltInParameter.GetValues(BuiltInParameter)}
    
    def __init__(self, doc, rule):
        self.doc = doc
        self.rule = rule
        self.workset_names = get_workset_names(doc)
        self.rule_value = None
        self.rule_eval = None
        self.get_rule_value()
    
    @property
    def rule_param_name(self):
        """Get the parameter name from the rule."""
        rule_param_id = self.rule.GetRuleParameter()
        
        # Try shared parameter
        rule_shared_param = self.doc.GetElement(rule_param_id)
        if rule_shared_param:
            return rule_shared_param.Name
        
        # Try built-in parameter
        try:
            bip_rule_param = self.dict_BIPs[str(rule_param_id)]
            readable_bip = LabelUtils.GetLabelFor(bip_rule_param)
            return readable_bip
        except:
            return "Unknown Parameter"
    
    def get_rule_value(self):
        """Extract rule evaluator and value."""
        rule_evaluator = ''
        rule_value = ''
        inverse = False
        
        # Handle inverse rules
        if type(self.rule) == FilterInverseRule:
            inverse = True
            rule = self.rule.GetInnerRule()
        else:
            rule = self.rule
        
        # Get evaluator
        try:
            rule_evaluator = rule.GetEvaluator().ToString().split('Filter')[1].replace("String", "").replace("Numeric", "")
            rule_evaluator = 'not {}'.format(rule_evaluator) if inverse else rule_evaluator
        except:
            if type(rule) == HasValueFilterRule:
                rule_evaluator = 'HasValue'
            elif type(rule) == HasNoValueFilterRule:
                rule_evaluator = 'HasNoValue'
            elif type(rule) == SharedParameterApplicableRule:
                rule_evaluator = 'Exists'
            else:
                rule_evaluator = 'Unknown'
        
        # Get value based on rule type
        if type(rule) == FilterStringRule:
            rule_value = rule.RuleString
        elif type(rule) == FilterIntegerRule:
            if self.rule_param_name == 'Workset' and rule.RuleValue in self.workset_names:
                rule_value = self.workset_names[rule.RuleValue]
            else:
                rule_value = rule.RuleValue
        elif type(rule) == FilterDoubleRule:
            rule_value = rule.RuleValue
        elif type(rule) == FilterElementIdRule:
            element_id = rule.RuleValue
            element = self.doc.GetElement(element_id)
            try:
                rule_value = element.Name if element else str(element_id)
            except:
                try:
                    rule_value = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                except:
                    rule_value = str(element_id)
        elif type(rule) in [HasNoValueFilterRule, HasValueFilterRule, SharedParameterApplicableRule]:
            rule_value = '-'
        else:
            rule_value = 'Unknown'
        
        self.rule_value = rule_value
        self.rule_eval = rule_evaluator


class ViewFilter:
    """Parse and store filter information."""
    
    def __init__(self, doc, pfe):
        self.doc = doc
        self.pfe = pfe
        self.cats = self.get_categories()
        self.cat_names = [cat.Name for cat in self.cats]
        self.rules = self.get_rules()
    
    def get_categories(self):
        """Get filter's categories."""
        cats = []
        for catid in self.pfe.GetCategories():
            cat = Category.GetCategory(self.doc, catid)
            if cat:
                cats.append(cat)
        return cats
    
    def get_rules(self):
        """Properly handle different filter types."""
        element_filter = self.pfe.GetElementFilter()
        if not element_filter:
            return []
        
        parsed_rules = []
        
        # Handle LogicalAndFilter or LogicalOrFilter
        if isinstance(element_filter, (LogicalAndFilter, LogicalOrFilter)):
            sub_filters = element_filter.GetFilters()
            for sub_filter in sub_filters:
                if isinstance(sub_filter, ElementParameterFilter):
                    rules = sub_filter.GetRules()
                    for rule in rules:
                        parsed_rule = FilterRule(self.doc, rule)
                        parsed_rules.append(parsed_rule)
        
        # Handle ElementParameterFilter directly
        elif isinstance(element_filter, ElementParameterFilter):
            rules = element_filter.GetRules()
            for rule in rules:
                parsed_rule = FilterRule(self.doc, rule)
                parsed_rules.append(parsed_rule)
        
        return parsed_rules


# ================================================================================================
# MAIN LEGEND GENERATION
# ================================================================================================

def create_legend_view(doc, existing_legends, name="Legend"):
    """Create a new legend view by duplicating an existing one."""
    if not existing_legends:
        raise Exception("No legend views available to duplicate")
    
    random_legend = existing_legends[0]
    new_legend_view_id = random_legend.Duplicate(ViewDuplicateOption.Duplicate)
    new_legend_view = doc.GetElement(new_legend_view_id)
    new_legend_view.Scale = 100
    
    # Rename with fallback
    for i in range(50):
        try:
            new_legend_view.Name = name
            break
        except:
            name += "*"
    
    return new_legend_view


def generate_filter_legends(doc, views_with_filters, text_type, 
                            region_width_mm=None, region_height_mm=None,
                            region_spacing_mm=None, line_width_mm=None):
    """
    Generate legends for views with filters applied
    
    Args:
        doc: Revit document
        views_with_filters: List of views that have filters
        text_type: TextNoteType for legend text
        region_width_mm: Width of filled regions (default from config)
        region_height_mm: Height of filled regions (default from config)
        region_spacing_mm: Spacing between elements (default from config)
        line_width_mm: Width of detail lines (default from config)
    
    Returns:
        List of created legend views
    """
    
    # Use defaults if not specified
    if region_width_mm is None:
        region_width_mm = LegendConfig.REGION_WIDTH_MM
    if region_height_mm is None:
        region_height_mm = LegendConfig.REGION_HEIGHT_MM
    if region_spacing_mm is None:
        region_spacing_mm = LegendConfig.SPACING_MM
    if line_width_mm is None:
        line_width_mm = LegendConfig.LINE_WIDTH_MM
    
    # Convert dimensions to feet
    region_width = convert_mm_to_feet(region_width_mm)
    region_height = convert_mm_to_feet(region_height_mm)
    region_spacing = convert_mm_to_feet(region_spacing_mm)
    line_width = convert_mm_to_feet(line_width_mm)
    text_offset = region_width + region_spacing
    
    # Get existing legends for duplication
    all_views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).ToElements()
    existing_legends = [v for v in all_views if v.ViewType == ViewType.Legend and not v.IsTemplate]
    
    if not existing_legends:
        raise Exception("No legend views exist in project - create at least one legend view first")
    
    created_legends = []
    
    for view in views_with_filters:
        X = 0.0
        Y = 0.0
        
        # Create legend
        legend_name = 'Legend_{}'.format(view.Name)
        legend_view = create_legend_view(doc, existing_legends, legend_name)
        
        # Headers
        create_text_note(doc, legend_view, X, Y + text_offset, 'Surface:', text_type)
        create_text_note(doc, legend_view, X + text_offset, Y + text_offset, 'Cut:', text_type)
        create_text_note(doc, legend_view, X + text_offset * 2, Y + text_offset, 'Surface:', text_type)
        create_text_note(doc, legend_view, X + text_offset * 3, Y + text_offset, 'Cut:', text_type)
        create_text_note(doc, legend_view, X + text_offset * 4, Y + text_offset, 'FilterName:', text_type)
        create_text_note(doc, legend_view, X + text_offset * 6, Y + text_offset, 'Categories:', text_type)
        create_text_note(doc, legend_view, X + text_offset * 8, Y + text_offset, 'Parameter:', text_type)
        create_text_note(doc, legend_view, X + text_offset * 11, Y + text_offset, 'Evaluator:', text_type)
        create_text_note(doc, legend_view, X + text_offset * 13, Y + text_offset, 'Value:', text_type)
        
        # Get and sort filters
        view_filters_ids = view.GetFilters()
        view_filters = [doc.GetElement(id) for id in view_filters_ids]
        view_filters.sort(key=lambda x: x.Name)
        
        # Process each filter
        for filter_elem in view_filters:
            X = 0.0
            
            try:
                # Create lines
                line_surf = create_horizontal_line(doc, legend_view, X, Y - region_height / 2, line_width)
                X += text_offset
                line_cut = create_horizontal_line(doc, legend_view, X, Y - region_height / 2, line_width)
                X += text_offset
                
                # Get filter overrides
                filter_overrides = view.GetFilterOverrides(filter_elem.Id)
                
                # Surface region
                region_surface = create_region(doc, legend_view, X, Y, region_width, region_height)
                override_graphics_region(doc, legend_view, region_surface,
                                         fg_pattern_id=filter_overrides.SurfaceForegroundPatternId,
                                         fg_color=filter_overrides.SurfaceForegroundPatternColor,
                                         bg_pattern_id=filter_overrides.SurfaceBackgroundPatternId,
                                         bg_color=filter_overrides.SurfaceBackgroundPatternColor,
                                         line_color=filter_overrides.ProjectionLineColor,
                                         line_pattern_id=filter_overrides.ProjectionLinePatternId,
                                         lineweight=filter_overrides.ProjectionLineWeight)
                
                override_graphics_line(doc, legend_view, line_surf,
                                       line_color=filter_overrides.ProjectionLineColor,
                                       line_pattern_id=filter_overrides.ProjectionLinePatternId,
                                       lineweight=filter_overrides.ProjectionLineWeight)
                X += text_offset
                
                # Cut region
                region_cut = create_region(doc, legend_view, X, Y, region_width, region_height)
                override_graphics_region(doc, legend_view, region_cut,
                                         fg_pattern_id=filter_overrides.CutForegroundPatternId,
                                         fg_color=filter_overrides.CutForegroundPatternColor,
                                         bg_pattern_id=filter_overrides.CutBackgroundPatternId,
                                         bg_color=filter_overrides.CutBackgroundPatternColor,
                                         line_color=filter_overrides.CutLineColor,
                                         line_pattern_id=filter_overrides.CutLinePatternId,
                                         lineweight=filter_overrides.CutLineWeight)
                
                override_graphics_line(doc, legend_view, line_cut,
                                       line_color=filter_overrides.CutLineColor,
                                       line_pattern_id=filter_overrides.CutLinePatternId,
                                       lineweight=filter_overrides.CutLineWeight)
                X += text_offset
                
                # Filter name
                create_text_note(doc, legend_view, X, Y, filter_elem.Name, text_type)
                X += text_offset * 2
                
                # Parse filter
                my_filter = ViewFilter(doc, filter_elem)
                
                # Categories
                categories = ', '.join(my_filter.cat_names)
                create_text_note(doc, legend_view, X, Y, categories, text_type)
                X += text_offset * 2
                
                # Rules
                rules = my_filter.rules
                if len(rules) == 1:
                    rule = rules[0]
                    rule_param_name = rule.rule_param_name
                    rule_eval = rule.rule_eval
                    rule_value = str(rule.rule_value)
                elif len(rules) == 0:
                    rule_param_name = '[No-Rules]'
                    rule_eval = '[No-Rules]'
                    rule_value = '[No-Rules]'
                else:
                    rule_param_name = '[Multi-Rules]'
                    rule_eval = '[Multi-Rules]'
                    rule_value = '[Multi-Rules]'
                
                create_text_note(doc, legend_view, X, Y, rule_param_name, text_type)
                X += text_offset * 3
                create_text_note(doc, legend_view, X, Y, rule_eval, text_type)
                X += text_offset * 2
                create_text_note(doc, legend_view, X, Y, rule_value, text_type)
                
            except Exception as e:
                print("Error processing filter {}: {}".format(filter_elem.Name, str(e)))
            
            Y -= region_spacing + region_height
        
        created_legends.append(legend_view)
    
    return created_legends