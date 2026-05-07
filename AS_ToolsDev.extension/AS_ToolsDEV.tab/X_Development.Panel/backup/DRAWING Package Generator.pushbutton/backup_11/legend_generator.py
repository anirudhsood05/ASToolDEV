# -*- coding: utf-8 -*-
"""
LEGEND GENERATOR MODULE - SIMPLIFIED ENHANCED VERSION
======================================================
Simplified two-column legend: Pattern (stacked projection/cut) | Value
Expands multi-rule filters to show all rules vertically

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
    """Default legend dimensions (in millimeters at 1:1 scale)"""
    # These are PAPER sizes - will be scaled by legend view scale (typically 1:100)
    REGION_WIDTH_MM = 10.0         # 10mm on paper = 1000mm (1m) in model at 1:100
    REGION_HEIGHT_MM = 10.0        # 10mm on paper = 1000mm (1m) in model at 1:100
    HALF_REGION_HEIGHT_MM = 5.0    # Half height for split projection/cut
    H_SPACING_MM = 15.0            # Horizontal spacing between pattern and text
    V_SPACING_MM = 12.5            # Vertical spacing between filter entries
    
    # AUK Standard text style
    TEXT_STYLE_NAME = 'AUK_Arial Bold_T_2.5mm'


# ================================================================================================
# HELPER FUNCTIONS
# ================================================================================================

def convert_mm_to_feet(mm):
    """Convert millimeters to feet."""
    return mm / 304.8


def get_text_type_by_name(doc, text_type_name):
    """Get TextNoteType by name with fallback"""
    all_text_types = FilteredElementCollector(doc)\
        .OfClass(TextNoteType)\
        .WhereElementIsElementType()\
        .ToElements()
    
    # Try exact match first
    for text_type in all_text_types:
        type_name = text_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        if type_name == text_type_name:
            return text_type
    
    # Fallback to first available
    if all_text_types:
        return all_text_types[0]
    
    return None


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


def override_graphics_region(doc, view, region, fg_pattern_id=None, fg_color=None, 
                              bg_pattern_id=None, bg_color=None):
    """Override ONLY fill pattern graphics for a filled region (no lines)."""
    settings = OverrideGraphicSettings()
    
    # Surface patterns only
    if fg_pattern_id and fg_pattern_id != ElementId.InvalidElementId:
        settings.SetSurfaceForegroundPatternId(fg_pattern_id)
    if fg_color:
        settings.SetSurfaceForegroundPatternColor(fg_color)
    if bg_pattern_id and bg_pattern_id != ElementId.InvalidElementId:
        settings.SetSurfaceBackgroundPatternId(bg_pattern_id)
    if bg_color:
        settings.SetSurfaceBackgroundPatternColor(bg_color)
    
    view.SetElementOverrides(region.Id, settings)


def get_workset_names(doc):
    """Get workset names dictionary"""
    try:
        all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
        return {int(str(w.Id)): w.Name for w in all_worksets}
    except:
        return {}


def patterns_are_identical(proj_fg_id, proj_bg_id, proj_fg_color, proj_bg_color,
                           cut_fg_id, cut_bg_id, cut_fg_color, cut_bg_color):
    """Check if projection and cut patterns are identical"""
    return (proj_fg_id == cut_fg_id and 
            proj_bg_id == cut_bg_id and
            proj_fg_color.Red == cut_fg_color.Red and
            proj_fg_color.Green == cut_fg_color.Green and
            proj_fg_color.Blue == cut_fg_color.Blue and
            proj_bg_color.Red == cut_bg_color.Red and
            proj_bg_color.Green == cut_bg_color.Green and
            proj_bg_color.Blue == cut_bg_color.Blue)


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
        self.rule_param_name = None
        self.get_rule_info()
    
    def get_parameter_name(self):
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
    
    def get_rule_info(self):
        """Extract rule evaluator, parameter name, and value."""
        rule_evaluator = ''
        rule_value = ''
        inverse = False
        
        # Handle inverse rules
        if type(self.rule) == FilterInverseRule:
            inverse = True
            rule = self.rule.GetInnerRule()
        else:
            rule = self.rule
        
        # Get parameter name
        self.rule_param_name = self.get_parameter_name()
        
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
    
    def format_display_text(self):
        """Format rule as display text: 'Parameter: Evaluator Value'"""
        return "{}: {} {}".format(self.rule_param_name, self.rule_eval, self.rule_value)


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
# MAIN LEGEND GENERATION - SIMPLIFIED TWO-COLUMN
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


def generate_filter_legends(doc, views_with_filters, text_type=None, 
                            region_width_mm=None, region_height_mm=None,
                            h_spacing_mm=None, v_spacing_mm=None):
    """
    Generate simplified two-column legends: Pattern | Value
    
    Args:
        doc: Revit document
        views_with_filters: List of views that have filters
        text_type: TextNoteType for legend text (optional - will use AUK standard)
        region_width_mm: Width of filled regions in mm (paper size, default from config)
        region_height_mm: Full height for combined patterns in mm (default from config)
        h_spacing_mm: Horizontal spacing between columns in mm (default from config)
        v_spacing_mm: Vertical spacing between filter entries in mm (default from config)
    
    Returns:
        List of created legend views
    """
    
    # Use AUK standard text type if not provided
    if text_type is None:
        text_type = get_text_type_by_name(doc, LegendConfig.TEXT_STYLE_NAME)
        if text_type is None:
            raise Exception("Text type '{}' not found and no fallback available".format(
                LegendConfig.TEXT_STYLE_NAME))
    
    # Use defaults if not specified
    if region_width_mm is None:
        region_width_mm = LegendConfig.REGION_WIDTH_MM
    if region_height_mm is None:
        region_height_mm = LegendConfig.REGION_HEIGHT_MM
    if h_spacing_mm is None:
        h_spacing_mm = LegendConfig.H_SPACING_MM
    if v_spacing_mm is None:
        v_spacing_mm = LegendConfig.V_SPACING_MM
    
    # Convert dimensions to feet
    region_width = convert_mm_to_feet(region_width_mm)
    region_height = convert_mm_to_feet(region_height_mm)
    half_region_height = convert_mm_to_feet(LegendConfig.HALF_REGION_HEIGHT_MM)
    h_spacing = convert_mm_to_feet(h_spacing_mm)
    v_spacing = convert_mm_to_feet(v_spacing_mm)
    
    # Column positions
    pattern_col_x = 0.0
    value_col_x = pattern_col_x + region_width + h_spacing
    
    # Get existing legends for duplication
    all_views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).ToElements()
    existing_legends = [v for v in all_views if v.ViewType == ViewType.Legend and not v.IsTemplate]
    
    if not existing_legends:
        raise Exception("No legend views exist in project - create at least one legend view first")
    
    created_legends = []
    
    for view in views_with_filters:
        current_y = 0.0
        
        # Create legend
        legend_name = 'Legend_{}'.format(view.Name)
        legend_view = create_legend_view(doc, existing_legends, legend_name)
        
        # HEADERS - Simplified two-column
        create_text_note(doc, legend_view, pattern_col_x, current_y, 'Pattern', text_type)
        create_text_note(doc, legend_view, value_col_x, current_y, 'Value', text_type)
        
        current_y -= v_spacing  # Move down for first filter entry
        
        # Get and sort filters
        view_filters_ids = view.GetFilters()
        view_filters = [doc.GetElement(id) for id in view_filters_ids]
        view_filters.sort(key=lambda x: x.Name)
        
        # Process each filter
        for filter_elem in view_filters:
            try:
                # Get filter overrides
                filter_overrides = view.GetFilterOverrides(filter_elem.Id)
                
                # Get pattern IDs and colors
                proj_fg_id = filter_overrides.SurfaceForegroundPatternId
                proj_bg_id = filter_overrides.SurfaceBackgroundPatternId
                proj_fg_color = filter_overrides.SurfaceForegroundPatternColor
                proj_bg_color = filter_overrides.SurfaceBackgroundPatternColor
                
                cut_fg_id = filter_overrides.CutForegroundPatternId
                cut_bg_id = filter_overrides.CutBackgroundPatternId
                cut_fg_color = filter_overrides.CutForegroundPatternColor
                cut_bg_color = filter_overrides.CutBackgroundPatternColor
                
                # Check if patterns are identical
                patterns_identical = patterns_are_identical(
                    proj_fg_id, proj_bg_id, proj_fg_color, proj_bg_color,
                    cut_fg_id, cut_bg_id, cut_fg_color, cut_bg_color
                )
                
                # === PATTERN COLUMN ===
                if patterns_identical:
                    # Single region showing combined pattern
                    region = create_region(doc, legend_view, pattern_col_x, current_y, 
                                          region_width, region_height)
                    override_graphics_region(doc, legend_view, region,
                                            fg_pattern_id=proj_fg_id,
                                            fg_color=proj_fg_color,
                                            bg_pattern_id=proj_bg_id,
                                            bg_color=proj_bg_color)
                else:
                    # Two stacked regions - Projection (top) and Cut (bottom)
                    # Projection region (top half)
                    region_proj_y = current_y + (half_region_height / 2.0)
                    region_proj = create_region(doc, legend_view, pattern_col_x, region_proj_y,
                                               region_width, half_region_height)
                    override_graphics_region(doc, legend_view, region_proj,
                                            fg_pattern_id=proj_fg_id,
                                            fg_color=proj_fg_color,
                                            bg_pattern_id=proj_bg_id,
                                            bg_color=proj_bg_color)
                    
                    # Cut region (bottom half)
                    region_cut_y = current_y - (half_region_height / 2.0)
                    region_cut = create_region(doc, legend_view, pattern_col_x, region_cut_y,
                                              region_width, half_region_height)
                    override_graphics_region(doc, legend_view, region_cut,
                                            fg_pattern_id=cut_fg_id,
                                            fg_color=cut_fg_color,
                                            bg_pattern_id=cut_bg_id,
                                            bg_color=cut_bg_color)
                
                # === VALUE COLUMN ===
                # Parse filter
                my_filter = ViewFilter(doc, filter_elem)
                
                # Get rules
                rules = my_filter.rules
                
                if len(rules) == 0:
                    # No rules - just show filter name
                    value_text = "{} [No Rules]".format(filter_elem.Name)
                    create_text_note(doc, legend_view, value_col_x, current_y, value_text, text_type)
                
                elif len(rules) == 1:
                    # Single rule - show parameter: evaluator value
                    rule = rules[0]
                    value_text = rule.format_display_text()
                    create_text_note(doc, legend_view, value_col_x, current_y, value_text, text_type)
                
                else:
                    # Multiple rules - expand vertically
                    # First line: filter name
                    create_text_note(doc, legend_view, value_col_x, current_y, 
                                    filter_elem.Name, text_type)
                    
                    # Subsequent lines: each rule indented
                    rule_y = current_y
                    for rule in rules:
                        rule_y -= (v_spacing * 0.5)  # Half spacing for sub-rules
                        rule_text = "  - {}".format(rule.format_display_text())
                        create_text_note(doc, legend_view, value_col_x, rule_y, 
                                        rule_text, text_type)
                    
                    # Adjust current_y to account for expanded rules
                    # Already moved down by (len(rules) - 1) * (v_spacing * 0.5)
                    # So we need to compensate by moving up slightly for next filter
                    current_y = rule_y
                
            except Exception as e:
                print("Error processing filter {}: {}".format(filter_elem.Name, str(e)))
            
            # Move down for next filter entry
            current_y -= v_spacing
        
        created_legends.append(legend_view)
    
    return created_legends
