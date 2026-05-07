# -*- coding: utf-8 -*-
"""
Script:   Family Visibility Type Creator
Desc:     Creates Yes/No visibility parameters for geometry elements in the
          Family Editor, associates them to geometry Visible properties, and
          generates one family type per parameter combination.
Author:   Anirudh Sood / Aukett Swanke
Usage:    Open a family (.rfa) in Family Editor, run this tool.
Result:   Yes/No type parameters created, associated to geometry Visible
          properties, and new family types generated.
"""
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms
from pyrevit.framework import wpf

import clr
clr.AddReference("System")

import System
from System.Windows import Controls, Media, Thickness
from System.Windows.Controls import CheckBox, TextBlock, StackPanel
from System.IO import StringReader

import traceback

logger = script.get_logger()
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc
app = __revit__.Application

# ── Constants ────────────────────────────────────────────────────────────────
REVIT_VERSION = int(app.VersionNumber)
MAX_GEOMETRY = 50

# AUK Colours
AUK_BLUE = "#375D6D"
AUK_GOLD = "#E8A735"
AUK_GREY = "#CCCCCC"
AUK_BG = "#F5F5F5"
AUK_TEXT = "#2C2C2C"


# ── Version-Safe Helpers ─────────────────────────────────────────────────────

def eid_int(element_id):
    """Return integer value of ElementId. Revit 2023-2026 safe."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def add_yesno_type_parameter(fam_mgr, param_name):
    """Add a Yes/No TYPE parameter. Version-safe for Revit 2023-2026.
    
    Revit 2025.3+ removed the (string, BuiltInParameterGroup, ParameterType, bool)
    overload. Uses ForgeTypeId-based overload for 2025+.
    """
    # Reuse if already exists
    for fp in fam_mgr.GetParameters():
        if DB.Element.Name.GetValue(fp) == param_name:
            logger.info("Parameter '{}' already exists, reusing.".format(param_name))
            return fp

    if REVIT_VERSION >= 2025:
        # ForgeTypeId API (Revit 2025+)
        try:
            param = fam_mgr.AddParameter(
                param_name,
                DB.SpecTypeId.Boolean.YesNo,
                DB.GroupTypeId.Visibility,
                True  # isType
            )
            return param
        except Exception:
            # Fallback to Graphics group
            param = fam_mgr.AddParameter(
                param_name,
                DB.SpecTypeId.Boolean.YesNo,
                DB.GroupTypeId.Graphics,
                True
            )
            return param
    else:
        # Legacy API (Revit 2023-2024)
        try:
            param = fam_mgr.AddParameter(
                param_name,
                DB.BuiltInParameterGroup.PG_VISIBILITY,
                DB.ParameterType.YesNo,
                True
            )
            return param
        except Exception:
            param = fam_mgr.AddParameter(
                param_name,
                DB.BuiltInParameterGroup.PG_GRAPHICS,
                DB.ParameterType.YesNo,
                True
            )
            return param


def get_geometry_elements(fam_doc):
    """Collect elements whose Visible property can be controlled.
    Works for model families (Extrusions, Blends, Sweeps) AND
    annotation families (Detail Lines, Filled Regions, Text, Masking Regions, etc.)."""
    geo_elements = []
    seen_ids = set()

    def _try_add(elem):
        """Add element if it has GEOM_VISIBILITY_PARAM and not already seen."""
        eid = eid_int(elem.Id)
        if eid in seen_ids:
            return
        try:
            vis = elem.get_Parameter(DB.BuiltInParameter.GEOM_VISIBILITY_PARAM)
            if vis is not None:
                geo_elements.append(elem)
                seen_ids.add(eid)
        except Exception:
            pass

    # GenericForm: Extrusion, Blend, Sweep, Revolution, SweptBlend
    for elem in DB.FilteredElementCollector(fam_doc)\
            .OfClass(DB.GenericForm).ToElements():
        _try_add(elem)

    # Nested FamilyInstance elements
    for elem in DB.FilteredElementCollector(fam_doc)\
            .OfClass(DB.FamilyInstance).ToElements():
        _try_add(elem)

    # CurveElement: Detail Lines, Symbolic Lines, Model Lines
    for elem in DB.FilteredElementCollector(fam_doc)\
            .OfClass(DB.CurveElement).ToElements():
        _try_add(elem)

    # FilledRegion: Filled/Masking Regions
    try:
        for elem in DB.FilteredElementCollector(fam_doc)\
                .OfClass(DB.FilledRegion).ToElements():
            _try_add(elem)
    except Exception:
        pass

    # TextElement: Text notes in annotation families
    try:
        for elem in DB.FilteredElementCollector(fam_doc)\
                .OfClass(DB.TextElement).ToElements():
            _try_add(elem)
    except Exception:
        pass

    # Fallback: broad scan for anything else with GEOM_VISIBILITY_PARAM
    # (catches ImportInstance, ModelText, and other edge cases)
    try:
        for elem in DB.FilteredElementCollector(fam_doc)\
                .WhereElementIsNotElementType().ToElements():
            if eid_int(elem.Id) not in seen_ids:
                _try_add(elem)
    except Exception:
        pass

    return geo_elements


def get_element_label(elem):
    """Human-readable label for an element."""
    eid = eid_int(elem.Id)
    parts = []

    # Element class hint
    cls = elem.GetType().Name
    friendly_map = {
        "Extrusion": "Extrusion",
        "Blend": "Blend",
        "Sweep": "Sweep",
        "Revolution": "Revolve",
        "SweptBlend": "SweptBlend",
        "FamilyInstance": "Nested",
        "DetailLine": "DetailLine",
        "DetailArc": "DetailArc",
        "DetailEllipse": "DetailEllipse",
        "DetailNurbSpline": "DetailSpline",
        "ModelLine": "ModelLine",
        "ModelArc": "ModelArc",
        "SymbolicCurve": "SymbolicLine",
        "FilledRegion": "FilledRegion",
        "TextNote": "Text",
        "TextElement": "Text",
        "ImportInstance": "Import",
    }
    class_label = friendly_map.get(cls, cls)
    parts.append(class_label)

    # Element name
    try:
        name = DB.Element.Name.GetValue(elem)
        if name and name != class_label:
            parts.append(name)
    except Exception:
        pass

    # Category
    try:
        if elem.Category and elem.Category.Name not in parts:
            parts.append(elem.Category.Name)
    except Exception:
        pass

    label = " : ".join(parts)
    return "{} [{}]".format(label, eid)


def sanitise_param_name(label):
    """Convert element label to a valid parameter name."""
    raw = label.split("[")[0].strip()
    raw = raw.replace(" ", "_").replace(".", "_")
    raw = "".join(c for c in raw if c.isalnum() or c == "_")
    if not raw or raw[0].isdigit():
        raw = "Geo_" + raw
    return raw


# ── XAML ─────────────────────────────────────────────────────────────────────
XAML_STR = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Family Visibility Type Creator"
    Width="720" Height="640"
    MinWidth="600" MinHeight="500"
    WindowStartupLocation="CenterScreen"
    ShowInTaskbar="False"
    ResizeMode="CanResize"
    FontFamily="Arial" FontSize="12"
    Background="#F5F5F5">

    <Window.Resources>
        <Style x:Key="HeaderStyle" TargetType="Border">
            <Setter Property="Background" Value="#375D6D"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="CornerRadius" Value="3"/>
        </Style>
        <Style x:Key="CardStyle" TargetType="Border">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="3"/>
            <Setter Property="Padding" Value="10"/>
        </Style>
        <Style x:Key="BtnPrimary" TargetType="Button">
            <Setter Property="Background" Value="#375D6D"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="MinWidth" Value="90"/>
            <Setter Property="Margin" Value="4,0"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="BtnGold" TargetType="Button">
            <Setter Property="Background" Value="#E8A735"/>
            <Setter Property="Foreground" Value="#2C2C2C"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="MinWidth" Value="90"/>
            <Setter Property="Margin" Value="4,0"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="BtnCancel" TargetType="Button">
            <Setter Property="Background" Value="#CCCCCC"/>
            <Setter Property="Foreground" Value="#2C2C2C"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="MinWidth" Value="90"/>
            <Setter Property="Margin" Value="4,0"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Style="{StaticResource HeaderStyle}">
            <StackPanel>
                <TextBlock Text="Family Visibility Type Creator"
                           FontWeight="Bold" FontSize="16"
                           Foreground="White" HorizontalAlignment="Center"/>
                <TextBlock Text="Creates one type per geometry element with visibility ON/OFF control"
                           FontSize="11" Foreground="#B0C4CE" HorizontalAlignment="Center"
                           Margin="0,2,0,0"/>
            </StackPanel>
        </Border>

        <!-- Info -->
        <Border Grid.Row="2" Background="#EEF4F7" Padding="10" CornerRadius="3"
                BorderBrush="#B0C4CE" BorderThickness="1">
            <TextBlock TextWrapping="Wrap" Foreground="#2C2C2C" FontSize="11">
                <Run FontWeight="SemiBold">How it works:</Run>
                <LineBreak/>
                For each ticked geometry element below, the tool will:
                <LineBreak/>
                1. Create a Yes/No type parameter (e.g. Show_ExtrusionA)
                <LineBreak/>
                2. Associate it to the element's Visible property
                <LineBreak/>
                3. Create a new family type where that element is ON (others OFF)
            </TextBlock>
        </Border>

        <!-- Geometry list with editable names -->
        <Border Grid.Row="4" Style="{StaticResource CardStyle}">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Text="Geometry Elements"
                           FontWeight="SemiBold" FontSize="12"
                           Foreground="#2C2C2C" Margin="0,0,0,2"/>
                <TextBlock Grid.Row="1" Foreground="#666666" FontSize="11"
                           Margin="0,0,0,6"
                           Text="Tick to include. Edit the Type Name and Parameter Name columns."/>

                <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Auto">
                    <StackPanel x:Name="UI_geo_panel"/>
                </ScrollViewer>

                <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,6,0,0">
                    <Button x:Name="UI_select_all" Content="Select All"
                            Style="{StaticResource BtnPrimary}" Height="24" MinWidth="75"/>
                    <Button x:Name="UI_select_none" Content="Deselect All"
                            Style="{StaticResource BtnCancel}" Height="24" MinWidth="75"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Options -->
        <Border Grid.Row="6" Style="{StaticResource CardStyle}">
            <StackPanel>
                <CheckBox x:Name="UI_all_on_type"
                          Content="Also create an 'All Visible' type (all parameters ON)"
                          Foreground="#2C2C2C" IsChecked="True"/>
                <CheckBox x:Name="UI_skip_types"
                          Content="Create parameters and associations only (skip type creation)"
                          Foreground="#2C2C2C" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Buttons -->
        <StackPanel Grid.Row="8" Orientation="Horizontal"
                    HorizontalAlignment="Right">
            <Button x:Name="UI_generate" Content="Generate"
                    Style="{StaticResource BtnGold}" MinWidth="120"/>
            <Button x:Name="UI_close" Content="Close"
                    Style="{StaticResource BtnCancel}"/>
        </StackPanel>
    </Grid>
</Window>
"""


# ── Data Class ───────────────────────────────────────────────────────────────

class GeoRow(object):
    """One geometry element row in the UI."""
    def __init__(self, element, label, param_name, type_name):
        self.element = element
        self.label = label
        self.param_name = param_name
        self.type_name = type_name
        self.include = True
        self.cb = None
        self.tb_param = None
        self.tb_type = None


# ── Main Window ──────────────────────────────────────────────────────────────

class VisTypeWindow(forms.WPFWindow):
    def __init__(self):
        xaml_stream = StringReader(XAML_STR)
        wpf.LoadComponent(self, xaml_stream)

        self.geo_rows = []
        self._result = False
        self._init_ok = False
        self._active_rows = []
        self._skip_types = False
        self._all_on = True

        self.initialize()
        if self._init_ok:
            self.connect_events()

    def initialize(self):
        """Load geometry elements and build editable rows."""
        geo_elements = get_geometry_elements(doc)

        if not geo_elements:
            forms.alert(
                "No geometry elements found in this family.\n\n"
                "Ensure the family contains Extrusions, Blends, Sweeps, "
                "Filled Regions, Detail Lines, Text elements, "
                "or nested family instances with a Visible parameter.",
                title="No Geometry Found"
            )
            return

        if len(geo_elements) > MAX_GEOMETRY:
            geo_elements = geo_elements[:MAX_GEOMETRY]
            logger.warning("Capped at {} geometry elements.".format(MAX_GEOMETRY))

        # Build rows with default names
        used_params = set()
        used_types = set()
        for elem in geo_elements:
            label = get_element_label(elem)
            base = sanitise_param_name(label)

            pname = "Show_{}".format(base)
            counter = 1
            while pname in used_params:
                pname = "Show_{}_{}".format(base, counter)
                counter += 1
            used_params.add(pname)

            tname = "Type_{}".format(base)
            counter = 1
            while tname in used_types:
                tname = "Type_{}_{}".format(base, counter)
                counter += 1
            used_types.add(tname)

            self.geo_rows.append(GeoRow(elem, label, pname, tname))

        self._build_geo_panel()
        self._init_ok = True

    def connect_events(self):
        self.UI_select_all.Click += self._on_select_all
        self.UI_select_none.Click += self._on_select_none
        self.UI_generate.Click += self._on_generate
        self.UI_close.Click += self._on_close

    # ── Build UI rows ────────────────────────────────────────────────────

    def _build_geo_panel(self):
        """Build one editable row per geometry element."""
        panel = self.UI_geo_panel
        panel.Children.Clear()
        brush = Media.BrushConverter()

        # Header row
        hdr = StackPanel()
        hdr.Orientation = Controls.Orientation.Horizontal
        hdr.Margin = Thickness(0, 0, 0, 4)

        hdr_cb = TextBlock()
        hdr_cb.Text = ""
        hdr_cb.Width = 24
        hdr.Children.Add(hdr_cb)

        for col_text, col_w in [("Geometry Element", 200),
                                ("Parameter Name", 180),
                                ("Type Name", 180)]:
            tb = TextBlock()
            tb.Text = col_text
            tb.Width = col_w
            tb.FontWeight = System.Windows.FontWeights.SemiBold
            tb.Foreground = brush.ConvertFromString(AUK_TEXT)
            hdr.Children.Add(tb)

        panel.Children.Add(hdr)

        sep = Controls.Separator()
        sep.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(sep)

        # Data rows
        for row in self.geo_rows:
            sp = StackPanel()
            sp.Orientation = Controls.Orientation.Horizontal
            sp.Margin = Thickness(0, 2, 0, 2)

            cb = CheckBox()
            cb.IsChecked = True
            cb.Margin = Thickness(0, 0, 4, 0)
            cb.VerticalAlignment = System.Windows.VerticalAlignment.Center
            row.cb = cb
            sp.Children.Add(cb)

            lbl = TextBlock()
            short = row.label if len(row.label) <= 28 else row.label[:25] + "..."
            lbl.Text = short
            lbl.Width = 200
            lbl.ToolTip = row.label
            lbl.Foreground = brush.ConvertFromString(AUK_TEXT)
            lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            sp.Children.Add(lbl)

            tb_p = Controls.TextBox()
            tb_p.Text = row.param_name
            tb_p.Width = 176
            tb_p.Height = 22
            tb_p.Margin = Thickness(0, 0, 4, 0)
            tb_p.Background = brush.ConvertFromString("#FFFFFF")
            tb_p.BorderBrush = brush.ConvertFromString("#CCCCCC")
            tb_p.Foreground = brush.ConvertFromString(AUK_TEXT)
            tb_p.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
            row.tb_param = tb_p
            sp.Children.Add(tb_p)

            tb_t = Controls.TextBox()
            tb_t.Text = row.type_name
            tb_t.Width = 176
            tb_t.Height = 22
            tb_t.Background = brush.ConvertFromString("#FFFFFF")
            tb_t.BorderBrush = brush.ConvertFromString("#CCCCCC")
            tb_t.Foreground = brush.ConvertFromString(AUK_TEXT)
            tb_t.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
            row.tb_type = tb_t
            sp.Children.Add(tb_t)

            panel.Children.Add(sp)

    # ── Button handlers ──────────────────────────────────────────────────

    def _on_select_all(self, sender, args):
        for row in self.geo_rows:
            row.cb.IsChecked = True

    def _on_select_none(self, sender, args):
        for row in self.geo_rows:
            row.cb.IsChecked = False

    def _on_close(self, sender, args):
        self.Close()

    def _on_generate(self, sender, args):
        """Validate, collect data, close window."""
        active_rows = []
        for row in self.geo_rows:
            if not row.cb.IsChecked:
                continue
            row.param_name = row.tb_param.Text.strip()
            row.type_name = row.tb_type.Text.strip()
            active_rows.append(row)

        if not active_rows:
            forms.alert("Select at least one geometry element.",
                        title="Nothing Selected")
            return

        skip_types = self.UI_skip_types.IsChecked

        # Validate param names
        param_names = []
        for r in active_rows:
            if not r.param_name:
                forms.alert(
                    "Parameter name is empty for '{}'.".format(r.label),
                    title="Invalid Name")
                return
            param_names.append(r.param_name)

        if len(param_names) != len(set(param_names)):
            forms.alert("Duplicate parameter names. Make each unique.",
                        title="Duplicate Names")
            return

        # Validate type names
        if not skip_types:
            type_names = []
            for r in active_rows:
                if not r.type_name:
                    forms.alert(
                        "Type name is empty for '{}'.".format(r.label),
                        title="Invalid Name")
                    return
                type_names.append(r.type_name)

            if self.UI_all_on_type.IsChecked:
                type_names.append("All_Visible")

            if len(type_names) != len(set(type_names)):
                forms.alert("Duplicate type names. Make each unique.",
                            title="Duplicate Names")
                return

            # Check against existing family types
            fam_mgr = doc.FamilyManager
            existing = set()
            for ft in fam_mgr.Types:
                existing.add(DB.Element.Name.GetValue(ft))
            for tn in type_names:
                if tn in existing:
                    forms.alert(
                        "Type '{}' already exists in the family.".format(tn),
                        title="Duplicate Type")
                    return

        self._active_rows = active_rows
        self._skip_types = skip_types
        self._all_on = self.UI_all_on_type.IsChecked
        self._result = True
        self.Close()

    @property
    def result(self):
        return self._result


# ── Execution ────────────────────────────────────────────────────────────────

def execute(window):
    """Run Revit transactions with data from the window."""
    rows = window._active_rows
    skip_types = window._skip_types
    create_all_on = window._all_on
    fam_mgr = doc.FamilyManager

    created_params = []  # (GeoRow, FamilyParameter)
    failed = []
    succeeded_params = 0
    succeeded_assoc = 0
    succeeded_types = 0

    try:
        with revit.TransactionGroup("Create Visibility Types"):

            # Phase 1: Create parameters + associate to geometry
            with revit.Transaction("Create Visibility Parameters"):
                for row in rows:
                    try:
                        fp = add_yesno_type_parameter(fam_mgr, row.param_name)
                        created_params.append((row, fp))
                        succeeded_params += 1
                    except Exception as e:
                        failed.append("Param '{}': {}".format(
                            row.param_name, str(e)))
                        created_params.append((row, None))
                        continue

                    try:
                        vis_param = row.element.get_Parameter(
                            DB.BuiltInParameter.GEOM_VISIBILITY_PARAM)
                        if vis_param is None:
                            failed.append(
                                "No Visible param: '{}'".format(row.label))
                            continue
                        fam_mgr.AssociateElementParameterToFamilyParameter(
                            vis_param, fp)
                        succeeded_assoc += 1
                    except Exception as e:
                        failed.append("Assoc '{}': {}".format(
                            row.param_name, str(e)))

            # Phase 2: Create one type per geometry element
            if not skip_types:
                with revit.Transaction("Create Family Types"):
                    for target_row, target_fp in created_params:
                        if target_fp is None:
                            continue
                        try:
                            new_type = fam_mgr.NewType(target_row.type_name)
                            fam_mgr.CurrentType = new_type

                            # This element ON, all others OFF
                            for row, fp in created_params:
                                if fp is None:
                                    continue
                                try:
                                    val = 1 if row is target_row else 0
                                    fam_mgr.Set(fp, val)
                                except Exception as ep:
                                    failed.append("Set {}/{}: {}".format(
                                        target_row.type_name,
                                        row.param_name, str(ep)))

                            succeeded_types += 1
                        except Exception as et:
                            failed.append("Type '{}': {}".format(
                                target_row.type_name, str(et)))

                    # All Visible type
                    if create_all_on:
                        try:
                            all_type = fam_mgr.NewType("All_Visible")
                            fam_mgr.CurrentType = all_type
                            for _, fp in created_params:
                                if fp is None:
                                    continue
                                try:
                                    fam_mgr.Set(fp, 1)
                                except Exception:
                                    pass
                            succeeded_types += 1
                        except Exception as eat:
                            failed.append(
                                "All_Visible type: {}".format(str(eat)))

        # ── Report ───────────────────────────────────────────────────
        output.print_md("## Family Visibility Type Creator - Results")
        output.print_md("**Parameters created:** {} of {}".format(
            succeeded_params, len(rows)))
        output.print_md("**Associations made:** {} of {}".format(
            succeeded_assoc, len(rows)))
        if not skip_types:
            valid = len([r for r, fp in created_params if fp])
            expected = valid + (1 if create_all_on else 0)
            output.print_md("**Types created:** {} of {}".format(
                succeeded_types, expected))

        if failed:
            output.print_md("### Issues")
            for f in failed:
                output.print_md("- {}".format(f))

        if succeeded_params > 0 and not failed:
            output.print_md("*All operations completed successfully.*")

        output.print_md("---")
        output.print_md("*Revit {} | pyRevit*".format(REVIT_VERSION))

    except Exception as e:
        logger.error("Critical error: {}".format(str(e)))
        logger.error(traceback.format_exc())
        forms.alert(
            "Critical error:\n\n{}\n\nCheck the output window.".format(str(e)),
            title="Error")


# ── Validation ───────────────────────────────────────────────────────────────

def validate():
    """Pre-condition checks."""
    if not doc.IsFamilyDocument:
        forms.alert(
            "This tool must be run from the Family Editor.\n\n"
            "Open a family file (.rfa) and try again.",
            title="Not a Family Document",
            exitscript=True)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    validate()
    try:
        window = VisTypeWindow()
        if not window._init_ok:
            script.exit()
        window.ShowDialog()
        if window.result:
            execute(window)
    except Exception as e:
        logger.error("Unexpected error: {}".format(str(e)))
        logger.error(traceback.format_exc())
        forms.alert(
            "Unexpected error:\n\n{}\n\nCheck logs.".format(str(e)),
            title="Error",
            exitscript=True)


main()
