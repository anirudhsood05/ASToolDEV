# -*- coding: utf-8 -*-
__title__ = 'DWG\nConverter'
__doc__ = """Two-mode DWG converter.
Mode A: Batch-converts DWG files into Revit detail item families (.rfa).
Mode B: Converts an imported/linked DWG in the current view into native
Revit lines (Detail, Model, Area Boundary, Room Separation, Space Boundary)
with simple or per-layer linestyle assignment."""
__author__ = "AS Tools"

import os
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows.Forms import OpenFileDialog, FolderBrowserDialog, DialogResult
import System.Windows.Controls as WpfControls
import System.Windows as Wpf

from Autodesk.Revit.DB import (
    Transaction, SaveAsOptions, DWGImportOptions, ImportPlacement,
    FilteredElementCollector, BuiltInCategory, ViewPlan, ViewType,
    IFailuresPreprocessor, FailureProcessingResult, FailureSeverity,
    Options, GeometryInstance, PolyLine, Curve, Line,
    GraphicsStyleType, ViewSchedule, ImportInstance, CurveArray,
    SketchPlane, Plane, XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, forms, script

uidoc  = revit.uidoc
doc    = revit.doc
app    = doc.Application
output = script.get_output()

# ---------------------------------------------------------------------------
# MODE A — DWG → DETAIL FAMILY (.rfa)
# ---------------------------------------------------------------------------

class WarningSwallower(IFailuresPreprocessor):
    """Suppresses DWG import warnings silently."""
    def PreprocessFailures(self, failuresAccessor):
        for msg in list(failuresAccessor.GetFailureMessages()):
            if msg.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(msg)
        return FailureProcessingResult.Continue


def pick_file(title, file_filter, multi=False):
    dlg = OpenFileDialog()
    dlg.Title      = title
    dlg.Filter     = file_filter
    dlg.Multiselect = multi
    if dlg.ShowDialog() == DialogResult.OK:
        return list(dlg.FileNames) if multi else dlg.FileName
    return None


def pick_folder(description):
    dlg = FolderBrowserDialog()
    dlg.Description = description
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.SelectedPath
    return None


def get_ref_level_floor_plan(fam_doc):
    """Return the 'Ref. Level' floor plan view from a family document."""
    collector = FilteredElementCollector(fam_doc).OfCategory(BuiltInCategory.OST_Views)
    for view in collector:
        if (isinstance(view, ViewPlan)
                and not view.IsTemplate
                and view.ViewType == ViewType.FloorPlan
                and view.Name == "Ref. Level"):
            return view
    return None


def create_family_from_dwg(dwg_path, template_path, save_folder):
    """
    Open a new family from template_path, import dwg_path, save as .rfa.
    Returns (save_path, None) on success or (None, error_message) on failure.
    """
    try:
        fam_doc = app.NewFamilyDocument(template_path)
    except Exception as ex:
        return None, "NewFamilyDocument failed: " + str(ex)

    if fam_doc is None:
        return None, "NewFamilyDocument returned None."

    t = Transaction(fam_doc, "Import DWG")
    try:
        fail_opts = t.GetFailureHandlingOptions()
        fail_opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(fail_opts)
        t.Start()

        floor_plan = get_ref_level_floor_plan(fam_doc)
        if floor_plan is None:
            t.RollBack()
            fam_doc.Close(False)
            return None, "No 'Ref. Level' floor plan found in template."

        dwg_opts           = DWGImportOptions()
        dwg_opts.Placement = ImportPlacement.Origin
        fam_doc.Import(dwg_path, dwg_opts, floor_plan)
        t.Commit()

    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        try:
            fam_doc.Close(False)
        except Exception:
            pass
        return None, "Import transaction failed: " + str(ex)

    dwg_filename = os.path.basename(dwg_path)
    rfa_name     = os.path.splitext(dwg_filename)[0] + ".rfa"
    save_path    = os.path.join(save_folder, rfa_name)

    try:
        save_opts = SaveAsOptions()
        save_opts.OverwriteExistingFile = True
        fam_doc.SaveAs(save_path, save_opts)
        fam_doc.Close(False)
        return save_path, None
    except Exception as ex:
        try:
            fam_doc.Close(False)
        except Exception:
            pass
        return None, "Save failed: " + str(ex)


def run_mode_a():
    template_path = pick_file(
        "Select Family Template (.rft)",
        "Family Templates (*.rft)|*.rft"
    )
    if not template_path:
        forms.alert("No template selected. Operation cancelled.", exitscript=True)
        return
    if not os.path.isfile(template_path):
        forms.alert("Template file not found:\n" + template_path, exitscript=True)
        return

    dwg_files = pick_file(
        "Select DWG Files to Convert",
        "DWG Files (*.dwg)|*.dwg",
        multi=True
    )
    if not dwg_files:
        forms.alert("No DWG files selected. Operation cancelled.", exitscript=True)
        return

    save_folder = pick_folder("Select Folder to Save RFA Families")
    if not save_folder:
        forms.alert("No save folder selected. Operation cancelled.", exitscript=True)
        return

    created = 0
    failed  = []

    for dwg_path in dwg_files:
        save_path, err = create_family_from_dwg(dwg_path, template_path, save_folder)
        name = os.path.basename(dwg_path)
        if save_path:
            created += 1
            output.print_md("**Created**: {}".format(os.path.basename(save_path)))
        else:
            failed.append("{} ({})".format(name, err or "Unknown error"))
            output.print_md("**Failed**: {} - {}".format(name, err or "Unknown error"))

    msg = "Created {} famil{} saved to:\n{}".format(
        created,
        "ies" if created != 1 else "y",
        save_folder
    )
    if failed:
        msg += "\n\nFailed ({}):\n- {}".format(len(failed), "\n- ".join(failed))
    forms.alert(msg)


# ---------------------------------------------------------------------------
# MODE B — DWG → NATIVE REVIT LINES
# ---------------------------------------------------------------------------

# ── Linestyle helpers ────────────────────────────────────────────────────────

def get_all_linestyles(document):
    """Return list of Category from OST_Lines SubCategories. Read-only, no tx needed."""
    try:
        lines_cat = document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        return list(lines_cat.SubCategories)
    except Exception:
        return []


def get_linestyle_by_name(document, name):
    """Return GraphicsStyle for the named linestyle, or None."""
    for cat in get_all_linestyles(document):
        if cat.Name == name:
            return cat.GetGraphicsStyle(GraphicsStyleType.Projection)
    return None


# ── Selection filter ─────────────────────────────────────────────────────────

class ImportInstanceFilter(ISelectionFilter):
    """Allows selection of ImportInstance elements only."""
    def AllowElement(self, element):
        return isinstance(element, ImportInstance)

    def AllowReference(self, reference, position):
        return False


# ── Geometry extraction ──────────────────────────────────────────────────────

def get_layer_name(geom_obj, document):
    """Return the DWG layer name from a geometry object's GraphicsStyle."""
    try:
        gs = document.GetElement(geom_obj.GraphicsStyleId)
        if gs is not None:
            return gs.GraphicsStyleCategory.Name
    except Exception:
        pass
    return "<Unknown>"


def _iter_geom_curves(import_instance, document):
    """
    Generator that yields (layer_name, Curve) pairs from an ImportInstance.
    PolyLines are split into individual Line segments; degenerate segments skipped.
    """
    opts = Options()
    opts.ComputeReferences     = False
    opts.IncludeNonVisibleObjects = False

    geom_elem = import_instance.get_Geometry(opts)
    if geom_elem is None:
        return

    for top_obj in geom_elem:
        if not isinstance(top_obj, GeometryInstance):
            continue
        for child in top_obj.GetInstanceGeometry():
            layer = get_layer_name(child, document)
            if isinstance(child, PolyLine):
                pts = child.GetCoordinates()
                for i in range(len(pts) - 1):
                    try:
                        seg = Line.CreateBound(pts[i], pts[i + 1])
                        yield layer, seg
                    except Exception:
                        pass  # degenerate segment (identical endpoints)
            elif isinstance(child, Curve):
                try:
                    if child.Length > 1e-9:
                        yield layer, child
                except Exception:
                    pass


def extract_curves_by_layer(import_instance, document):
    """Return dict {layer_name: [Curve]} from an ImportInstance."""
    layer_map = {}
    for layer, curve in _iter_geom_curves(import_instance, document):
        layer_map.setdefault(layer, []).append(curve)
    return layer_map


def extract_all_curves(import_instance, document):
    """Return flat list of all Curves from an ImportInstance (layer-agnostic)."""
    return [curve for _, curve in _iter_geom_curves(import_instance, document)]


# ── View / line-type validation ──────────────────────────────────────────────

LINE_TYPE_LABELS = {
    'detail': 'Detail Line',
    'model':  'Model Line',
    'area':   'Area Boundary',
    'room':   'Room Separation',
    'space':  'Space Boundary',
}

LABEL_TO_KEY = {v: k for k, v in LINE_TYPE_LABELS.items()}


def validate_view_for_line_type(document, line_type):
    """
    Check the active view supports the requested line type.
    Returns (True, '') on success or (False, reason_string) on failure.
    """
    view = document.ActiveView
    is_family = document.IsFamilyDocument

    if line_type in ('area', 'room', 'space') and is_family:
        return False, "{} lines are not supported in family documents.".format(
            LINE_TYPE_LABELS[line_type])

    if line_type == 'area':
        if not (isinstance(view, ViewPlan) and view.ViewType == ViewType.AreaPlan):
            return False, "Area Boundary lines require an Area Plan view."

    if line_type in ('room', 'space'):
        if not isinstance(view, ViewPlan):
            return False, "{} lines require a Floor Plan or Reflected Ceiling Plan view.".format(
                LINE_TYPE_LABELS[line_type])

    return True, ''


# ── Line creation ────────────────────────────────────────────────────────────

def create_line(document, view, curve, line_type):
    """
    Create a single Revit line element from a Curve.
    Returns the created element, or raises on failure.
    """
    is_family = document.IsFamilyDocument

    if line_type == 'detail':
        if is_family:
            return document.FamilyCreate.NewDetailCurve(view, curve)
        return document.Create.NewDetailCurve(view, curve)

    if line_type == 'model':
        sp = view.SketchPlane
        if sp is None:
            # Family views may not have a sketch plane set — create one on XY plane at Z=0
            plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero)
            sp = SketchPlane.Create(document, plane)
        if is_family:
            return document.FamilyCreate.NewModelCurve(curve, sp)
        return document.Create.NewModelCurve(curve, sp)

    if line_type == 'area':
        sp = view.SketchPlane
        ca = CurveArray()
        ca.Append(curve)
        result = document.Create.NewAreaBoundaryLine(sp, ca, view)
        # NewAreaBoundaryLine returns a ModelCurveArray
        if result and result.Size > 0:
            return result.get_Item(0)
        return None

    if line_type == 'room':
        sp = view.SketchPlane
        ca = CurveArray()
        ca.Append(curve)
        result = document.Create.NewRoomBoundaryLines(sp, ca, view)
        if result and result.Size > 0:
            return result.get_Item(0)
        return None

    if line_type == 'space':
        sp = view.SketchPlane
        ca = CurveArray()
        ca.Append(curve)
        result = document.Create.NewSpaceBoundaryLines(sp, ca, view)
        if result and result.Size > 0:
            return result.get_Item(0)
        return None

    raise ValueError("Unknown line_type: " + line_type)


def apply_linestyle(element, document, linestyle_name):
    """Assign a linestyle to a created line element. Silently skips if not found."""
    if element is None or not linestyle_name:
        return
    gs = get_linestyle_by_name(document, linestyle_name)
    if gs is None:
        return
    try:
        element.LineStyle = gs
    except Exception:
        pass


# ── Conversion core (called inside an open Transaction) ─────────────────────

def convert_simple(document, view, curves, linestyle_name, line_type):
    """
    Create lines for all curves using a single linestyle.
    Returns (created_count, [failure_strings]).
    Must be called inside an active transaction.
    """
    created  = 0
    failures = []
    for curve in curves:
        try:
            elem = create_line(document, view, curve, line_type)
            apply_linestyle(elem, document, linestyle_name)
            created += 1
        except Exception as ex:
            failures.append(str(ex))
    return created, failures


def convert_by_layer(document, view, layer_curve_map, layer_style_map, line_type):
    """
    Create lines per layer using the provided layer→linestyle mapping.
    Layers with no mapping entry are skipped.
    Returns (created_count, [failure_strings]).
    Must be called inside an active transaction.
    """
    created  = 0
    failures = []
    for layer_name, curves in layer_curve_map.items():
        linestyle_name = layer_style_map.get(layer_name, '')
        if not linestyle_name:
            output.print_md("**Skipped layer** (no style assigned): {}".format(layer_name))
            continue
        for curve in curves:
            try:
                elem = create_line(document, view, curve, line_type)
                apply_linestyle(elem, document, linestyle_name)
                created += 1
            except Exception as ex:
                failures.append("Layer '{}': {}".format(layer_name, str(ex)))
    return created, failures


# ── By-Layer WPF dialog ──────────────────────────────────────────────────────

BY_LAYER_XAML = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Map DWG Layers to Linestyles"
    Width="540"
    SizeToContent="Height"
    MinHeight="150"
    MaxHeight="680"
    ResizeMode="CanResize"
    WindowStartupLocation="CenterScreen"
    ShowInTaskbar="False"
    FontFamily="Arial"
    FontSize="12"
    Background="#F5F5F5">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0"
               Text="Assign a linestyle to each DWG layer (leave blank to skip):"
               FontWeight="Bold"
               Margin="0,0,0,8"/>

    <ScrollViewer Grid.Row="1"
                  MaxHeight="500"
                  VerticalScrollBarVisibility="Auto">
      <StackPanel x:Name="layerStack" Margin="0,0,4,0"/>
    </ScrollViewer>

    <Separator Grid.Row="2" Margin="0,10"/>

    <StackPanel Grid.Row="3"
                Orientation="Horizontal"
                HorizontalAlignment="Right">
      <Button x:Name="btnCancel"
              Content="Cancel"
              Width="80"
              Height="26"
              Margin="0,0,8,0"/>
      <Button x:Name="btnOk"
              Content="OK"
              Width="80"
              Height="26"
              Background="#143D5C"
              Foreground="White"
              FontWeight="Bold"/>
    </StackPanel>
  </Grid>
</Window>"""


class ByLayerDialog(forms.WPFWindow):
    """WPF dialog that lets the user map each DWG layer to a linestyle."""

    def __init__(self, layer_names, linestyle_names):
        forms.WPFWindow.__init__(self, BY_LAYER_XAML, literal_string=True)
        self._ok_clicked = False
        self._combos     = {}   # {layer_name: ComboBox}

        # Build one row per layer imperatively
        blank = ['']
        all_choices = blank + list(linestyle_names)

        for layer in layer_names:
            # Container row
            row = WpfControls.Grid()
            col0 = WpfControls.ColumnDefinition()
            col0.Width = Wpf.GridLength(200)
            col1 = WpfControls.ColumnDefinition()
            col1.Width = Wpf.GridLength(1, Wpf.GridUnitType.Star)
            row.ColumnDefinitions.Add(col0)
            row.ColumnDefinitions.Add(col1)
            row.Margin = Wpf.Thickness(0, 3, 0, 3)

            # Layer name label
            lbl = WpfControls.TextBlock()
            lbl.Text              = layer
            lbl.VerticalAlignment = Wpf.VerticalAlignment.Center
            lbl.TextTrimming      = Wpf.TextTrimming.CharacterEllipsis
            lbl.ToolTip           = layer
            WpfControls.Grid.SetColumn(lbl, 0)
            row.Children.Add(lbl)

            # Linestyle ComboBox
            combo = WpfControls.ComboBox()
            for choice in all_choices:
                combo.Items.Add(choice)
            combo.SelectedIndex   = 0   # blank by default
            combo.Height          = 22
            combo.Margin          = Wpf.Thickness(8, 0, 0, 0)
            WpfControls.Grid.SetColumn(combo, 1)
            row.Children.Add(combo)

            self.layerStack.Children.Add(row)
            self._combos[layer] = combo

        self.btnOk.Click     += self.on_ok
        self.btnCancel.Click += self.on_cancel

    def on_ok(self, sender, args):
        self._ok_clicked = True
        self.Close()

    def on_cancel(self, sender, args):
        self.Close()

    def get_mapping(self):
        """Return {layer_name: linestyle_name} for layers with a selection, or {} if cancelled."""
        if not self._ok_clicked:
            return {}
        result = {}
        for layer, combo in self._combos.items():
            chosen = combo.SelectedItem
            if chosen:
                result[layer] = chosen
        return result


# ── Mode B entry point ───────────────────────────────────────────────────────

def run_mode_b():
    # 1. Guard: schedule view
    if isinstance(doc.ActiveView, ViewSchedule):
        forms.alert(
            "Please switch to a non-schedule view before running this tool.",
            exitscript=True
        )
        return

    # 2. Pick ImportInstance
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            ImportInstanceFilter(),
            "Select an imported or linked DWG file"
        )
        import_instance = doc.GetElement(ref.ElementId)
    except OperationCanceledException:
        script.exit()
        return
    except Exception as ex:
        forms.alert("Selection failed: " + str(ex), exitscript=True)
        return

    if not isinstance(import_instance, ImportInstance):
        forms.alert("Selected element is not an imported DWG. Please try again.", exitscript=True)
        return

    # 3. Choose line type — restrict options in a family document
    if doc.IsFamilyDocument:
        type_choices = ['Detail Line', 'Model Line']
    else:
        type_choices = sorted(LINE_TYPE_LABELS.values())

    line_type_label = forms.ask_for_one_item(
        type_choices,
        prompt="Select the type of Revit lines to create:",
        title="Line Type"
    )
    if not line_type_label:
        script.exit()
        return
    line_type = LABEL_TO_KEY[line_type_label]

    # 4. Validate view supports line type
    valid, reason = validate_view_for_line_type(doc, line_type)
    if not valid:
        forms.alert(reason, exitscript=True)
        return

    # 5. Assignment mode
    assign_mode = forms.ask_for_one_item(
        ['Simple (one style for all layers)', 'By Layer'],
        prompt="How should linestyles be assigned?",
        title="Assignment Mode"
    )
    if not assign_mode:
        script.exit()
        return

    # 6. Collect linestyle names
    linestyle_names = sorted([c.Name for c in get_all_linestyles(doc)])
    if not linestyle_names:
        forms.alert("No linestyles found in the current document.", exitscript=True)
        return

    # 7a. Simple mode
    if assign_mode.startswith('Simple'):
        chosen_style = forms.SelectFromList.show(
            linestyle_names,
            title="Select Linestyle",
            prompt="Choose a linestyle for all converted lines:",
            multiselect=False
        )
        if not chosen_style:
            script.exit()
            return

        curves = extract_all_curves(import_instance, doc)
        if not curves:
            forms.alert("No convertible geometry found in the selected DWG.", exitscript=True)
            return

        output.print_md("Converting **{}** curves using style **{}**...".format(
            len(curves), chosen_style))

        t = Transaction(doc, "DWG to Native Lines")
        t.Start()
        try:
            created, failures = convert_simple(doc, doc.ActiveView, curves, chosen_style, line_type)
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            forms.alert("Conversion failed: " + str(ex), exitscript=True)
            return

    # 7b. By-layer mode
    else:
        layer_curve_map = extract_curves_by_layer(import_instance, doc)
        if not layer_curve_map:
            forms.alert("No convertible geometry found in the selected DWG.", exitscript=True)
            return

        layer_names = sorted(layer_curve_map.keys())
        dlg = ByLayerDialog(layer_names, linestyle_names)
        dlg.show_dialog()
        layer_style_map = dlg.get_mapping()

        if not layer_style_map:
            script.exit()
            return

        total_curves = sum(len(v) for v in layer_curve_map.values())
        output.print_md("Converting **{}** curves across **{}** layers...".format(
            total_curves, len(layer_style_map)))

        t = Transaction(doc, "DWG to Native Lines")
        t.Start()
        try:
            created, failures = convert_by_layer(
                doc, doc.ActiveView, layer_curve_map, layer_style_map, line_type)
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            forms.alert("Conversion failed: " + str(ex), exitscript=True)
            return

    # 8. Report
    for fail_msg in failures:
        output.print_md("**Warning**: {}".format(fail_msg))

    msg = "Created {} {} line{}.".format(
        created,
        LINE_TYPE_LABELS[line_type],
        "s" if created != 1 else ""
    )
    if failures:
        msg += "\n{} curve(s) could not be converted — see output for details.".format(
            len(failures))
    forms.alert(msg)


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

def main():
    if doc is None:
        forms.alert("No document is open. Please open a project or family first.", exitscript=True)
        return

    selected = forms.CommandSwitchWindow.show(
        ['DWG to Detail Family (.rfa)', 'DWG to Native Revit Lines'],
        message='Select conversion mode:'
    )
    if not selected:
        script.exit()
        return

    if selected == 'DWG to Detail Family (.rfa)':
        run_mode_a()
    else:
        run_mode_b()


main()
