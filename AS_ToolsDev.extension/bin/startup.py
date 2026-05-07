# -*- coding: utf-8 -*-
u"""
startup.py  --  AUK Tips + Tricks System
Registers a DocumentOpened event handler to display rotating tips/tutorials.

Revit 2023-2026 | pyRevit v6.0.0 | IronPython 2.7
Place at: AS_ToolsDev.extension/startup.py  (dev)  or  AS_Tools.extension/startup.py  (prod)

CONFIG FILE:  %APPDATA%\pyRevit\pyRevit_config.ini   section: [auk-tips]
"""

__persistentengine__ = True

# ============================================================================
# CONFIGURATION
# ============================================================================

TESTING_MODE = False   # SET FALSE FOR PRODUCTION

TIPS_METADATA_CSV  = r"I:\002_BIM\001_Document Reference Library\tips_metadata.csv"
TIPS_FOLDER        = r"I:\002_BIM\001_Document Reference Library\Digital Tips and Tricks"
NEWSLETTERS_FOLDER = r"I:\002_BIM\001_Document Reference Library\Digital Newsletters"
MASCOT_IMAGE_PATH  = r"I:\002_BIM\007_Addins\Aukett Swanke\Extensions\ASTool\AS_Tools.extension\resources\tip_image.png"

NETWORK_DRIVE_ROOT    = u"I:\\"    # fast pre-check before any network op
ALLOWED_EXTENSIONS    = (u".mp4", u".avi", u".mov", u".wmv", u".pdf")
TOAST_TIMEOUT_SECONDS = 15
MAX_RESOURCE_SIZE_MB  = 500
MAX_CSV_SIZE          = 1048576
MAX_CSV_ROWS          = 100
NETWORK_TIMEOUT       = 2.0
MAX_CONFIG_VALUE_LEN  = 1000
CONFIG_SECTION        = u"auk-tips"

VALID_FREQUENCIES = (u"always", u"daily", u"weekly")
DEFAULT_FREQUENCY = u"always"

AUK_BLUE     = u"#143D5C"
AUK_GOLD     = u"#E8A735"
AUK_BG       = u"#F5F5F5"
AUK_TEXT     = u"#2C2C2C"
AUK_TEXT_SEC = u"#666666"
AUK_BORDER   = u"#CCCCCC"
AUK_WHITE    = u"#FFFFFF"
AUK_BLUE_LT  = u"#B8D4DD"

FALLBACK_TIP = (
    u"AUK Tools",
    u"Explore the AUK toolbar for powerful Revit automation.",
    u"Access tools via the AUK tab in the Revit ribbon.",
    None,
)

# ============================================================================
# IMPORTS
# ============================================================================

import random
import datetime
import os
import csv
import threading
import codecs

from pyrevit import HOST_APP, framework, DB, script

import clr
clr.AddReference(u"PresentationFramework")
clr.AddReference(u"PresentationCore")

from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows.Threading import DispatcherTimer
from System import Uri, TimeSpan
from System.IO import MemoryStream
from System.Text import UTF8Encoding

# ── Module-level state ───────────────────────────────────────────────────────
_CONFIG_CACHE           = {}
_RESOURCE_CACHE         = {}
_TIPS_LOADED            = None
_ACTIVE_WINDOW          = None
_TIP_SHOWN_THIS_SESSION = False
_NETWORK_AVAILABLE      = None   # None = unchecked; True/False after first check


# ============================================================================
# SECTION 1 — DEBUG & CONFIG
# ============================================================================

def _dbg(msg):
    if TESTING_MODE:
        print(u"[AUK TIPS] {}".format(msg))


def _clear_session_cache():
    u"""Reset all module-level state on every pyRevit reload."""
    global _CONFIG_CACHE, _RESOURCE_CACHE, _TIPS_LOADED
    global _TIP_SHOWN_THIS_SESSION, _NETWORK_AVAILABLE
    _CONFIG_CACHE.clear()
    _RESOURCE_CACHE.clear()
    _TIPS_LOADED            = None
    _TIP_SHOWN_THIS_SESSION = False
    _NETWORK_AVAILABLE      = None   # re-check on next document open

    if TESTING_MODE:
        try:
            cfg = script.get_config(CONFIG_SECTION)
            setattr(cfg, u"tip_frequency", u"always")
            setattr(cfg, u"last_tip_date", u"")
            script.save_config()
            _dbg(u"TESTING_MODE: reset tip_frequency=always, last_tip_date=''")
        except Exception as ex:
            _dbg(u"TESTING_MODE config reset failed: {}".format(ex))


def safe_read_config(key, default_value):
    cache_key = u"{}|{}".format(CONFIG_SECTION, key)
    if cache_key in _CONFIG_CACHE:
        return _CONFIG_CACHE[cache_key]
    try:
        cfg     = script.get_config(CONFIG_SECTION)
        val     = cfg.get_option(key, default_value)
        str_val = u"{}".format(val)
        _CONFIG_CACHE[cache_key] = str_val
        _dbg(u"Config read  [{}.{}] = {}".format(CONFIG_SECTION, key, str_val))
        return str_val
    except Exception as ex:
        _dbg(u"Config read failed [{}.{}]: {}".format(CONFIG_SECTION, key, ex))
        return u"{}".format(default_value)


def safe_write_config(key, value):
    try:
        str_val = u"{}".format(value)
        if len(str_val) > MAX_CONFIG_VALUE_LEN:
            return False
        cfg = script.get_config(CONFIG_SECTION)
        setattr(cfg, key, str_val)
        script.save_config()
        cache_key = u"{}|{}".format(CONFIG_SECTION, key)
        _CONFIG_CACHE[cache_key] = str_val
        _dbg(u"Config write [{}.{}] = {}".format(CONFIG_SECTION, key, str_val))
        return True
    except Exception as ex:
        _dbg(u"Config write failed [{}.{}]: {}".format(CONFIG_SECTION, key, ex))
        return False


# ============================================================================
# SECTION 2 — NETWORK GUARD (zero-wait when drive unmapped)
# ============================================================================

def _check_network_available():
    u"""
    Single synchronous drive-root check — completes in <5 ms when mapped,
    instantly when unmapped (os.path.exists returns False immediately for
    a missing drive letter). Result cached for the session.
    """
    global _NETWORK_AVAILABLE
    if _NETWORK_AVAILABLE is not None:
        return _NETWORK_AVAILABLE
    try:
        _NETWORK_AVAILABLE = os.path.exists(NETWORK_DRIVE_ROOT)
    except Exception:
        _NETWORK_AVAILABLE = False
    _dbg(u"Network drive '{}' available: {}".format(
        NETWORK_DRIVE_ROOT, _NETWORK_AVAILABLE))
    return _NETWORK_AVAILABLE


def validate_resource_path(path):
    if not path:
        return False
    if not _check_network_available():
        return False
    if path in _RESOURCE_CACHE:
        return _RESOURCE_CACHE[path]
    valid = False
    try:
        if u".." not in path and os.path.isfile(path):
            if any(path.lower().endswith(ext) for ext in ALLOWED_EXTENSIONS):
                if os.path.getsize(path) <= MAX_RESOURCE_SIZE_MB * 1024 * 1024:
                    abs_path = os.path.abspath(path)
                    roots    = [os.path.abspath(TIPS_FOLDER),
                                os.path.abspath(NEWSLETTERS_FOLDER)]
                    if any(abs_path.startswith(r) for r in roots):
                        valid = True
    except Exception:
        pass
    _RESOURCE_CACHE[path] = valid
    return valid


def validate_csv_path():
    if not _check_network_available():
        _dbg(u"CSV skipped — network drive unavailable")
        return False
    try:
        return (os.path.isfile(TIPS_METADATA_CSV)
                and os.access(TIPS_METADATA_CSV, os.R_OK)
                and os.path.getsize(TIPS_METADATA_CSV) <= MAX_CSV_SIZE)
    except Exception:
        return False


def build_resource_inventory():
    u"""
    Scan tips/newsletters folders in background threads.
    Skipped entirely (instant return) if network drive is unmapped.
    """
    if not _check_network_available():
        _dbg(u"Resource scan skipped — network drive unavailable")
        return {}

    inventory = {}
    for folder in [TIPS_FOLDER, NEWSLETTERS_FOLDER]:
        bucket = {}
        err    = [None]

        def _scan(f=folder, b=bucket, e=err):
            try:
                if os.path.isdir(f):
                    for fname in os.listdir(f):
                        if fname.lower().endswith(ALLOWED_EXTENSIONS):
                            b[fname.lower()] = os.path.join(f, fname)
            except Exception as ex:
                e[0] = ex

        t = threading.Thread(target=_scan)
        t.daemon = True
        t.start()
        t.join(NETWORK_TIMEOUT)
        if not t.is_alive() and err[0] is None:
            inventory.update(bucket)
        elif t.is_alive():
            _dbg(u"Folder scan timed out: {}".format(folder))
        else:
            _dbg(u"Folder scan error: {}".format(err[0]))

    _dbg(u"Resource inventory: {} files found".format(len(inventory)))
    return inventory


# ============================================================================
# SECTION 3 — TIP LOADING
# ============================================================================

def _sanitize(value):
    v         = u"{}".format(value).strip()
    dangerous = (u"=", u"+", u"-", u"@", u"\t", u"\r", u"\n")
    while v and v[0] in dangerous:
        v = v[1:]
    return v


def load_tips_from_csv():
    u"""
    Load tips; cached after first call.
    Returns [FALLBACK_TIP] immediately if network is unavailable — no threads,
    no timeout wait, no file I/O attempted.
    """
    global _TIPS_LOADED
    if _TIPS_LOADED is not None:
        _dbg(u"Tips from cache ({} tips)".format(len(_TIPS_LOADED)))
        return _TIPS_LOADED

    # Network unavailable → instant fallback, no I/O
    if not _check_network_available():
        _dbg(u"Tips skipped — network drive unavailable; using fallback")
        _TIPS_LOADED = [FALLBACK_TIP]
        return _TIPS_LOADED

    tips = []
    if not validate_csv_path():
        _dbg(u"CSV not found or unreadable: {}".format(TIPS_METADATA_CSV))
        _TIPS_LOADED = [FALLBACK_TIP]
        return _TIPS_LOADED

    inventory = build_resource_inventory()

    try:
        with codecs.open(TIPS_METADATA_CSV, u"r", encoding=u"utf-8") as fh:
            reader = csv.DictReader(fh)
            for i, row in enumerate(reader):
                if i >= MAX_CSV_ROWS:
                    break
                try:
                    filename  = _sanitize(row.get(u"filename",    u""))
                    tool_name = _sanitize(row.get(u"tool_name",   u""))
                    desc      = _sanitize(row.get(u"description", u""))
                    usage     = _sanitize(row.get(u"when_to_use", u""))
                    if not filename or not tool_name:
                        continue
                    file_path = inventory.get(filename.lower())
                    if file_path and not validate_resource_path(file_path):
                        file_path = None
                    tips.append((tool_name, desc, usage, file_path))
                except Exception:
                    continue
    except Exception as ex:
        _dbg(u"CSV read error: {}".format(ex))

    _TIPS_LOADED = tips if tips else [FALLBACK_TIP]
    _dbg(u"Loaded {} tips".format(len(_TIPS_LOADED)))
    return _TIPS_LOADED


# ============================================================================
# SECTION 4 — FREQUENCY LOGIC
# ============================================================================

def get_date_string():
    return datetime.date.today().strftime(u"%Y-%m-%d")


def _validate_frequency(value):
    v = u"{}".format(value).strip().lower()
    return v if v in VALID_FREQUENCIES else DEFAULT_FREQUENCY


def should_show_tip(frequency):
    try:
        if frequency == u"always":
            return True
        today_str  = get_date_string()
        last_shown = u"{}".format(safe_read_config(u"last_tip_date", u""))
        today      = datetime.date.today()
        _dbg(u"Frequency={} today={} last={}".format(frequency, today_str, last_shown))
        if frequency == u"daily":
            return last_shown != today_str
        elif frequency == u"weekly":
            return today.weekday() == 4 and last_shown != today_str
        return False
    except Exception as ex:
        _dbg(u"should_show_tip error: {}".format(ex))
        return False


def record_tip_shown():
    if TESTING_MODE:
        _dbg(u"TESTING_MODE: skipping last_tip_date write")
        return
    safe_write_config(u"last_tip_date", get_date_string())


# ============================================================================
# SECTION 5 — XAML (with prev/next nav buttons)
# ============================================================================

def _escape_xml(text):
    if not text:
        return u""
    return (u"{}".format(text)
            .replace(u"&",  u"&amp;")
            .replace(u"<",  u"&lt;")
            .replace(u">",  u"&gt;")
            .replace(u'"',  u"&quot;")
            .replace(u"'",  u"&apos;"))


def create_toast_xaml(name, desc, usage, path, tip_index, tip_total):
    u"""
    Toast layout — two-column body:
      LEFT  : 180x180 mascot image (AUK Blue panel)
      RIGHT : tool name, description, resource link,
              prev/next nav, frequency selector, countdown
    Width: 520px. Header spans full width.
    """
    has_resource = bool(path)
    is_pdf       = has_resource and path.lower().endswith(u".pdf")
    link_text    = u"View Guide (PDF)" if is_pdf else u"Watch Tutorial"
    link_vis     = u"Visible" if has_resource else u"Collapsed"

    desc_preview = desc or usage or u""
    if len(desc_preview) > 200:
        desc_preview = desc_preview[:197] + u"..."

    counter_text = u"{} of {}".format(tip_index + 1, tip_total)

    name_escaped    = _escape_xml(name)
    desc_escaped    = _escape_xml(desc_preview)
    link_escaped    = _escape_xml(link_text)
    counter_escaped = _escape_xml(counter_text)

    # Shared nav button style (inline — no StaticResource needed)
    nav_btn = (
        u' Width="26" Height="26"'
        u' Background="' + AUK_BG + u'"'
        u' BorderBrush="' + AUK_BORDER + u'"'
        u' BorderThickness="1"'
        u' FontFamily="Arial"'
        u' FontSize="12"'
        u' Cursor="Hand"'
        u' VerticalAlignment="Center"'
    )

    return (
        u'<Window'
        u' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        u' Title="AUK Tip"'
        u' Width="520"'
        u' SizeToContent="Height"'
        u' ResizeMode="NoResize"'
        u' WindowStyle="None"'
        u' AllowsTransparency="True"'
        u' Background="Transparent"'
        u' Topmost="True"'
        u' ShowInTaskbar="False"'
        u' Name="ToastRoot">'

        # ── Outer shadow container ────────────────────────────────────
        u'<Border CornerRadius="8" Margin="14"'
        u' Background="' + AUK_WHITE + u'"'
        u' BorderBrush="' + AUK_BORDER + u'"'
        u' BorderThickness="1">'
        u'<Border.Effect>'
        u'<DropShadowEffect BlurRadius="16" ShadowDepth="3"'
        u' Opacity="0.22" Color="#000000"/>'
        u'</Border.Effect>'

        u'<Grid>'
        u'<Grid.RowDefinitions>'
        u'<RowDefinition Height="Auto"/>'   # header
        u'<RowDefinition Height="Auto"/>'   # body
        u'</Grid.RowDefinitions>'

        # ── Header ───────────────────────────────────────────────────
        u'<Border Grid.Row="0"'
        u' Background="' + AUK_BLUE + u'"'
        u' CornerRadius="8,8,0,0"'
        u' Padding="16,10,12,10">'
        u'<Grid>'
        u'<Grid.ColumnDefinitions>'
        u'<ColumnDefinition Width="*"/>'
        u'<ColumnDefinition Width="Auto"/>'
        u'</Grid.ColumnDefinitions>'
        u'<StackPanel Grid.Column="0" VerticalAlignment="Center">'
        u'<TextBlock Text="AUK Tips + Tricks"'
        u' FontFamily="Arial" FontSize="13" FontWeight="Bold"'
        u' Foreground="' + AUK_WHITE + u'"/>'
        u'<TextBlock Name="ToolName"'
        u' Text="' + name_escaped + u'"'
        u' FontFamily="Arial" FontSize="11"'
        u' Foreground="' + AUK_BLUE_LT + u'"'
        u' TextTrimming="CharacterEllipsis"/>'
        u'</StackPanel>'
        u'<Button Name="CloseButton" Grid.Column="1"'
        u' Content="&#x2715;" Width="24" Height="24"'
        u' Background="Transparent"'
        u' Foreground="' + AUK_WHITE + u'"'
        u' BorderThickness="0" FontSize="12"'
        u' Cursor="Hand" VerticalAlignment="Top"'
        u' FontFamily="Arial" ToolTip="Dismiss"/>'
        u'</Grid>'
        u'</Border>'

        # ── Body ─────────────────────────────────────────────────────
        u'<Grid Grid.Row="1">'
        u'<Grid.ColumnDefinitions>'
        u'<ColumnDefinition Width="180"/>'
        u'<ColumnDefinition Width="*"/>'
        u'</Grid.ColumnDefinitions>'

        # LEFT: mascot image
        u'<Border Grid.Column="0"'
        u' Background="' + AUK_BLUE + u'"'
        u' CornerRadius="0,0,0,8">'
        u'<Image Name="MascotImage"'
        u' Width="180" Height="180"'
        u' Stretch="UniformToFill"'
        u' HorizontalAlignment="Center"'
        u' VerticalAlignment="Center"/>'
        u'</Border>'

        # RIGHT: content
        u'<StackPanel Grid.Column="1"'
        u' Margin="16,14,14,14"'
        u' VerticalAlignment="Top">'

        # Description
        u'<TextBlock Name="TipContent"'
        u' Text="' + desc_escaped + u'"'
        u' FontFamily="Arial" FontSize="11"'
        u' Foreground="' + AUK_TEXT + u'"'
        u' TextWrapping="Wrap"'
        u' LineHeight="17"'
        u' Margin="0,0,0,10"/>'

        # Gold divider
        u'<Rectangle Height="2" Fill="' + AUK_GOLD + u'"'
        u' HorizontalAlignment="Left" Width="40"'
        u' Margin="0,0,0,10"/>'

        # Resource hyperlink
        u'<TextBlock Margin="0,0,0,12"'
        u' Visibility="' + link_vis + u'"'
        u' Name="LearnMorePanel">'
        u'<Hyperlink Name="LearnMoreLink"'
        u' Foreground="' + AUK_BLUE + u'"'
        u' FontFamily="Arial" FontSize="11"'
        u' FontWeight="Bold">'
        u'<Run Name="LinkRun" Text="' + link_escaped + u'"/>'
        u'</Hyperlink>'
        u'</TextBlock>'

        # ── Nav row: prev | counter | next ───────────────────────────
        u'<Grid Margin="0,0,0,10">'
        u'<Grid.ColumnDefinitions>'
        u'<ColumnDefinition Width="Auto"/>'   # prev button
        u'<ColumnDefinition Width="*"/>'      # counter (centred)
        u'<ColumnDefinition Width="Auto"/>'   # next button
        u'</Grid.ColumnDefinitions>'

        u'<Button Name="PrevButton" Grid.Column="0"'
        u' Content="&#x2039;"'               # ‹
        + nav_btn +
        u' ToolTip="Previous tip"/>'

        u'<TextBlock Name="TipCounter"'
        u' Grid.Column="1"'
        u' Text="' + counter_escaped + u'"'
        u' FontFamily="Arial" FontSize="10"'
        u' Foreground="' + AUK_TEXT_SEC + u'"'
        u' HorizontalAlignment="Center"'
        u' VerticalAlignment="Center"/>'

        u'<Button Name="NextButton" Grid.Column="2"'
        u' Content="&#x203A;"'               # ›
        + nav_btn +
        u' ToolTip="Next tip"/>'

        u'</Grid>'

        # ── Bottom row: frequency + countdown ────────────────────────
        u'<Grid>'
        u'<Grid.ColumnDefinitions>'
        u'<ColumnDefinition Width="Auto"/>'
        u'<ColumnDefinition Width="*"/>'
        u'<ColumnDefinition Width="Auto"/>'
        u'</Grid.ColumnDefinitions>'
        u'<StackPanel Grid.Column="0" Orientation="Horizontal">'
        u'<TextBlock Text="Show:"'
        u' FontFamily="Arial" FontSize="10"'
        u' Foreground="' + AUK_TEXT_SEC + u'"'
        u' VerticalAlignment="Center"'
        u' Margin="0,0,4,0"/>'
        u'<ComboBox Name="FrequencyCombo"'
        u' Width="118" Height="22"'
        u' FontFamily="Arial" FontSize="10"'
        u' BorderBrush="' + AUK_BORDER + u'">'
        u'<ComboBoxItem Content="Every open"      Tag="always"/>'
        u'<ComboBoxItem Content="Daily"           Tag="daily"/>'
        u'<ComboBoxItem Content="Weekly (Friday)" Tag="weekly"/>'
        u'</ComboBox>'
        u'</StackPanel>'
        u'<TextBlock Name="CountdownText"'
        u' Grid.Column="2"'
        u' FontFamily="Arial" FontSize="10"'
        u' Foreground="' + AUK_TEXT_SEC + u'"'
        u' VerticalAlignment="Center"'
        u' Text=""/>'
        u'</Grid>'

        u'</StackPanel>'   # end right content
        u'</Grid>'         # end body
        u'</Grid>'         # end outer grid
        u'</Border>'       # end shadow container
        u'</Window>'
    )


# ============================================================================
# SECTION 6 — IN-PLACE TIP NAVIGATION
# ============================================================================

def _update_tip_content(window, tips, index):
    u"""
    Update all dynamic controls in the existing window to show a new tip.
    No window rebuild — purely updates named WPF element properties.
    """
    tip       = tips[index]
    name      = tip[0] if len(tip) > 0 else u""
    desc      = tip[1] if len(tip) > 1 else u""
    usage     = tip[2] if len(tip) > 2 else u""
    doc_path  = tip[3] if len(tip) > 3 else None

    desc_preview = desc or usage or u""
    if len(desc_preview) > 200:
        desc_preview = desc_preview[:197] + u"..."

    try:
        tool_name_tb = window.FindName(u"ToolName")
        if tool_name_tb:
            tool_name_tb.Text = name

        tip_content_tb = window.FindName(u"TipContent")
        if tip_content_tb:
            tip_content_tb.Text = desc_preview

        counter_tb = window.FindName(u"TipCounter")
        if counter_tb:
            counter_tb.Text = u"{} of {}".format(index + 1, len(tips))

        # Resource link visibility + text
        panel = window.FindName(u"LearnMorePanel")
        link  = window.FindName(u"LearnMoreLink")
        run   = window.FindName(u"LinkRun")

        if panel and link and run:
            if doc_path and validate_resource_path(doc_path):
                from System.Windows import Visibility as Vis
                panel.Visibility = Vis.Visible
                is_pdf           = doc_path.lower().endswith(u".pdf")
                run.Text         = u"View Guide (PDF)" if is_pdf else u"Watch Tutorial"

                # Rewire click handler — remove old, add new
                try:
                    link.Click -= link.Click
                except Exception:
                    pass
                link.Click += lambda s, a, p=doc_path: open_resource_safe(p)
            else:
                from System.Windows import Visibility as Vis
                panel.Visibility = Vis.Collapsed

    except Exception as ex:
        _dbg(u"_update_tip_content error: {}".format(ex))


def _wire_nav_buttons(window, tips, start_index):
    u"""
    Wire Prev/Next buttons. State held in a mutable list so both closures
    share and mutate the same index without nonlocal (IronPython 2.7).
    """
    idx = [start_index]

    prev_btn = window.FindName(u"PrevButton")
    next_btn = window.FindName(u"NextButton")

    if not prev_btn or not next_btn:
        _dbg(u"Nav buttons not found in XAML")
        return

    def _on_prev(sender, args):
        idx[0] = (idx[0] - 1) % len(tips)
        _dbg(u"Prev → tip index {}".format(idx[0]))
        _update_tip_content(window, tips, idx[0])

    def _on_next(sender, args):
        idx[0] = (idx[0] + 1) % len(tips)
        _dbg(u"Next → tip index {}".format(idx[0]))
        _update_tip_content(window, tips, idx[0])

    prev_btn.Click += _on_prev
    next_btn.Click += _on_next


# ============================================================================
# SECTION 7 — WINDOW WIRING & DISPLAY
# ============================================================================

def safe_close_window(window):
    global _ACTIVE_WINDOW
    try:
        window.Close()
    except Exception:
        pass
    finally:
        _ACTIVE_WINDOW = None


def open_resource_safe(path):
    if validate_resource_path(path):
        try:
            os.startfile(path)
        except Exception as ex:
            _dbg(u"open_resource_safe failed: {}".format(ex))


def _load_mascot_image(window):
    if not _check_network_available():
        return
    if not os.path.isfile(MASCOT_IMAGE_PATH):
        _dbg(u"Mascot not found: {}".format(MASCOT_IMAGE_PATH))
        return
    try:
        mascot = window.FindName(u"MascotImage")
        if not mascot:
            return
        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.UriSource   = Uri(MASCOT_IMAGE_PATH)
        bmp.CacheOption = BitmapCacheOption.OnLoad
        bmp.EndInit()
        mascot.Source = bmp
    except Exception as ex:
        _dbg(u"Mascot load failed: {}".format(ex))


def _wire_hyperlink(window, doc_path):
    if not doc_path or not validate_resource_path(doc_path):
        return
    try:
        link = window.FindName(u"LearnMoreLink")
        if link:
            link.Click += lambda s, a, p=doc_path: open_resource_safe(p)
    except Exception as ex:
        _dbg(u"Hyperlink wire failed: {}".format(ex))


def _setup_frequency_combo(window):
    try:
        saved_freq = _validate_frequency(
            safe_read_config(u"tip_frequency", DEFAULT_FREQUENCY)
        )
        combo = window.FindName(u"FrequencyCombo")
        if not combo:
            return
        for item in combo.Items:
            if hasattr(item, u"Tag") and u"{}".format(item.Tag) == saved_freq:
                combo.SelectedItem = item
                break
    except Exception as ex:
        _dbg(u"Freq combo setup failed: {}".format(ex))


def _save_frequency_from_combo(window):
    try:
        cb = window.FindName(u"FrequencyCombo")
        if cb and cb.SelectedItem and hasattr(cb.SelectedItem, u"Tag"):
            freq = _validate_frequency(u"{}".format(cb.SelectedItem.Tag))
        else:
            freq = DEFAULT_FREQUENCY
        if not TESTING_MODE:
            safe_write_config(u"tip_frequency", freq)
    except Exception as ex:
        _dbg(u"Save freq failed: {}".format(ex))


def _position_bottom_right(window):
    u"""Called from Loaded event — ActualHeight is resolved by then."""
    try:
        from System.Windows import SystemParameters
        area        = SystemParameters.WorkArea
        window.Left = area.Right  - window.ActualWidth  - 8
        window.Top  = area.Bottom - window.ActualHeight - 8
    except Exception as ex:
        _dbg(u"Position failed: {}".format(ex))


def _start_countdown(window, total_seconds):
    try:
        ct        = window.FindName(u"CountdownText")
        remaining = [total_seconds]

        timer          = DispatcherTimer()
        timer.Interval = TimeSpan.FromSeconds(1)

        def _tick(sender, args):
            remaining[0] -= 1
            try:
                if ct:
                    ct.Text = u"Closing in {}s".format(remaining[0])
            except Exception:
                pass
            if remaining[0] <= 0:
                try:
                    sender.Stop()
                except Exception:
                    pass
                _save_frequency_from_combo(window)
                safe_close_window(window)

        timer.Tick += _tick
        timer.Start()
        if ct:
            ct.Text = u"Closing in {}s".format(total_seconds)
    except Exception as ex:
        _dbg(u"Countdown failed: {}".format(ex))


def show_tip_toast(tips, start_index):
    u"""
    Build the toast for the tip at start_index; wire nav for the full list.
    tips      : full list of (tool_name, desc, usage, file_path) tuples
    start_index: which tip to display first
    """
    global _ACTIVE_WINDOW

    if _ACTIVE_WINDOW is not None:
        safe_close_window(_ACTIVE_WINDOW)

    tip      = tips[start_index]
    name     = tip[0] if len(tip) > 0 else u""
    desc     = tip[1] if len(tip) > 1 else u""
    usage    = tip[2] if len(tip) > 2 else u""
    doc_path = tip[3] if len(tip) > 3 else None

    _dbg(u"Showing tip {}/{}: '{}'".format(start_index + 1, len(tips), name))

    def _build_and_show(xaml_str):
        enc    = UTF8Encoding()
        stream = MemoryStream(enc.GetBytes(xaml_str))
        try:
            win = XamlReader.Load(stream)
        finally:
            try:
                stream.Dispose()
            except Exception:
                pass

        _load_mascot_image(win)
        _wire_hyperlink(win, doc_path)
        _setup_frequency_combo(win)
        _wire_nav_buttons(win, tips, start_index)

        close_btn = win.FindName(u"CloseButton")
        if close_btn:
            def _on_close(s, a, w=win):
                _save_frequency_from_combo(w)
                safe_close_window(w)
            close_btn.Click += _on_close

        def _on_loaded(s, a, w=win):
            _position_bottom_right(w)
            _start_countdown(w, TOAST_TIMEOUT_SECONDS)

        win.Loaded += _on_loaded
        return win

    try:
        xaml           = create_toast_xaml(name, desc, usage, doc_path,
                                           start_index, len(tips))
        window         = _build_and_show(xaml)
        _ACTIVE_WINDOW = window
        window.Show()

    except Exception as ex:
        _dbg(u"Toast error (transparent): {}".format(ex))
        try:
            xaml_fb = create_toast_xaml(name, desc, usage, doc_path,
                                        start_index, len(tips))
            xaml_fb = (xaml_fb
                       .replace(u'AllowsTransparency="True"',  u'AllowsTransparency="False"')
                       .replace(u'Background="Transparent"',   u'Background="' + AUK_WHITE + u'"'))
            window2        = _build_and_show(xaml_fb)
            _ACTIVE_WINDOW = window2
            window2.Show()
        except Exception as ex2:
            _dbg(u"Toast fallback also failed: {}".format(ex2))


# ============================================================================
# SECTION 8 — EVENT HANDLER
# ============================================================================

def _on_document_opened(sender, args):
    global _TIP_SHOWN_THIS_SESSION
    try:
        frequency = _validate_frequency(
            safe_read_config(u"tip_frequency", DEFAULT_FREQUENCY)
        )
        _dbg(u"DocumentOpened — frequency={} session_shown={}".format(
            frequency, _TIP_SHOWN_THIS_SESSION))

        if _TIP_SHOWN_THIS_SESSION and frequency != u"always" and not TESTING_MODE:
            _dbg(u"Skipping: already shown this session")
            return

        # Network check is instant — cached after first call
        tips = load_tips_from_csv()

        show = TESTING_MODE or should_show_tip(frequency)
        _dbg(u"show={}".format(show))

        if show:
            start_index = random.randint(0, len(tips) - 1)
            record_tip_shown()
            _TIP_SHOWN_THIS_SESSION = True
            show_tip_toast(tips, start_index)

    except Exception as ex:
        _dbg(u"Handler error: {}".format(ex))


# ============================================================================
# SECTION 9 — REGISTRATION
# ============================================================================

_clear_session_cache()

try:
    HOST_APP.app.DocumentOpened += \
        framework.EventHandler[DB.Events.DocumentOpenedEventArgs](
            _on_document_opened
        )
    _dbg(
        u"DocumentOpened handler registered. "
        u"Config: %APPDATA%\\pyRevit\\pyRevit_config.ini  section=[{}]".format(CONFIG_SECTION)
    )
except Exception as ex:
    print(u"[AUK TIPS] Failed to register handler: {}".format(ex))