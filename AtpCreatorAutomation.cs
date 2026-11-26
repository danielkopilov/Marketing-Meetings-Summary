using System.IO;
using System.Threading;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

namespace Marketing_Meetings_Summary;

/// <summary>
/// Launches ATP Creator and fills its fields via UI Automation.
/// Strategy: find each GroupBox by its UIA Name, then locate
/// controls inside it by ControlType + order or by text — avoids
/// all screen-coordinate geometry which breaks under DPI scaling.
/// </summary>
public static class AtpCreatorAutomation
{
    /// <summary>Set true to write the ATP Creator UIA tree to the Desktop for debugging.</summary>
    public static bool DumpTreeEnabled = true;

    private static void DumpTree(AutomationElement window)
    {
        try
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== ATP Creator UIA Tree Dump ===");
            sb.AppendLine($"Timestamp: {DateTime.Now}");
            sb.AppendLine();
            DumpNode(window, 0, sb);
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                "atp_uia_dump.txt");
            File.WriteAllText(path, sb.ToString());
        }
        catch { /* ignore */ }
    }

    private static void DumpNode(AutomationElement el, int depth, System.Text.StringBuilder sb)
    {
        try
        {
            string indent = new string(' ', depth * 2);
            string ct = el.ControlType.ToString();
            string name = (el.Name ?? "").Replace("\r", " ").Replace("\n", " ");
            string aid = el.AutomationId ?? "";
            string rect;
            try
            {
                var r = el.BoundingRectangle;
                rect = $"[{(int)r.X},{(int)r.Y} {(int)r.Width}x{(int)r.Height}]";
            }
            catch { rect = "[no-rect]"; }

            sb.AppendLine($"{indent}{ct} | Name=\"{name}\" | Id=\"{aid}\" | {rect}");

            foreach (var child in el.FindAllChildren())
                DumpNode(child, depth + 1, sb);
        }
        catch { /* ignore */ }
    }

    public class AtpFillData
    {
        public string PMName          = "";
        public string SystemSN        = "";
        public string OrderNumber     = "";
        public string SystemType      = "";
        public string Variant         = "";
        public string Aperture        = "";
        public string BBType          = "";
        public string BBSize          = "";
        public string ISExitAperture  = "";
        public string BacklightType   = "";
        public string Vrs1            = "";
        public string Vrs2            = "";
        public string Vrs3            = "";
        public string GimbalSize           = "";
        public string NewPortStageMaxWeight = "";
        public string FocusStageFiniteDistance = "";
        public string CsvPath  = "";
        public string SavePath = "";
        public readonly Dictionary<string, bool> Checkboxes = new();
    }

    public static void LaunchAndFill(string exePath, AtpFillData data)
    {
        if (string.IsNullOrWhiteSpace(exePath) || !File.Exists(exePath))
            throw new FileNotFoundException("ATP Creator executable was not found.", exePath);

        var app = Application.Launch(exePath);
        using var automation = new UIA3Automation();

        // Wait for the main window.
        Window? window = null;
        for (int i = 0; i < 60 && window == null; i++)
        {
            try { window = app.GetMainWindow(automation, TimeSpan.FromSeconds(1)); }
            catch { /* keep retrying */ }
            if (window == null) Thread.Sleep(300);
        }
        if (window == null)
            throw new InvalidOperationException("Timed out waiting for the ATP Creator window.");

        // Give WinForms time to fully paint and wire up event handlers.
        Thread.Sleep(1000);

        // DIAGNOSTIC: dump the entire UIA tree so we can see exactly what
        // ControlType / Name / AutomationId / Rect each control exposes.
        // Writes to the Desktop as "atp_uia_dump.txt".
        if (DumpTreeEnabled)
            DumpTree(window);

        // ── Header text fields (located by their label-derived Name) ──────────
        SetText(FindEdit(window, "PM Name:"),    data.PMName);
        SetText(FindEdit(window, "Order#:"),     data.OrderNumber);
        SetText(FindEdit(window, "System S/N:"), data.SystemSN);

        Thread.Sleep(200);

        // Group rectangles (used to disambiguate duplicate control Names across groups).
        var rcSystemType = FindGroupRect(window, "System Type");
        var rcRad        = FindGroupRect(window, "Radiation Source");
        var rcComp       = FindGroupRect(window, "System Components");

        // All controls are searched window-wide (WinForms Groups don't reliably
        // parent their children in the UIA tree), then filtered by the group
        // rectangle they fall inside.

        // ── System Type ───────────────────────────────────────────────────────
        SelectComboByText(FindCombo(window, "Type:", rcSystemType), data.SystemType);
        Thread.Sleep(600); // cascade rebuilds Variant
        SelectComboByText(FindCombo(window, "Variant:", rcSystemType), data.Variant);
        Thread.Sleep(500); // cascade rebuilds Aperture
        SelectComboByText(FindCombo(window, "Aperture:", rcSystemType), data.Aperture);

        Thread.Sleep(200);

        // ── Radiation Source ──────────────────────────────────────────────────
        SetCheckbox(window, "B.B",       data.Checkboxes, rcRad);
        SetCheckbox(window, "I.S",       data.Checkboxes, rcRad);
        SetCheckbox(window, "Backlight", data.Checkboxes, rcRad);
        SetCheckbox(window, "LOS Laser", data.Checkboxes, rcRad);
        SetCheckbox(window, "QTH Lamp",  data.Checkboxes, rcRad);
        Thread.Sleep(150);

        // Two combos share Name "Type:" inside Radiation Source — order top→bottom:
        //   [0]=BB Type (y≈319), [1]=Backlight Type (y≈379).
        var radTypeCombos = window
            .FindAllDescendants(cf => cf.ByControlType(ControlType.ComboBox))
            .Where(c => string.Equals((c.Name ?? "").Trim(), "Type:", StringComparison.OrdinalIgnoreCase)
                        && RectContains(rcRad, c))
            .OrderBy(GetTop)
            .ToList();
        if (radTypeCombos.Count > 0) SelectComboByText(radTypeCombos[0], data.BBType);
        if (radTypeCombos.Count > 1) SelectComboByText(radTypeCombos[1], data.BacklightType);

        SelectComboByText(FindCombo(window, "Size:", rcRad),          data.BBSize);
        SelectComboByText(FindCombo(window, "Exit Aperture:", rcRad), data.ISExitAperture);

        Thread.Sleep(150);

        // ── System Components (checkboxes by Name) ────────────────────────────
        SetCheckbox(window, "Source Stage",         data.Checkboxes, rcComp);
        SetCheckbox(window, "Rackmount",            data.Checkboxes, rcComp);
        SetCheckbox(window, "CTE",                  data.Checkboxes, rcComp);
        SetCheckbox(window, "Device Center",        data.Checkboxes, rcComp);
        SetCheckbox(window, "NewPort Stage",        data.Checkboxes, rcComp);
        SetCheckbox(window, "LOS alignment target", data.Checkboxes, rcComp);
        SetCheckbox(window, "Focus Stage",          data.Checkboxes, rcComp);
        SetCheckbox(window, "VRS",                  data.Checkboxes, rcComp);
        SetCheckbox(window, "XY Stage",             data.Checkboxes, rcComp);
        SetCheckbox(window, "Power Meter",          data.Checkboxes, rcComp);
        SetCheckbox(window, "Energy Meter",         data.Checkboxes, rcComp);
        SetCheckbox(window, "Manual Choke",         data.Checkboxes, rcComp);
        SetCheckbox(window, "Gimbal",               data.Checkboxes, rcComp);

        try { window.Focus(); } catch { /* ignore */ }
    }

    /// <summary>Find an Edit by Name (window-wide).</summary>
    private static AutomationElement? FindEdit(AutomationElement window, string name)
    {
        return window
            .FindAllDescendants(cf => cf.ByControlType(ControlType.Edit))
            .FirstOrDefault(c => string.Equals((c.Name ?? "").Trim(), name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Find a ComboBox by Name, optionally restricted to a group rectangle.</summary>
    private static AutomationElement? FindCombo(AutomationElement window, string name, System.Drawing.Rectangle? rect)
    {
        return window
            .FindAllDescendants(cf => cf.ByControlType(ControlType.ComboBox))
            .FirstOrDefault(c => string.Equals((c.Name ?? "").Trim(), name, StringComparison.OrdinalIgnoreCase)
                                 && RectContains(rect, c));
    }

    /// <summary>Check/uncheck a CheckBox by Name, optionally restricted to a group rectangle.</summary>
    private static void SetCheckbox(
        AutomationElement window,
        string text,
        Dictionary<string, bool> states,
        System.Drawing.Rectangle? rect)
    {
        if (!states.TryGetValue(text, out bool desired)) return;

        var el = window
            .FindAllDescendants(cf => cf.ByControlType(ControlType.CheckBox))
            .FirstOrDefault(c => string.Equals((c.Name ?? "").Trim(), text, StringComparison.OrdinalIgnoreCase)
                                 && RectContains(rect, c));
        if (el == null) return;

        try
        {
            var cb = el.AsCheckBox();
            bool current = cb.IsChecked ?? false;
            if (current != desired) cb.IsChecked = desired;
        }
        catch { /* ignore */ }
    }

    /// <summary>Returns the bounding rectangle of a GroupBox (Group or Pane) by Name.</summary>
    private static System.Drawing.Rectangle? FindGroupRect(AutomationElement window, string name)
    {
        var gb = FindGroupBox(window, name);
        if (gb == null) return null;
        try { return gb.BoundingRectangle; }
        catch { return null; }
    }

    /// <summary>True if no rect filter, or the element's center lies within the rect.</summary>
    private static bool RectContains(System.Drawing.Rectangle? rect, AutomationElement el)
    {
        if (rect == null) return true;
        try
        {
            var r = el.BoundingRectangle;
            int cx = r.X + r.Width / 2;
            int cy = r.Y + r.Height / 2;
            return rect.Value.Contains(cx, cy);
        }
        catch { return false; }
    }

    private static double GetTop(AutomationElement el)
    {
        try { return el.BoundingRectangle.Y; }
        catch { return double.MaxValue; }
    }

    // ── Finders ───────────────────────────────────────────────────────────────

    /// <summary>Find a GroupBox by its text (UIA Name).</summary>
    private static AutomationElement? FindGroupBox(AutomationElement root, string name)
    {
        // WinForms GroupBox is a Pane or Group in UIA.
        var candidates = root.FindAllDescendants(cf => cf.ByControlType(ControlType.Group));
        var match = candidates.FirstOrDefault(el =>
            string.Equals((el.Name ?? "").Trim(), name, StringComparison.OrdinalIgnoreCase));
        if (match != null) return match;

        // Fallback: Pane
        candidates = root.FindAllDescendants(cf => cf.ByControlType(ControlType.Pane));
        return candidates.FirstOrDefault(el =>
            string.Equals((el.Name ?? "").Trim(), name, StringComparison.OrdinalIgnoreCase));
    }

    // ── Combo selection ───────────────────────────────────────────────────────

    private static void SelectComboByText(AutomationElement? el, string value)
    {
        if (el == null || string.IsNullOrEmpty(value)) return;
        string target = value.Trim();

        try
        {
            var combo = el.AsComboBox();

            // 0) FlaUI native select (handles WinForms popup windows reliably).
            try
            {
                var selected = combo.Select(target);
                if (selected != null)
                {
                    try { selected.IsSelected = true; } catch { /* ignore */ }
                    try { combo.Collapse(); } catch { /* ignore */ }
                    return;
                }
            }
            catch { /* fall through */ }

            // Expand to realize items.
            try { combo.Expand(); } catch { /* ignore */ }
            Thread.Sleep(300);

            // 1) Try items exposed directly.
            var items = el.FindAllDescendants(cf => cf.ByControlType(ControlType.ListItem));
            var match = items.FirstOrDefault(i =>
                string.Equals((i.Name ?? "").Trim(), target, StringComparison.OrdinalIgnoreCase));

            if (match == null)
            {
                match = combo.Items.FirstOrDefault(i =>
                    string.Equals(i.Text?.Trim(), target, StringComparison.OrdinalIgnoreCase));
            }

            if (match != null)
            {
                try
                {
                    if (match.Patterns.SelectionItem.IsSupported)
                        match.Patterns.SelectionItem.Pattern.Select();
                    else
                        match.AsListBoxItem().Select();
                }
                catch { try { match.Click(); } catch { /* ignore */ } }

                try { combo.Collapse(); } catch { /* ignore */ }
                return;
            }

            // 2) Fallback: keyboard — type the value into the focused combo.
            try
            {
                combo.Collapse();
                Thread.Sleep(100);
                el.Focus();
                Thread.Sleep(100);
                Keyboard.Type(target);
                Thread.Sleep(200);
                Keyboard.Press(VirtualKeyShort.RETURN);
            }
            catch { /* ignore */ }
        }
        catch { /* ignore */ }
    }

    // ── Text fill ─────────────────────────────────────────────────────────────

    private static void SetText(AutomationElement? el, string value)
    {
        if (el == null || string.IsNullOrEmpty(value)) return;
        try
        {
            var tb = el.AsTextBox();
            tb.Focus();
            // Select all then type so we don't append.
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
            Thread.Sleep(60);
            tb.Enter(value);
        }
        catch
        {
            try { el.AsTextBox().Text = value; } catch { /* ignore */ }
        }
    }
}
