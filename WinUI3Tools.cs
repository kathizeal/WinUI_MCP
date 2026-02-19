using System.ComponentModel;
using System.Drawing.Imaging;
using System.Text;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Capturing;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using ModelContextProtocol.Server;

namespace WinUI3McpServer;

/// <summary>
/// MCP Tools for WinUI 3 app automation via FlaUI (Windows UI Automation).
/// No WinAppDriver required — talks directly to Windows UIA3.
///
/// Workflow:
///   1. LaunchApp  → get a window handle
///   2. GetSnapshot → see the accessibility tree with element refs (e.g. w1e5)
///   3. ClickElement / TypeText / FillText → interact by ref
///   4. CaptureScreenshot → visual confirmation
///   5. CloseApp
/// </summary>
[McpServerToolType]
public static class WinUI3Tools
{
    private static UIA3Automation? _automation;
    private static Window? _currentWindow;
    private static string _currentHandle = "";
    private static readonly Dictionary<string, AutomationElement> _elementRegistry = new();
    private static int _elementCounter = 0;
    private static int _windowCounter = 0;

    private static UIA3Automation Automation => _automation ??= new UIA3Automation();

    private static void EnsureWindowReady()
    {
        if (_currentWindow == null)
            throw new InvalidOperationException("No app is running. Call LaunchApp first.");
    }

    // -------------------------------------------------------------------------
    // 1. LAUNCH APP
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Launch a Windows application. Supports three modes:\n" +
        "1. Plain .exe path  → Process.Start directly\n" +
        "2. AppxManifest.xml → runs Add-AppxPackage to deploy, then launches via AUMID\n" +
        "3. AUMID string     → launches a packaged app that is already installed " +
        "   (format: 'PackageFamilyName!App')\n" +
        "After launch, call GetSnapshot to see the UI.")]
    public static string LaunchApp(
        [Description(
            "One of:\n" +
            "• Full path to .exe (e.g. C:\\MyApp\\MyApp.exe)\n" +
            "• Full path to AppxManifest.xml — deploys and launches the MSIX package\n" +
            "• AUMID of an installed packaged app (e.g. 'Abc.Xyz_1.0_x64__abc123!App')")]
        string appPath,
        [Description("Optional: working directory (only used for plain .exe launches)")]
        string workingDirectory = "")
    {
        try
        {
            _currentWindow = null;
            _elementRegistry.Clear();
            _elementCounter = 0;

            // ---- Mode 2: AppxManifest.xml deploy + launch ----
            if (appPath.EndsWith("AppxManifest.xml", StringComparison.OrdinalIgnoreCase))
                return DeployAndLaunchMsix(appPath);

            // ---- Mode 3: AUMID (contains '!' but is not a file path) ----
            if (appPath.Contains('!') && !File.Exists(appPath))
                return LaunchByAumid(appPath);

            // ---- Mode 1: Plain .exe ----
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = true,
                WorkingDirectory = string.IsNullOrWhiteSpace(workingDirectory) ? "" : workingDirectory
            };

            var process = System.Diagnostics.Process.Start(psi)
                ?? throw new Exception("Process.Start returned null.");

            try { process.WaitForInputIdle(5000); } catch { }
            Thread.Sleep(1000);

            var window = FindWindowByPid(process.Id);
            if (window == null)
                throw new Exception(
                    "Window not found after launch. Try ListWindows() to locate it manually.");

            _currentWindow = window;
            _currentHandle = $"w{++_windowCounter}";
            return $"Launched '{_currentWindow.Title}' — handle: {_currentHandle}\n" +
                   "Next: call GetSnapshot to see the accessibility tree.";
        }
        catch (Exception ex)
        {
            return $"ERROR launching app: {ex.Message}";
        }
    }

    private static string DeployAndLaunchMsix(string manifestPath)
    {
        if (!File.Exists(manifestPath))
            return $"ERROR: AppxManifest.xml not found: {manifestPath}\n" +
                   "Build the project first with BuildApp.";

        // Run Add-AppxPackage + get AUMID + launch in one PowerShell call
        var ps = $"Add-AppxPackage -Register '{manifestPath}' -ForceApplicationShutdown; " +
                 "$pkg = Get-AppxPackage | Where-Object {{ $_.InstallLocation -and " +
                 $"'{manifestPath.Replace("'", "''")}' -like \"$($_.InstallLocation)*\" }}; " +
                 "if ($pkg) { $aumid = \"$($pkg.PackageFamilyName)!App\"; " +
                 "& explorer.exe \"shell:AppsFolder\\$aumid\"; $aumid } else { 'AUMID_NOT_FOUND' }";

        var psi = new System.Diagnostics.ProcessStartInfo
        {
            FileName               = "powershell.exe",
            Arguments              = $"-NoProfile -ExecutionPolicy Bypass -Command \"{ps.Replace("\"", "\\\"")}\"",
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true
        };

        using var psProcess = System.Diagnostics.Process.Start(psi)
            ?? throw new Exception("Failed to start PowerShell.");
        var output = psProcess.StandardOutput.ReadToEnd().Trim();
        var errors = psProcess.StandardError.ReadToEnd().Trim();
        psProcess.WaitForExit();

        if (!string.IsNullOrWhiteSpace(errors))
            return $"ERROR during Add-AppxPackage:\n{errors}";

        // Wait for the app window to appear (explorer.exe spawned the real process)
        Thread.Sleep(2000);
        var window = FindNewWindow(TimeSpan.FromSeconds(10));
        if (window == null)
            return $"Package deployed (AUMID: {output}) but window not detected.\n" +
                   "Call ListWindows() to find it, then try GetSnapshot.";

        _currentWindow = window;
        _currentHandle = $"w{++_windowCounter}";
        return $"Deployed and launched '{_currentWindow.Title}' — handle: {_currentHandle}\n" +
               $"AUMID: {output}\n" +
               "Next: call GetSnapshot to see the accessibility tree.";
    }

    private static string LaunchByAumid(string aumid)
    {
        // Snapshot existing windows before launch
        var desktop = Automation.GetDesktop();
        var before = new HashSet<IntPtr>(
            desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window))
                   .Select(w => w.Properties.NativeWindowHandle.ValueOrDefault));

        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName        = "explorer.exe",
            Arguments       = $"shell:AppsFolder\\{aumid}",
            UseShellExecute = true
        });

        Thread.Sleep(2000);
        var window = FindNewWindow(TimeSpan.FromSeconds(10), before);
        if (window == null)
            return $"Launched AUMID '{aumid}' but window not detected.\n" +
                   "Call ListWindows() to find it manually.";

        _currentWindow = window;
        _currentHandle = $"w{++_windowCounter}";
        return $"Launched '{_currentWindow.Title}' — handle: {_currentHandle}\n" +
               "Next: call GetSnapshot to see the accessibility tree.";
    }

    private static Window? FindWindowByPid(int pid)
    {
        var desktop = Automation.GetDesktop();
        var element = desktop.FindFirstDescendant(cf => cf.ByProcessId(pid));
        if (element != null) return element.AsWindow();

        for (int i = 0; i < 10; i++)
        {
            Thread.Sleep(500);
            var all = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));
            var match = all.FirstOrDefault(w =>
            {
                try { return w.Properties.ProcessId.ValueOrDefault == pid; }
                catch { return false; }
            });
            if (match != null) return match.AsWindow();
        }
        return null;
    }

    private static Window? FindNewWindow(
        TimeSpan timeout, HashSet<IntPtr>? existingHandles = null)
    {
        var desktop = Automation.GetDesktop();
        if (existingHandles == null)
        {
            existingHandles = new HashSet<IntPtr>(
                desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window))
                       .Select(w => w.Properties.NativeWindowHandle.ValueOrDefault));
        }

        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            Thread.Sleep(500);
            var all = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));
            foreach (var w in all)
            {
                var hwnd = w.Properties.NativeWindowHandle.ValueOrDefault;
                if (existingHandles.Contains(hwnd)) continue;
                var win = w.AsWindow();
                if (win != null && !string.IsNullOrWhiteSpace(win.Title))
                    return win;
            }
        }
        return null;
    }

    // -------------------------------------------------------------------------
    // 2. BUILD APP
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Build a WinUI 3 project. " +
        "For unpackaged apps use dotnet build (leave platform empty). " +
        "For packaged MSIX apps (like Raptor) set platform to 'x64', 'ARM64', or 'x86' — " +
        "this switches to msbuild automatically. " +
        "Returns full output with errors so the agent can fix and rebuild. " +
        "Typical loop: edit XAML/C# → BuildApp → fix errors → BuildApp → LaunchApp → GetSnapshot.")]
    public static string BuildApp(
        [Description("Full path to the .csproj or .sln file (e.g. Z:\\source\\Raptor\\Raptor.csproj)")]
        string projectPath,
        [Description("Build configuration: 'Debug' or 'Release' (default: Debug)")]
        string configuration = "Debug",
        [Description("Target platform for packaged MSIX apps: 'x64', 'ARM64', or 'x86'. " +
                     "Leave empty for standard dotnet build.")]
        string platform = "")
    {
        try
        {
            if (!File.Exists(projectPath))
                return $"ERROR: Project file not found: {projectPath}";

            // Packaged WinUI3 apps require msbuild + explicit Platform
            bool useMsBuild = !string.IsNullOrWhiteSpace(platform);

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName               = useMsBuild ? "msbuild" : "dotnet",
                Arguments              = useMsBuild
                    ? $"\"{projectPath}\" /p:Configuration={configuration} /p:Platform={platform} /nologo"
                    : $"build \"{projectPath}\" --configuration {configuration} --nologo",
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true
            };

            using var process = System.Diagnostics.Process.Start(psi)
                ?? throw new Exception("Failed to start build process.");

            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            var sb = new StringBuilder();
            sb.AppendLine(
                $"=== {(useMsBuild ? "msbuild" : "dotnet build")}: " +
                $"{Path.GetFileName(projectPath)}" +
                $" [{configuration}{(useMsBuild ? $"|{platform}" : "")}] ===");

            if (!string.IsNullOrWhiteSpace(stdout)) sb.AppendLine(stdout);
            if (!string.IsNullOrWhiteSpace(stderr)) sb.AppendLine(stderr);

            sb.AppendLine(process.ExitCode == 0
                ? "✅ BUILD SUCCEEDED — call LaunchApp to run the app."
                : "❌ BUILD FAILED — read the errors above, fix them, then call BuildApp again.");

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"ERROR running build: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 3. GET SNAPSHOT  (replaces GetUiTree + FindElement)
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Capture the accessibility tree of the running app. Every element gets a stable ref " +
        "(e.g. w1e5). Use those refs directly with ClickElement, TypeText, FillText — " +
        "no more searching by id/name/xpath. Call this after LaunchApp or after navigating.")]
    public static string GetSnapshot(
        [Description("Maximum tree depth (1–10, default 8). Lower values are faster on large apps.")]
        int maxDepth = 8)
    {
        try
        {
            EnsureWindowReady();
            maxDepth = Math.Clamp(maxDepth, 1, 10);

            _elementRegistry.Clear();
            _elementCounter = 0;

            var sb = new StringBuilder();
            sb.AppendLine(
                $"=== Snapshot: '{_currentWindow!.Title}' [handle={_currentHandle}] ===");
            sb.AppendLine("Refs like 'w1e5' → use with ClickElement / TypeText / FillText");
            sb.AppendLine();

            BuildSnapshot(sb, _currentWindow, 0, maxDepth);
            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"ERROR getting snapshot: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 4. CLICK ELEMENT
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Click a UI element by its ref from GetSnapshot (e.g. 'w1e5'). " +
        "Uses Invoke/Toggle/SelectionItem accessibility patterns first for reliability, " +
        "falls back to mouse click.")]
    public static string ClickElement(
        [Description("Element ref from GetSnapshot, e.g. 'w1e5'")]
        string elementRef)
    {
        try
        {
            EnsureWindowReady();
            if (!_elementRegistry.TryGetValue(elementRef, out var element))
                return $"Element '{elementRef}' not found. Call GetSnapshot to refresh refs.";

            var name = element.Properties.Name.ValueOrDefault ?? elementRef;

            if (element.Patterns.Invoke.IsSupported)
            {
                element.Patterns.Invoke.Pattern.Invoke();
                return $"Invoked '{name}' [{elementRef}]";
            }

            if (element.Patterns.Toggle.IsSupported)
            {
                element.Patterns.Toggle.Pattern.Toggle();
                var state = element.Patterns.Toggle.Pattern.ToggleState.ValueOrDefault;
                return $"Toggled '{name}' [{elementRef}] → {state}";
            }

            if (element.Patterns.SelectionItem.IsSupported)
            {
                element.Patterns.SelectionItem.Pattern.Select();
                return $"Selected '{name}' [{elementRef}]";
            }

            // Mouse click fallback
            var pt = element.GetClickablePoint();
            Mouse.Click(pt);
            return $"Mouse-clicked '{name}' [{elementRef}]";
        }
        catch (Exception ex)
        {
            return $"ERROR clicking '{elementRef}': {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 5. TYPE TEXT
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Focus an element by ref and type text (appends to existing content). " +
        "Use FillText to replace content entirely.")]
    public static string TypeText(
        [Description("Element ref from GetSnapshot, e.g. 'w1e5'")]
        string elementRef,
        [Description("Text to type")]
        string text,
        [Description("Press Enter after typing (default: false)")]
        bool submit = false)
    {
        try
        {
            EnsureWindowReady();
            if (!_elementRegistry.TryGetValue(elementRef, out var element))
                return $"Element '{elementRef}' not found. Call GetSnapshot to refresh refs.";

            element.Focus();
            Thread.Sleep(50);
            Keyboard.Type(text);
            if (submit) Keyboard.Press(VirtualKeyShort.ENTER);

            return $"Typed \"{text}\" into [{elementRef}]{(submit ? " + Enter" : "")}";
        }
        catch (Exception ex)
        {
            return $"ERROR typing into '{elementRef}': {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 6. FILL TEXT
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Clear and fill a text field with a new value. Prefers the Value accessibility pattern " +
        "for reliability. Use TypeText to append without clearing.")]
    public static string FillText(
        [Description("Element ref from GetSnapshot, e.g. 'w1e5'")]
        string elementRef,
        [Description("Value to set")]
        string value)
    {
        try
        {
            EnsureWindowReady();
            if (!_elementRegistry.TryGetValue(elementRef, out var element))
                return $"Element '{elementRef}' not found. Call GetSnapshot to refresh refs.";

            var name = element.Properties.Name.ValueOrDefault ?? elementRef;

            if (element.Patterns.Value.IsSupported &&
                !element.Patterns.Value.Pattern.IsReadOnly.ValueOrDefault)
            {
                element.Patterns.Value.Pattern.SetValue(value);
                return $"Filled '{name}' [{elementRef}] with \"{value}\"";
            }

            // Keyboard fallback: select all + type
            element.Focus();
            Thread.Sleep(50);
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
            Thread.Sleep(50);
            Keyboard.Type(value);
            return $"Filled '{name}' [{elementRef}] with \"{value}\" (keyboard fallback)";
        }
        catch (Exception ex)
        {
            return $"ERROR filling '{elementRef}': {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 7. CAPTURE SCREENSHOT
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Capture a screenshot of the running app window. Returns Base64-encoded PNG. " +
        "Uses FlaUI's Capture API — no WinAppDriver needed.")]
    public static string CaptureScreenshot(
        [Description("Optional file path to also save the PNG (e.g. C:\\screenshots\\ui.png).")]
        string saveToPath = "")
    {
        try
        {
            EnsureWindowReady();

            var capture = Capture.Element(_currentWindow!);
            using var stream = new MemoryStream();
            capture.Bitmap.Save(stream, ImageFormat.Png);
            var base64 = Convert.ToBase64String(stream.ToArray());

            if (!string.IsNullOrWhiteSpace(saveToPath))
            {
                var dir = Path.GetDirectoryName(saveToPath);
                if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
                capture.Bitmap.Save(saveToPath, ImageFormat.Png);
                return $"Screenshot saved to: {saveToPath}\n\ndata:image/png;base64,{base64}";
            }

            return $"data:image/png;base64,{base64}";
        }
        catch (Exception ex)
        {
            return $"ERROR capturing screenshot: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 8. GET WINDOW INFO
    // -------------------------------------------------------------------------

    [McpServerTool, Description("Get the current window's title, size, and position.")]
    public static string GetWindowInfo()
    {
        try
        {
            EnsureWindowReady();
            var bounds = _currentWindow!.BoundingRectangle;

            var sb = new StringBuilder();
            sb.AppendLine("=== Window Info ===");
            sb.AppendLine($"Handle   : {_currentHandle}");
            sb.AppendLine($"Title    : {_currentWindow.Title}");
            sb.AppendLine($"Position : ({bounds.X}, {bounds.Y})");
            sb.AppendLine($"Size     : {bounds.Width} x {bounds.Height}");
            sb.AppendLine($"Backend  : FlaUI UIA3 (no WinAppDriver required)");
            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"ERROR getting window info: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 9. LIST WINDOWS
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "List all visible top-level windows on the desktop. " +
        "Use this to find an already-running app before calling LaunchApp.")]
    public static string ListWindows()
    {
        try
        {
            var desktop = Automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));

            var sb = new StringBuilder();
            sb.AppendLine("=== Open Windows ===");

            foreach (var w in windows)
            {
                var win = w.AsWindow();
                var title = win?.Title;
                if (string.IsNullOrWhiteSpace(title)) continue;

                try
                {
                    var pid = w.Properties.ProcessId.ValueOrDefault;
                    var proc = System.Diagnostics.Process.GetProcessById(pid).ProcessName;
                    sb.AppendLine($"  \"{title}\" [{proc}]");
                }
                catch
                {
                    sb.AppendLine($"  \"{title}\"");
                }
            }

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"ERROR listing windows: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 10. CLOSE APP
    // -------------------------------------------------------------------------

    [McpServerTool, Description("Close the current app window and clear the session.")]
    public static string CloseApp()
    {
        try
        {
            _currentWindow?.Close();
            _currentWindow = null;
            _elementRegistry.Clear();
            _elementCounter = 0;
            _currentHandle = "";
            return "App closed.";
        }
        catch (Exception ex)
        {
            return $"ERROR closing app: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // SNAPSHOT HELPERS
    // -------------------------------------------------------------------------

    private static void BuildSnapshot(
        StringBuilder sb, AutomationElement element, int depth, int maxDepth)
    {
        if (depth > maxDepth) return;

        var name = GetElementName(element);
        var role = GetElementRole(element);

        if (!ShouldSkipElement(name, role))
        {
            var refId = $"{_currentHandle}e{++_elementCounter}";
            _elementRegistry[refId] = element;

            var indent  = new string(' ', depth * 2);
            var nameStr = string.IsNullOrEmpty(name) ? "" : $" \"{name}\"";
            var states  = GetStateIndicators(element);
            var stateStr = states.Count > 0
                ? " " + string.Join(" ", states.Select(s => $"[{s}]"))
                : "";

            sb.AppendLine($"{indent}- {role}{nameStr} [ref={refId}]{stateStr}");
        }

        if (depth < maxDepth)
        {
            try
            {
                foreach (var child in element.FindAllChildren())
                    BuildSnapshot(sb, child, depth + 1, maxDepth);
            }
            catch { /* some elements don't support child traversal */ }
        }
    }

    private static string GetElementRole(AutomationElement element)
    {
        try
        {
            return element.Properties.ControlType.ValueOrDefault switch
            {
                ControlType.Button      => "button",
                ControlType.Edit        => "textbox",
                ControlType.Text        => "text",
                ControlType.CheckBox    => "checkbox",
                ControlType.RadioButton => "radio",
                ControlType.ComboBox    => "combobox",
                ControlType.List        => "list",
                ControlType.ListItem    => "listitem",
                ControlType.Menu        => "menu",
                ControlType.MenuItem    => "menuitem",
                ControlType.MenuBar     => "menubar",
                ControlType.Tree        => "tree",
                ControlType.TreeItem    => "treeitem",
                ControlType.Tab         => "tablist",
                ControlType.TabItem     => "tab",
                ControlType.Table       => "table",
                ControlType.DataItem    => "row",
                ControlType.Slider      => "slider",
                ControlType.ProgressBar => "progressbar",
                ControlType.Hyperlink   => "link",
                ControlType.Image       => "image",
                ControlType.Pane        => "group",
                ControlType.Group       => "group",
                ControlType.Window      => "window",
                ControlType.ToolBar     => "toolbar",
                ControlType.DataGrid    => "grid",
                _                       => "element"
            };
        }
        catch { return "element"; }
    }

    private static string? GetElementName(AutomationElement element)
    {
        try
        {
            var name = element.Properties.Name.ValueOrDefault;
            if (!string.IsNullOrWhiteSpace(name)) return name;

            var aid = element.Properties.AutomationId.ValueOrDefault;
            if (!string.IsNullOrWhiteSpace(aid) && aid.Length < 50)
                return $"[{aid}]";

            return null;
        }
        catch { return null; }
    }

    private static List<string> GetStateIndicators(AutomationElement element)
    {
        var states = new List<string>();
        try
        {
            if (!element.Properties.IsEnabled.ValueOrDefault)
                states.Add("disabled");
            if (element.Properties.IsOffscreen.ValueOrDefault)
                states.Add("offscreen");
            if (element.Patterns.Value.IsSupported &&
                element.Patterns.Value.Pattern.IsReadOnly.ValueOrDefault)
                states.Add("readonly");
            if (element.Patterns.Toggle.IsSupported)
            {
                var ts = element.Patterns.Toggle.Pattern.ToggleState.ValueOrDefault;
                if (ts == ToggleState.On) states.Add("checked");
                else if (ts == ToggleState.Indeterminate) states.Add("indeterminate");
            }
            if (element.Patterns.SelectionItem.IsSupported &&
                element.Patterns.SelectionItem.Pattern.IsSelected.ValueOrDefault)
                states.Add("selected");
            if (element.Patterns.ExpandCollapse.IsSupported)
            {
                var es = element.Patterns.ExpandCollapse.Pattern.ExpandCollapseState.ValueOrDefault;
                if (es == ExpandCollapseState.Expanded) states.Add("expanded");
                else if (es == ExpandCollapseState.Collapsed) states.Add("collapsed");
            }
        }
        catch { }
        return states;
    }

    private static bool ShouldSkipElement(string? name, string role)
    {
        if (!string.IsNullOrEmpty(name)) return false;

        // Always include actionable types even without a name
        if (role is "button" or "textbox" or "checkbox" or "radio" or "combobox"
                 or "listitem" or "menuitem" or "tab" or "treeitem" or "link" or "slider")
            return false;

        // Include structural containers
        if (role is "window" or "group" or "list" or "tree" or "tablist"
                 or "menu" or "menubar" or "toolbar" or "grid" or "table")
            return false;

        return true;
    }
}
