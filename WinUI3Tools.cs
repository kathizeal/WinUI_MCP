using System.ComponentModel;
using System.Drawing.Imaging;
using System.Text;
using System.Text.RegularExpressions;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Capturing;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using ModelContextProtocol.Protocol;
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
        "3. AUMID string     → launches a packaged app that is already installed\n" +
        "   (format: 'PackageFamilyName!App', e.g. from GetInstalledPackages or DeployApp)\n" +
        "Use forceRestart=true after DeployApp to kill any stale process and start fresh.\n" +
        "After launch, call GetSnapshot to see the UI.\n" +
        "To attach to an app already running (launched outside this tool), use AttachToApp.")]
    public static string LaunchApp(
        [Description(
            "One of:\n" +
            "• Full path to .exe (e.g. C:\\MyApp\\MyApp.exe)\n" +
            "• Full path to AppxManifest.xml — deploys and launches the MSIX package\n" +
            "• AUMID of an installed packaged app (e.g. 'Abc.Xyz_1.0_x64__abc123!App')")]
        string appPath,
        [Description("Optional: working directory (only used for plain .exe launches)")]
        string workingDirectory = "",
        [Description("Kill any existing instance before launching. Use true after DeployApp to avoid attaching to stale process.")]
        bool forceRestart = false)
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
                return LaunchByAumid(appPath, forceRestart);

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

        var aumid = DeployManifest(manifestPath, out var deployError);
        if (aumid == null) return deployError!;

        Thread.Sleep(2000);
        var window = FindNewWindow(TimeSpan.FromSeconds(10));
        if (window == null)
            return $"Package deployed (AUMID: {aumid}) but window not detected.\n" +
                   "Try AttachToApp(title) or ListWindows().";

        _currentWindow = window;
        _currentHandle = $"w{++_windowCounter}";
        return $"Deployed and launched '{_currentWindow.Title}' — handle: {_currentHandle}\n" +
               $"AUMID: {aumid}\n" +
               "Next: call GetSnapshot to see the accessibility tree.";
    }

    // Shared: deploy manifest and return AUMID (or null + error message)
    private static string? DeployManifest(string manifestPath, out string? error)
    {
        var manifestEsc = manifestPath.Replace("'", "''");
        var manifestDir = Path.GetDirectoryName(manifestPath)!.Replace("'", "''");
        var ps1 = Path.GetTempFileName() + ".ps1";
        File.WriteAllText(ps1,
            $"Add-AppxPackage -Register '{manifestEsc}' -ForceApplicationShutdown -ErrorAction Stop\n" +
            $"$d = '{manifestDir}'\n" +
            "$pkg = Get-AppxPackage | Where-Object { $_.InstallLocation -and $d.StartsWith($_.InstallLocation, [System.StringComparison]::OrdinalIgnoreCase) }\n" +
            "if ($pkg) {\n" +
            "    $aumid = \"$($pkg.PackageFamilyName)!App\"\n" +
            "    & explorer.exe \"shell:AppsFolder\\$aumid\"\n" +
            "    Write-Output $aumid\n" +
            "} else {\n" +
            "    Write-Output 'AUMID_NOT_FOUND'\n" +
            "}");
        try
        {
            var output = RunPs1(ps1, out var errors);
            if (!string.IsNullOrWhiteSpace(errors) && errors.Contains("Error", StringComparison.OrdinalIgnoreCase))
            {
                error = $"ERROR during Add-AppxPackage:\n{errors}";
                return null;
            }
            if (output == "AUMID_NOT_FOUND")
            {
                error = "Package registered but AUMID not found. Try GetInstalledPackages().";
                return null;
            }
            error = null;
            return output;
        }
        finally { try { File.Delete(ps1); } catch { } }
    }

    private static string LaunchByAumid(string aumid, bool forceRestart = false)
    {
        var pfn = aumid.Contains('!') ? aumid.Split('!')[0] : aumid;
        var existing = FindWindowByPackageFamilyName(pfn);

        if (existing != null && !forceRestart)
        {
            _currentWindow = existing;
            _currentHandle = $"w{++_windowCounter}";
            return $"App already running — attached to '{_currentWindow.Title}' [handle: {_currentHandle}].\n" +
                   "Next: call GetSnapshot to see the accessibility tree.";
        }

        // forceRestart=true or no existing window: kill old instance then launch fresh
        if (existing != null && forceRestart)
        {
            try
            {
                var pid = existing.Properties.ProcessId.ValueOrDefault;
                System.Diagnostics.Process.GetProcessById(pid).Kill();
                Thread.Sleep(1500);
            }
            catch { }
        }
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
                   "Call AttachToApp(title) or ListWindows() to find it manually.";

        _currentWindow = window;
        _currentHandle = $"w{++_windowCounter}";
        return $"Launched '{_currentWindow.Title}' — handle: {_currentHandle}\n" +
               "Next: call GetSnapshot to see the accessibility tree.";
    }

    private static Window? FindWindowByPackageFamilyName(string pfn)
    {
        var ps1 = Path.GetTempFileName() + ".ps1";
        File.WriteAllText(ps1,
            $"$p=Get-AppxPackage|Where-Object{{$_.PackageFamilyName -eq '{pfn.Replace("'","''")}'}};if($p){{$p.InstallLocation}}");
        try
        {
            var loc = RunPs1(ps1, out _).Trim();
            if (string.IsNullOrEmpty(loc)) return null;

            var desktop = Automation.GetDesktop();
            foreach (var w in desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window)))
            {
                try
                {
                    var pid = w.Properties.ProcessId.ValueOrDefault;
                    var proc = System.Diagnostics.Process.GetProcessById(pid);
                    if (proc.MainModule?.FileName?.StartsWith(loc, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        var win = w.AsWindow();
                        if (win != null && !string.IsNullOrWhiteSpace(win.Title)) return win;
                    }
                }
                catch { }
            }
        }
        finally { try { File.Delete(ps1); } catch { } }
        return null;
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

                // Skip explorer.exe windows — it's the launcher, not the target app
                try
                {
                    var pid = w.Properties.ProcessId.ValueOrDefault;
                    var pName = System.Diagnostics.Process.GetProcessById(pid).ProcessName;
                    if (pName.Equals("explorer", StringComparison.OrdinalIgnoreCase)) continue;
                }
                catch { }

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
        "For packaged MSIX apps set platform to 'x64', 'ARM64', or 'x86' — uses msbuild automatically. " +
        "Returns a concise summary on success. Returns filtered errors/warnings on failure. " +
        "Typical flow: BuildApp → DeployApp → LaunchApp(aumid) → GetSnapshot.")]
    public static string BuildApp(
        [Description("Full path to the .csproj or .sln file (e.g. Z:\\source\\Raptor\\Raptor.csproj)")]
        string projectPath,
        [Description("Build configuration: 'Debug' or 'Release' (default: Debug)")]
        string configuration = "Debug",
        [Description("Target platform for packaged MSIX apps: 'x64', 'ARM64', or 'x86'. Leave empty for dotnet build.")]
        string platform = "",
        [Description("Set true to get the full build log. Default false returns a one-line summary on success.")]
        bool verbose = false)
    {
        try
        {
            if (!File.Exists(projectPath))
                return $"ERROR: Project file not found: {projectPath}";

            bool useMsBuild = !string.IsNullOrWhiteSpace(platform);

            var sw = System.Diagnostics.Stopwatch.StartNew();
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
            sw.Stop();

            var allOutput = stdout + stderr;
            var succeeded = process.ExitCode == 0;

            if (succeeded)
            {
                // Parse warning count
                var warnCount = Regex.Matches(allOutput, @"\bwarning\b", RegexOptions.IgnoreCase).Count;
                var elapsed   = $"{sw.Elapsed.TotalSeconds:F1}s";
                var syncMsg   = useMsBuild ? SyncAppXFolder(projectPath, configuration, platform) : "";
                var summary   = $"✅ Build succeeded in {elapsed} (0 errors, {warnCount} warnings)";
                if (!string.IsNullOrEmpty(syncMsg)) summary += $"\n{syncMsg}";
                summary += "\nNext: call DeployApp(manifestPath) then LaunchApp(aumid).";

                return verbose ? $"{summary}\n\n{allOutput}" : summary;
            }
            else
            {
                // Return only error/warning lines — skip the noise
                var errorLines = allOutput.Split('\n')
                    .Where(l => Regex.IsMatch(l, @": ?(error|warning) ", RegexOptions.IgnoreCase)
                             || l.Contains("Build FAILED", StringComparison.OrdinalIgnoreCase)
                             || Regex.IsMatch(l, @"\d+ Error\(s\)"))
                    .Select(l => l.Trim())
                    .Where(l => l.Length > 0);

                return $"❌ BUILD FAILED\n{string.Join("\n", errorLines)}";
            }
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
        "(e.g. w1e5). Use those refs with ClickElement, TypeText, FillText, ScrollElement. " +
        "Call this after LaunchApp or after navigating. Use includeBounds=true to see " +
        "element positions and sizes for layout analysis.")]
    public static string GetSnapshot(
        [Description("Maximum tree depth (1–10, default 8). Lower values are faster on large apps.")]
        int maxDepth = 8,
        [Description("Include (x,y w×h) bounding rect for each element. Default false.")]
        bool includeBounds = false)
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
            sb.AppendLine("Refs like 'w1e5' → use with ClickElement / TypeText / FillText / ScrollElement");
            sb.AppendLine();

            BuildSnapshot(sb, _currentWindow, 0, maxDepth, includeBounds);
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
        "Capture a screenshot of the running app window. " +
        "Returns an inline image content block (vision models can see it directly) plus " +
        "a text summary of visible UI elements. No WinAppDriver needed.")]
    public static IEnumerable<ContentBlock> CaptureScreenshot(
        [Description("Optional file path to also save the PNG (e.g. C:\\screenshots\\ui.png).")]
        string saveToPath = "")
    {
        EnsureWindowReady();

        var capture = FlaUI.Core.Capturing.Capture.Element(_currentWindow!);
        using var stream = new MemoryStream();
        capture.Bitmap.Save(stream, ImageFormat.Png);
        var base64 = Convert.ToBase64String(stream.ToArray());

        if (!string.IsNullOrWhiteSpace(saveToPath))
        {
            var dir = Path.GetDirectoryName(saveToPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
            capture.Bitmap.Save(saveToPath, ImageFormat.Png);
        }

        // Text summary: window info + named elements for text-mode reasoning
        var meta = new StringBuilder();
        meta.AppendLine($"=== Screenshot: '{_currentWindow!.Title}' ===");
        meta.AppendLine($"Size: {capture.Bitmap.Width}x{capture.Bitmap.Height}px");
        if (!string.IsNullOrWhiteSpace(saveToPath)) meta.AppendLine($"Saved: {saveToPath}");
        meta.AppendLine("Visible UI elements:");
        CollectVisibleText(meta, _currentWindow, 0, 4);

        return
        [
            new TextContentBlock  { Text = meta.ToString().TrimEnd() },
            new ImageContentBlock { Data = base64, MimeType = "image/png" }
        ];
    }

    private static void CollectVisibleText(StringBuilder sb, AutomationElement el, int depth, int maxDepth)
    {
        if (depth > maxDepth) return;
        try
        {
            var role = GetElementRole(el);
            var name = GetElementName(el);
            if (!string.IsNullOrEmpty(name) && role is not ("window" or "element" or "group"))
            {
                var indent = new string(' ', depth * 2);
                sb.AppendLine($"{indent}[{role}] {name}");
            }
            foreach (var child in el.FindAllChildren())
                CollectVisibleText(sb, child, depth + 1, maxDepth);
        }
        catch { }
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
    // 11. ATTACH TO RUNNING APP
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Attach to an already-running application by process name or window title. " +
        "Use this when the app was started outside of LaunchApp (e.g. via F5 in VS, " +
        "Start-Process, or a deploy task). Partial match, case-insensitive.")]
    public static string AttachToApp(
        [Description("Process name (e.g. 'Raptor') or window title (e.g. 'Raptor') — partial match OK")]
        string nameOrTitle)
    {
        try
        {
            var desktop = Automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));
            var search = nameOrTitle.ToLowerInvariant();

            Window? match = null;
            string? matchedBy = null;

            foreach (var w in windows)
            {
                var win = w.AsWindow();
                if (win == null) continue;

                var title = win.Title ?? "";
                if (title.ToLowerInvariant().Contains(search))
                {
                    match = win;
                    matchedBy = $"title \"{title}\"";
                    break;
                }

                try
                {
                    var pid = w.Properties.ProcessId.ValueOrDefault;
                    var proc = System.Diagnostics.Process.GetProcessById(pid);
                    if (proc.ProcessName.ToLowerInvariant().Contains(search))
                    {
                        match = win;
                        matchedBy = $"process \"{proc.ProcessName}\" (pid {pid})";
                        break;
                    }
                }
                catch { }
            }

            if (match == null)
                return $"No window found matching '{nameOrTitle}'.\n" +
                       "Call ListWindows() to see all open windows.";

            _currentWindow = match;
            _currentHandle = $"w{++_windowCounter}";
            return $"Attached to '{_currentWindow.Title}' via {matchedBy} — handle: {_currentHandle}\n" +
                   "Next: call GetSnapshot to see the accessibility tree.";
        }
        catch (Exception ex)
        {
            return $"ERROR in AttachToApp: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 12. DEPLOY APP
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Deploy a packaged WinUI 3 MSIX app by registering its AppxManifest.xml. " +
        "Automatically syncs ALL build outputs (DLL, EXE, XBF, PRI, etc.) from the output " +
        "directory into the AppX folder before registering, so stale files are never deployed. " +
        "Returns the AUMID. " +
        "Typical flow: BuildApp → DeployApp → LaunchApp(aumid, forceRestart=true).")]
    public static string DeployApp(
        [Description("Full path to AppxManifest.xml, e.g. Z:\\source\\Raptor\\bin\\x64\\Debug\\net8.0-windows10.0.19041.0\\win-x64\\AppX\\AppxManifest.xml")]
        string manifestPath)
    {
        try
        {
            if (!File.Exists(manifestPath))
                return $"ERROR: AppxManifest.xml not found: {manifestPath}";

            // Sync all output files into AppX before registering
            var syncMsg = SyncAppXFromManifest(manifestPath);

            var aumid = DeployManifest(manifestPath, out var deployError);
            if (aumid == null) return deployError!;

            return $"✅ Package deployed successfully.\n" +
                   $"{syncMsg}\n" +
                   $"AUMID: {aumid}\n" +
                   $"Next: call LaunchApp(\"{aumid}\", forceRestart=true) to launch fresh.";
        }
        catch (Exception ex)
        {
            return $"ERROR deploying package: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 13. GET INSTALLED PACKAGES
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "List installed packaged (MSIX) apps matching a name filter. " +
        "Returns the AUMID for each match — paste into LaunchApp or DeployApp.")]
    public static string GetInstalledPackages(
        [Description("Partial name to filter (e.g. 'Raptor'). Leave empty to list all.")]
        string filter = "")
    {
        var ps1 = Path.GetTempFileName() + ".ps1";
        File.WriteAllText(ps1,
            "Get-AppxPackage | Select-Object Name,PackageFamilyName,Version | ConvertTo-Json -Compress");
        try
        {
            var json = RunPs1(ps1, out _);
            var sb = new StringBuilder();
            sb.AppendLine("=== Installed Packages (Name | Version | AUMID) ===");

            var entries = Regex.Matches(json,
                @"""Name"":""([^""]+)""[^}]*""PackageFamilyName"":""([^""]+)""[^}]*""Version"":""([^""]+)""");

            int count = 0;
            foreach (Match m in entries)
            {
                var name = m.Groups[1].Value;
                var pfn  = m.Groups[2].Value;
                var ver  = m.Groups[3].Value;
                if (!string.IsNullOrEmpty(filter) &&
                    !name.Contains(filter, StringComparison.OrdinalIgnoreCase) &&
                    !pfn.Contains(filter, StringComparison.OrdinalIgnoreCase))
                    continue;

                sb.AppendLine($"  {name}  v{ver}");
                sb.AppendLine($"    AUMID: {pfn}!App");
                count++;
            }

            if (count == 0)
                sb.AppendLine(string.IsNullOrEmpty(filter)
                    ? "No packages found."
                    : $"No packages matching '{filter}'.");

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"ERROR listing packages: {ex.Message}";
        }
        finally { try { File.Delete(ps1); } catch { } }
    }

    // -------------------------------------------------------------------------
    // 14. SCROLL ELEMENT
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Scroll a scrollable element (ScrollViewer, list, panel) by ref. " +
        "Use this to reveal offscreen items before calling GetSnapshot again.")]
    public static string ScrollElement(
        [Description("Element ref from GetSnapshot, e.g. 'w1e5'. The element or a scrollable ancestor will be scrolled.")]
        string elementRef,
        [Description("Scroll direction: 'up', 'down', 'left', 'right'")]
        string direction = "down",
        [Description("Number of scroll increments (default 3)")]
        int amount = 3)
    {
        try
        {
            EnsureWindowReady();
            if (!_elementRegistry.TryGetValue(elementRef, out var element))
                return $"Element '{elementRef}' not found. Call GetSnapshot to refresh refs.";

            // Walk up to find a scrollable ancestor
            var scrollable = element;
            while (scrollable != null && !scrollable.Patterns.Scroll.IsSupported)
            {
                try { scrollable = scrollable.Parent; } catch { scrollable = null; }
            }

            if (scrollable == null || !scrollable.Patterns.Scroll.IsSupported)
                return $"No scrollable element found at or above '{elementRef}'.";

            var scroll = scrollable.Patterns.Scroll.Pattern;
            var dir = direction.ToLowerInvariant();

            for (int i = 0; i < amount; i++)
            {
                if (dir == "down")  scroll.Scroll(FlaUI.Core.Definitions.ScrollAmount.NoAmount, FlaUI.Core.Definitions.ScrollAmount.SmallIncrement);
                else if (dir == "up")    scroll.Scroll(FlaUI.Core.Definitions.ScrollAmount.NoAmount, FlaUI.Core.Definitions.ScrollAmount.SmallDecrement);
                else if (dir == "right") scroll.Scroll(FlaUI.Core.Definitions.ScrollAmount.SmallIncrement, FlaUI.Core.Definitions.ScrollAmount.NoAmount);
                else if (dir == "left")  scroll.Scroll(FlaUI.Core.Definitions.ScrollAmount.SmallDecrement, FlaUI.Core.Definitions.ScrollAmount.NoAmount);
                Thread.Sleep(50);
            }

            return $"Scrolled {direction} ×{amount} on [{elementRef}]. Call GetSnapshot to see updated content.";
        }
        catch (Exception ex)
        {
            return $"ERROR scrolling '{elementRef}': {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // PRIVATE HELPERS
    // -------------------------------------------------------------------------

    // Run a .ps1 file, return trimmed stdout, set errors to trimmed stderr
    private static string RunPs1(string ps1Path, out string errors)
    {
        var psi = new System.Diagnostics.ProcessStartInfo
        {
            FileName               = "powershell.exe",
            Arguments              = $"-NoProfile -ExecutionPolicy Bypass -File \"{ps1Path}\"",
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true
        };
        using var proc = System.Diagnostics.Process.Start(psi)
            ?? throw new Exception("Failed to start PowerShell.");
        var stdout = proc.StandardOutput.ReadToEnd().Trim();
        errors = proc.StandardError.ReadToEnd().Trim();
        proc.WaitForExit();
        return stdout;
    }

    // Sync ALL changed output files from the directory containing AppX into AppX itself
    private static string SyncAppXFromManifest(string manifestPath)
    {
        try
        {
            var appxDir   = Path.GetDirectoryName(manifestPath)!;
            var outputDir = Path.GetDirectoryName(appxDir)!;
            var synced    = new List<string>();
            var skipped   = new List<string>();

            // Extensions that matter for WinUI3 packaged apps
            var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { ".dll", ".exe", ".xbf", ".pri", ".runtimeconfig.json", ".deps.json" };

            foreach (var srcFile in Directory.GetFiles(outputDir))
            {
                if (!exts.Contains(Path.GetExtension(srcFile))) continue;
                var fname   = Path.GetFileName(srcFile);
                var dstFile = Path.Combine(appxDir, fname);
                if (!File.Exists(dstFile)) continue; // Only update existing AppX files
                try
                {
                    if (File.GetLastWriteTimeUtc(srcFile) > File.GetLastWriteTimeUtc(dstFile))
                    {
                        File.Copy(srcFile, dstFile, overwrite: true);
                        synced.Add(fname);
                    }
                }
                catch { skipped.Add(fname); }
            }

            if (synced.Count > 0) return $"📦 Synced {synced.Count} file(s): {string.Join(", ", synced)}";
            if (skipped.Count > 0) return $"⚠️ AppX up to date but {skipped.Count} file(s) locked (app running?)";
            return "📦 AppX already up to date";
        }
        catch (Exception ex) { return $"⚠️ Sync warning: {ex.Message}"; }
    }

    // After a successful msbuild, copy .dll/.exe outputs into the AppX folder
    private static string SyncAppXFolder(string projectPath, string configuration, string platform)
    {
        try
        {
            var projectDir  = Path.GetDirectoryName(projectPath)!;
            var projectName = Path.GetFileNameWithoutExtension(projectPath);
            var binRoot     = Path.Combine(projectDir, "bin", platform, configuration);
            if (!Directory.Exists(binRoot)) return "";

            var appxDirs = Directory.GetDirectories(binRoot, "AppX", SearchOption.AllDirectories);
            if (appxDirs.Length == 0) return "";

            // Use the manifest-based sync for completeness (handles XBF/PRI too)
            var manifestPath = Path.Combine(appxDirs[0], "AppxManifest.xml");
            return File.Exists(manifestPath) ? SyncAppXFromManifest(manifestPath) : "";
        }
        catch { return ""; }
    }

    private static void BuildSnapshot(
        StringBuilder sb, AutomationElement element, int depth, int maxDepth, bool includeBounds = false)
    {
        if (depth > maxDepth) return;

        var name = GetElementName(element);
        var role = GetElementRole(element);

        if (!ShouldSkipElement(name, role))
        {
            var refId = $"{_currentHandle}e{++_elementCounter}";
            _elementRegistry[refId] = element;

            var indent   = new string(' ', depth * 2);
            var nameStr  = string.IsNullOrEmpty(name) ? "" : $" \"{name}\"";
            var states   = GetStateIndicators(element);
            var stateStr = states.Count > 0
                ? " " + string.Join(" ", states.Select(s => $"[{s}]"))
                : "";
            var boundsStr = "";
            if (includeBounds)
            {
                try
                {
                    var r = element.BoundingRectangle;
                    boundsStr = $" ({r.X},{r.Y} {r.Width}×{r.Height})";
                }
                catch { }
            }

            sb.AppendLine($"{indent}- {role}{nameStr} [ref={refId}]{stateStr}{boundsStr}");
        }

        if (depth < maxDepth)
        {
            try
            {
                foreach (var child in element.FindAllChildren())
                    BuildSnapshot(sb, child, depth + 1, maxDepth, includeBounds);
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
