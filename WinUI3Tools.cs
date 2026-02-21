using System.ComponentModel;
using System.Drawing.Imaging;
using System.Net.Http;
using System.Text;
using System.Text.Json;
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
        "3. AUMID string     → launches a packaged app already installed\n" +
        "   (format: 'PackageFamilyName!App', e.g. from GetInstalledPackages or DeployApp)\n" +
        "Use forceRestart=true to kill any stale process and start fresh.\n" +
        "Set autoScreenshot=true (default) to immediately capture the window after launch.\n" +
        "To attach to an app started outside this tool, use AttachToApp.")]
    public static IEnumerable<ContentBlock> LaunchApp(
        [Description(
            "One of:\n" +
            "• Full path to .exe (e.g. C:\\MyApp\\MyApp.exe)\n" +
            "• Full path to AppxManifest.xml — deploys and launches the MSIX package\n" +
            "• AUMID of an installed packaged app (e.g. 'Abc.Xyz_1.0_x64__abc123!App')")]
        string appPath,
        [Description("Optional: working directory (only used for plain .exe launches)")]
        string workingDirectory = "",
        [Description("Kill any existing instance before launching. Use true after DeployApp.")]
        bool forceRestart = false,
        [Description("Automatically capture a screenshot after launch so you can see the initial UI. Default true.")]
        bool autoScreenshot = true)
    {
        string resultText;
        try
        {
            _currentWindow = null;
            _elementRegistry.Clear();
            _elementCounter = 0;
            _windowCounter  = 0;

            // ---- Mode 2: AppxManifest.xml deploy + launch ----
            if (appPath.EndsWith("AppxManifest.xml", StringComparison.OrdinalIgnoreCase))
                resultText = DeployAndLaunchMsix(appPath);
            // ---- Mode 3: AUMID (contains '!' but is not a file path) ----
            else if (appPath.Contains('!') && !File.Exists(appPath))
                resultText = LaunchByAumid(appPath, forceRestart);
            else
            {
                // ---- Mode 1: Plain .exe ----
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName         = appPath,
                    UseShellExecute  = true,
                    WorkingDirectory = string.IsNullOrWhiteSpace(workingDirectory) ? "" : workingDirectory
                };
                var process = System.Diagnostics.Process.Start(psi)
                    ?? throw new Exception("Process.Start returned null.");
                try { process.WaitForInputIdle(5000); } catch { }
                Thread.Sleep(1000);
                var window = FindWindowByPid(process.Id);
                if (window == null)
                    throw new Exception("Window not found after launch. Try ListWindows() to locate it manually.");
                _currentWindow = window;
                _currentHandle = $"w{++_windowCounter}";
                resultText = $"Launched '{_currentWindow.Title}' — handle: {_currentHandle}\n" +
                             "Next: call GetSnapshot to see the accessibility tree.";
            }
        }
        catch (Exception ex)
        {
            return [new TextContentBlock { Text = $"ERROR launching app: {ex.Message}" }];
        }

        if (autoScreenshot && _currentWindow != null)
        {
            try
            {
                WaitForWindowReadyInternal(2000, 2);
                var hwnd = _currentWindow.Properties.NativeWindowHandle.ValueOrDefault;
                using var bmp = CaptureHwnd(hwnd);
                using var stream = new MemoryStream();
                bmp.Save(stream, ImageFormat.Png);
                return
                [
                    new TextContentBlock  { Text = resultText },
                    new ImageContentBlock { Data = Convert.ToBase64String(stream.ToArray()), MimeType = "image/png" }
                ];
            }
            catch { /* screenshot failed — still return text */ }
        }

        return [new TextContentBlock { Text = resultText }];
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

        // Fallback: FindNewWindow might miss if the app was already running or launched slowly
        if (window == null)
            window = FindWindowByPackageFamilyName(pfn);

        if (window == null)
            return $"Launched AUMID '{aumid}' but window not detected after 12s.\n" +
                   "The app may still be starting — try AttachToApp(title, processName=) or ListWindows().";

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

                try
                {
                    var pid   = w.Properties.ProcessId.ValueOrDefault;
                    var pName = System.Diagnostics.Process.GetProcessById(pid).ProcessName;
                    if (IsIdeProcess(pName)) continue;
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
        "Returns a concise summary on success. Returns classified errors on failure. " +
        "Typical flow: BuildApp → DeployApp → LaunchApp(aumid) → GetSnapshot.")]
    public static string BuildApp(
        [Description("Full path to the .csproj or .sln file (e.g. Z:\\source\\Raptor\\Raptor.csproj)")]
        string projectPath,
        [Description("Build configuration: 'Debug' or 'Release' (default: Debug)")]
        string configuration = "Debug",
        [Description("Target platform for packaged MSIX apps: 'x64', 'ARM64', or 'x86'. Leave empty for dotnet build.")]
        string platform = "",
        [Description("Set true to get the full build log. Default false returns a summary on success, classified errors on failure.")]
        bool verbose = false,
        [Description("Incremental build (default true). Set false to force a full rebuild (/t:Rebuild).")]
        bool incremental = true)
    {
        try
        {
            if (!File.Exists(projectPath))
                return $"ERROR: Project file not found: {projectPath}";

            bool useMsBuild = !string.IsNullOrWhiteSpace(platform);
            var  target     = (!incremental && useMsBuild) ? " /t:Rebuild" : "";

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName               = useMsBuild ? "msbuild" : "dotnet",
                Arguments              = useMsBuild
                    ? $"\"{projectPath}\" /p:Configuration={configuration} /p:Platform={platform} /nologo{target}"
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
                // Classify each error line: XAML (WMC*/XBF*), C# (CS*), MSBuild (MSB*)
                var errorLines = allOutput.Split('\n')
                    .Where(l => Regex.IsMatch(l, @": ?(error|warning) [A-Z0-9]+:", RegexOptions.IgnoreCase)
                             || l.Contains("Build FAILED", StringComparison.OrdinalIgnoreCase))
                    .Select(l => l.Trim())
                    .Where(l => l.Length > 0)
                    .Distinct()
                    .ToList();

                if (errorLines.Count == 0)
                    errorLines = allOutput.Split('\n')
                        .Where(l => l.Contains("error", StringComparison.OrdinalIgnoreCase) && l.Trim().Length > 0)
                        .Select(l => l.Trim())
                        .Take(20)
                        .ToList();

                var sb = new StringBuilder("❌ BUILD FAILED\n");
                var xamlErrors = errorLines.Where(l => Regex.IsMatch(l, @"\b(WMC|XBF)\d+\b")).ToList();
                var csErrors   = errorLines.Where(l => Regex.IsMatch(l, @"\bCS\d+\b")).ToList();
                var otherErrors = errorLines.Except(xamlErrors).Except(csErrors).ToList();

                if (xamlErrors.Count > 0)
                {
                    sb.AppendLine($"\n[XAML errors — {xamlErrors.Count}]");
                    xamlErrors.ForEach(e => sb.AppendLine("  " + e));
                }
                if (csErrors.Count > 0)
                {
                    sb.AppendLine($"\n[C# errors — {csErrors.Count}]");
                    csErrors.ForEach(e => sb.AppendLine("  " + e));
                }
                if (otherErrors.Count > 0)
                {
                    sb.AppendLine($"\n[Other — {otherErrors.Count}]");
                    otherErrors.ForEach(e => sb.AppendLine("  " + e));
                }
                return sb.ToString();
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
        int maxDepth = 10,
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
        "Capture a screenshot of the attached app window (identified by its handle, not Z-order). " +
        "Returns an inline image the vision model can see directly, plus a text element summary. " +
        "Works correctly even when the window is behind other windows.")]
    public static IEnumerable<ContentBlock> CaptureScreenshot(
        [Description("Optional file path to also save the PNG (e.g. C:\\screenshots\\ui.png).")]
        string saveToPath = "")
    {
        EnsureWindowReady();

        var hwnd = _currentWindow!.Properties.NativeWindowHandle.ValueOrDefault;
        var bmp  = CaptureHwnd(hwnd);

        using var stream = new MemoryStream();
        bmp.Save(stream, ImageFormat.Png);
        var base64 = Convert.ToBase64String(stream.ToArray());

        if (!string.IsNullOrWhiteSpace(saveToPath))
        {
            var dir = Path.GetDirectoryName(saveToPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
            bmp.Save(saveToPath, ImageFormat.Png);
        }
        bmp.Dispose();

        var meta = new StringBuilder();
        meta.AppendLine($"=== Screenshot: '{_currentWindow!.Title}' [handle={_currentHandle}] ===");
        var r = _currentWindow.BoundingRectangle;
        meta.AppendLine($"Size: {r.Width}x{r.Height}px  Position: ({r.X},{r.Y})");
        if (!string.IsNullOrWhiteSpace(saveToPath)) meta.AppendLine($"Saved: {saveToPath}");
        meta.AppendLine("Visible UI elements:");
        CollectVisibleText(meta, _currentWindow, 0, 4);

        return
        [
            new TextContentBlock  { Text = meta.ToString().TrimEnd() },
            new ImageContentBlock { Data = base64, MimeType = "image/png" }
        ];
    }

    // Capture a specific HWND using PrintWindow(PW_RENDERFULLCONTENT=0x2),
    // which works for DirectX/WinUI windows regardless of Z-order.
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr hwnd, out NativeRect rect);

    private struct NativeRect { public int Left, Top, Right, Bottom; }

    private static System.Drawing.Bitmap CaptureHwnd(IntPtr hwnd)
    {
        GetClientRect(hwnd, out var r);
        int w = Math.Max(r.Right - r.Left, 1);
        int h = Math.Max(r.Bottom - r.Top, 1);

        var bmp = new System.Drawing.Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var g   = System.Drawing.Graphics.FromImage(bmp);
        var hdc = g.GetHdc();
        PrintWindow(hwnd, hdc, 0x2); // PW_RENDERFULLCONTENT
        g.ReleaseHdc(hdc);
        return bmp;
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
    // 7b. CAPTURE ELEMENT (element-region screenshot)
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Capture a screenshot of a specific UI element by ref (from GetSnapshot). " +
        "Returns an inline image of just that element's bounding region — " +
        "useful for focused before/after comparisons of individual UI components.")]
    public static IEnumerable<ContentBlock> CaptureElement(
        [Description("Element ref from GetSnapshot, e.g. 'w1e5'")]
        string elementRef,
        [Description("Optional file path to also save the PNG.")]
        string saveToPath = "")
    {
        EnsureWindowReady();
        if (!_elementRegistry.TryGetValue(elementRef, out var element))
            return [new TextContentBlock { Text = $"Element '{elementRef}' not found. Call GetSnapshot first." }];

        var hwnd = _currentWindow!.Properties.NativeWindowHandle.ValueOrDefault;
        using var fullBmp = CaptureHwnd(hwnd);

        // Get element bounds relative to window client area
        var winRect = _currentWindow.BoundingRectangle;
        var elRect  = element.BoundingRectangle;

        var relX = elRect.X - winRect.X;
        var relY = elRect.Y - winRect.Y;
        var w    = Math.Max(elRect.Width,  1);
        var h    = Math.Max(elRect.Height, 1);

        // Clamp to bitmap bounds
        relX = Math.Max(0, Math.Min(relX, fullBmp.Width  - 1));
        relY = Math.Max(0, Math.Min(relY, fullBmp.Height - 1));
        w    = Math.Min(w, fullBmp.Width  - relX);
        h    = Math.Min(h, fullBmp.Height - relY);

        var cropped = fullBmp.Clone(new System.Drawing.Rectangle(relX, relY, w, h),
                                    fullBmp.PixelFormat);
        using var stream = new MemoryStream();
        cropped.Save(stream, ImageFormat.Png);
        var base64 = Convert.ToBase64String(stream.ToArray());

        if (!string.IsNullOrWhiteSpace(saveToPath))
        {
            var dir = Path.GetDirectoryName(saveToPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
            cropped.Save(saveToPath, ImageFormat.Png);
        }
        cropped.Dispose();

        var name = GetElementName(element) ?? elementRef;
        var role = GetElementRole(element);
        return
        [
            new TextContentBlock  { Text = $"[{role}] \"{name}\" ({w}×{h}px at {elRect.X},{elRect.Y})" },
            new ImageContentBlock { Data = base64, MimeType = "image/png" }
        ];
    }

    // -------------------------------------------------------------------------
    // 7c. COMPARE SCREENSHOT
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Capture the current window and compare it against a reference image file. " +
        "Returns a diff summary: overall similarity %, per-region breakdown, and " +
        "a list of areas where the images differ significantly. " +
        "Use this to verify a UI change matched an expected design.")]
    public static string CompareScreenshot(
        [Description("Full path to the reference/target image (PNG or JPEG).")]
        string referenceImagePath,
        [Description("Difference threshold per-pixel (0–255, default 30). Lower = more strict.")]
        int threshold = 30)
    {
        try
        {
            EnsureWindowReady();
            if (!File.Exists(referenceImagePath))
                return $"ERROR: Reference image not found: {referenceImagePath}";

            var hwnd = _currentWindow!.Properties.NativeWindowHandle.ValueOrDefault;
            using var current = CaptureHwnd(hwnd);
            using var reference = (System.Drawing.Bitmap)System.Drawing.Image.FromFile(referenceImagePath);

            // Resize reference to match current if sizes differ
            System.Drawing.Bitmap? refResized = null;
            var refBmp = reference;
            if (reference.Width != current.Width || reference.Height != current.Height)
            {
                refResized = new System.Drawing.Bitmap(reference, current.Width, current.Height);
                refBmp = refResized;
            }

            int totalPixels   = current.Width * current.Height;
            int diffPixels    = 0;
            int w = current.Width, h = current.Height;

            // Divide into a 4×4 grid to report which region differs
            int gridW = Math.Max(1, w / 4);
            int gridH = Math.Max(1, h / 4);
            var gridDiff = new int[4, 4];
            var gridTotal = new int[4, 4];

            // Lock bits for fast pixel access
            var curData = current.LockBits(
                new System.Drawing.Rectangle(0, 0, w, h),
                System.Drawing.Imaging.ImageLockMode.ReadOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var refData = refBmp.LockBits(
                new System.Drawing.Rectangle(0, 0, w, h),
                System.Drawing.Imaging.ImageLockMode.ReadOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            unsafe
            {
                var curPtr = (byte*)curData.Scan0;
                var refPtr = (byte*)refData.Scan0;
                for (int y = 0; y < h; y++)
                {
                    for (int x = 0; x < w; x++)
                    {
                        int idx = y * curData.Stride + x * 4;
                        int dr  = Math.Abs(curPtr[idx + 2] - refPtr[idx + 2]);
                        int dg  = Math.Abs(curPtr[idx + 1] - refPtr[idx + 1]);
                        int db  = Math.Abs(curPtr[idx + 0] - refPtr[idx + 0]);
                        int diff = (dr + dg + db) / 3;

                        int gx = Math.Min(x / gridW, 3);
                        int gy = Math.Min(y / gridH, 3);
                        gridTotal[gy, gx]++;

                        if (diff > threshold)
                        {
                            diffPixels++;
                            gridDiff[gy, gx]++;
                        }
                    }
                }
            }
            current.UnlockBits(curData);
            refBmp.UnlockBits(refData);
            refResized?.Dispose();

            double similarity = 100.0 * (totalPixels - diffPixels) / totalPixels;
            var sb = new StringBuilder();
            sb.AppendLine($"=== Screenshot Comparison ===");
            sb.AppendLine($"Current : {_currentWindow!.Title} ({w}×{h}px)");
            sb.AppendLine($"Reference: {Path.GetFileName(referenceImagePath)}");
            sb.AppendLine($"Similarity: {similarity:F1}%  ({diffPixels:N0} / {totalPixels:N0} pixels differ, threshold={threshold})");
            sb.AppendLine();

            // Grid map
            sb.AppendLine("Diff grid (% of region changed) — rows=top→bottom, cols=left→right:");
            string[] rowLabels = ["Top   ", "Upper ", "Lower ", "Bottom"];
            string[] colLabels = ["Left", "CtrL", "CtrR", "Rght"];
            sb.Append("        ");
            foreach (var c in colLabels) sb.Append($"  {c}");
            sb.AppendLine();
            for (int gy = 0; gy < 4; gy++)
            {
                sb.Append($"  {rowLabels[gy]}");
                for (int gx = 0; gx < 4; gx++)
                {
                    double pct = gridTotal[gy, gx] > 0
                        ? 100.0 * gridDiff[gy, gx] / gridTotal[gy, gx]
                        : 0;
                    sb.Append(pct > 5 ? $"  {pct,4:F0}%" : "     ·");
                }
                sb.AppendLine();
            }

            sb.AppendLine();
            if (similarity >= 95)
                sb.AppendLine("✅ Images are visually very similar.");
            else if (similarity >= 75)
                sb.AppendLine("⚠️ Moderate differences detected. Check the grid above for locations.");
            else
                sb.AppendLine("❌ Significant differences. The UI likely does not match the reference.");

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"ERROR comparing screenshot: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 7d. VALIDATE XAML
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Statically validate WinUI 3 XAML files before building. Catches common errors instantly " +
        "without running msbuild: malformed XML, unknown WinUI attributes on Window/Page, " +
        "duplicate x:Name values, missing closing tags. " +
        "Run this after editing XAML and before BuildApp to catch mistakes early.")]
    public static string ValidateXaml(
        [Description("Full path to a .xaml file, or a directory to scan all .xaml files recursively.")]
        string path)
    {
        try
        {
            var files = File.Exists(path)
                ? [path]
                : Directory.Exists(path)
                    ? Directory.GetFiles(path, "*.xaml", SearchOption.AllDirectories)
                    : Array.Empty<string>();

            if (files.Length == 0)
                return $"No .xaml files found at: {path}";

            var sb = new StringBuilder();
            int errorCount = 0, fileCount = 0;

            foreach (var file in files)
            {
                var issues = ValidateXamlFile(file);
                fileCount++;
                if (issues.Count > 0)
                {
                    sb.AppendLine($"\n❌ {Path.GetFileName(file)}:");
                    foreach (var issue in issues)
                    {
                        sb.AppendLine($"   {issue}");
                        errorCount++;
                    }
                }
            }

            if (errorCount == 0)
                return $"✅ {fileCount} XAML file(s) validated — no issues found.";

            sb.Insert(0, $"Found {errorCount} issue(s) across {fileCount} file(s):\n");
            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"ERROR validating XAML: {ex.Message}";
        }
    }

    private static List<string> ValidateXamlFile(string path)
    {
        var issues = new List<string>();
        string xml;
        try { xml = File.ReadAllText(path); }
        catch (Exception ex) { return [$"Cannot read file: {ex.Message}"]; }

        // 1. Well-formed XML check
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(xml);

            // 2. Unknown attributes on <Window> / <Page>
            var ns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
            var root = doc.Root;
            if (root != null)
            {
                // Attributes that do NOT belong on Window/Page root
                var badWindowAttribs = new HashSet<string>
                    { "RequestedTheme", "Background", "Foreground" };
                var rootName = root.Name.LocalName;
                if (rootName is "Window" or "Page")
                {
                    foreach (var attr in root.Attributes())
                    {
                        if (badWindowAttribs.Contains(attr.Name.LocalName))
                            issues.Add($"Line ~1: '{attr.Name.LocalName}' is not valid on <{rootName}>. " +
                                       $"Move it to the root Grid/StackPanel instead.");
                    }
                }

                // 3. Duplicate x:Name values
                var xNs = "http://schemas.microsoft.com/winfx/2006/xaml";
                var names = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (var el in doc.Descendants())
                {
                    var xName = el.Attribute(System.Xml.Linq.XName.Get("Name", xNs))?.Value
                             ?? el.Attribute("x:Name")?.Value;
                    if (string.IsNullOrEmpty(xName)) continue;
                    names.TryGetValue(xName, out int cnt);
                    names[xName] = cnt + 1;
                }
                foreach (var kv in names.Where(kv => kv.Value > 1))
                    issues.Add($"Duplicate x:Name=\"{kv.Key}\" ({kv.Value} occurrences).");

                // 4. Grid with ColumnDefinitions but no Column assignments
                foreach (var grid in doc.Descendants(System.Xml.Linq.XName.Get("Grid", ns)))
                {
                    var colDefs = grid.Element(System.Xml.Linq.XName.Get("Grid.ColumnDefinitions", ns));
                    if (colDefs == null || colDefs.Elements().Count() <= 1) continue;
                    bool anyAssigned = grid.Elements().Any(c =>
                        c.Attribute(System.Xml.Linq.XName.Get("Column", ns)) != null ||
                        c.Attribute("Grid.Column") != null);
                    if (!anyAssigned)
                        issues.Add($"Grid has {colDefs.Elements().Count()} ColumnDefinitions but no child has Grid.Column set.");
                }
            }
        }
        catch (System.Xml.XmlException ex)
        {
            issues.Add($"XML parse error at line {ex.LineNumber}: {ex.Message}");
        }

        return issues;
    }

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
                    var pid   = w.Properties.ProcessId.ValueOrDefault;
                    var proc  = System.Diagnostics.Process.GetProcessById(pid).ProcessName;
                    var isIde = IsIdeProcess(proc) ? " [IDE]" : "";
                    sb.AppendLine($"  \"{title}\" [process={proc} pid={pid}]{isIde}");
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
            _currentWindow    = null;
            _elementRegistry.Clear();
            _elementCounter   = 0;
            _windowCounter    = 0;
            _currentHandle    = "";
            return "App closed. Session reset — next LaunchApp will use handle w1.";
        }
        catch (Exception ex)
        {
            return $"ERROR closing app: {ex.Message}";
        }
    }

    // -------------------------------------------------------------------------
    // 11. ATTACH TO RUNNING APP
    // -------------------------------------------------------------------------

    // Well-known IDE/shell processes whose windows must never be mistaken for a target app
    private static readonly HashSet<string> KnownIdeProcesses = new(StringComparer.OrdinalIgnoreCase)
    {
        "Code", "Code - Insiders", "devenv", "rider", "idea64", "studio64",
        "AndroidStudio", "fleet", "notepad", "notepad++", "sublime_text",
        "explorer", "SearchHost", "StartMenuExperienceHost", "ShellExperienceHost",
        "ApplicationFrameHost"
    };

    [McpServerTool, Description(
        "Attach to an already-running application window. " +
        "Use this when the app was started outside of LaunchApp (e.g. via VS F5, Start-Process). " +
        "Supply pid for the most precise attach (no ambiguity). " +
        "Without pid: exact title > exact process name > partial (IDEs always excluded). " +
        "Prefix title with 'regex:' for pattern matching, e.g. 'regex:^Raptor$'.")]
    public static string AttachToApp(
        [Description("Window title to match. Prefix with 'regex:' for a regex pattern, e.g. 'regex:^Raptor$' for exact.")]
        string title,
        [Description("Optional: process name filter (e.g. 'Raptor'). Narrows the search when title alone is ambiguous.")]
        string processName = "",
        [Description("Optional: attach directly by process ID (most precise — no title ambiguity).")]
        int pid = 0)
    {
        try
        {
            // PID-based attach — most precise, bypasses all title matching
            if (pid > 0)
            {
                System.Diagnostics.Process proc;
                try { proc = System.Diagnostics.Process.GetProcessById(pid); }
                catch { return $"ERROR: No process with PID {pid}."; }

                var win = FindWindowByPid(pid);
                if (win == null)
                    return $"Process PID={pid} ({proc.ProcessName}) found but has no visible window.\n" +
                           "Try ListWindows() to confirm.";

                _currentWindow = win;
                _currentHandle = $"w{++_windowCounter}";
                return $"Attached to '{_currentWindow.Title}' via PID {pid} ({proc.ProcessName}) — handle: {_currentHandle}\n" +
                       "Next: call GetSnapshot to see the accessibility tree.";
            }
            var desktop = Automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));

            // Collect all candidates with their match quality
            var candidates = new List<(Window win, string procName, int quality, string reason)>();
            bool useRegex  = title.StartsWith("regex:", StringComparison.OrdinalIgnoreCase);
            var pattern    = useRegex ? title[6..] : "";
            var searchLo   = title.ToLowerInvariant();
            var procFilter = processName.ToLowerInvariant();

            foreach (var w in windows)
            {
                Window? win;
                try { win = w.AsWindow(); } catch { continue; }
                if (win == null) continue;

                string pName = "";
                int    winPid = 0;
                try
                {
                    winPid = w.Properties.ProcessId.ValueOrDefault;
                    pName  = System.Diagnostics.Process.GetProcessById(winPid).ProcessName;
                }
                catch { }

                // Apply processName filter early
                if (!string.IsNullOrEmpty(procFilter) &&
                    !pName.ToLowerInvariant().Contains(procFilter)) continue;

                var winTitle = win.Title ?? "";

                // quality: 3=exact title, 2=exact proc name, 1=partial title, 0=partial proc
                int quality = 0;
                string reason = "";

                if (useRegex)
                {
                    if (!Regex.IsMatch(winTitle, pattern, RegexOptions.IgnoreCase)) continue;
                    quality = 3; reason = $"regex match on title \"{winTitle}\"";
                }
                else if (winTitle.Equals(title, StringComparison.OrdinalIgnoreCase))
                {
                    quality = 3; reason = $"exact title \"{winTitle}\"";
                }
                else if (pName.Equals(title, StringComparison.OrdinalIgnoreCase))
                {
                    quality = 2; reason = $"exact process \"{pName}\" (pid {winPid})";
                }
                else if (winTitle.Contains(searchLo, StringComparison.OrdinalIgnoreCase))
                {
                    // Skip IDE windows for partial matches
                    if (IsIdeProcess(pName)) continue;
                    quality = 1; reason = $"partial title \"{winTitle}\"";
                }
                else if (pName.Contains(searchLo, StringComparison.OrdinalIgnoreCase))
                {
                    if (IsIdeProcess(pName)) continue;
                    quality = 0; reason = $"partial process \"{pName}\" (pid {winPid})";
                }
                else continue;

                candidates.Add((win, pName, quality, reason));
            }

            if (candidates.Count == 0)
                return $"No window found matching title='{title}'" +
                       (string.IsNullOrEmpty(processName) ? "" : $" processName='{processName}'") +
                       ".\nTip: call ListWindows() to see all open windows, or use 'regex:^Raptor$' for exact match.";

            // Pick best match (highest quality; on tie prefer shorter title = less noise)
            var best = candidates
                .OrderByDescending(c => c.quality)
                .ThenBy(c => c.win.Title?.Length ?? 999)
                .First();

            if (candidates.Count > 1)
            {
                var others = string.Join(", ", candidates
                    .Where(c => c != best)
                    .Select(c => $"\"{c.win.Title}\" [{c.procName}]"));
                // Only warn if there were lower-quality or same-quality alternatives
                var msg = candidates.Any(c => c.quality == best.quality && c != best)
                    ? $"⚠️ Multiple windows matched — picked best: \"{best.win.Title}\" [{best.procName}].\n" +
                      $"Others: {others}\nUse processName= or 'regex:^exact$' to be precise.\n"
                    : "";
                _currentWindow = best.win;
                _currentHandle = $"w{++_windowCounter}";
                return $"{msg}Attached to '{_currentWindow.Title}' via {best.reason} — handle: {_currentHandle}\n" +
                       "Next: call GetSnapshot to see the accessibility tree.";
            }

            _currentWindow = best.win;
            _currentHandle = $"w{++_windowCounter}";
            return $"Attached to '{_currentWindow.Title}' via {best.reason} — handle: {_currentHandle}\n" +
                   "Next: call GetSnapshot to see the accessibility tree.";
        }
        catch (Exception ex)
        {
            return $"ERROR in AttachToApp: {ex.Message}";
        }
    }

    private static bool IsIdeProcess(string procName) =>
        KnownIdeProcesses.Contains(procName) ||
        procName.StartsWith("Code", StringComparison.OrdinalIgnoreCase) ||
        procName.StartsWith("devenv", StringComparison.OrdinalIgnoreCase);

    // -------------------------------------------------------------------------
    // 12. DEPLOY APP
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Deploy a packaged WinUI 3 MSIX app by registering its AppxManifest.xml. " +
        "Automatically kills any running instance of the app, syncs ALL build outputs " +
        "(DLL, EXE, XBF, PRI, etc.) into the AppX folder, then registers with Add-AppxPackage. " +
        "Returns the AUMID. " +
        "Typical flow: BuildApp → DeployApp → LaunchApp(aumid). " +
        "Or use BuildDeployLaunch for the full flow in a single call.")]
    public static string DeployApp(
        [Description("Full path to AppxManifest.xml, e.g. Z:\\source\\Raptor\\bin\\x64\\Debug\\net8.0-windows10.0.19041.0\\win-x64\\AppX\\AppxManifest.xml")]
        string manifestPath)
    {
        try
        {
            if (!File.Exists(manifestPath))
                return $"ERROR: AppxManifest.xml not found: {manifestPath}";

            // Auto-kill running instance so files aren't locked during sync
            var killMsg = KillPackageProcess(manifestPath);

            // Sync all output files into AppX — fail loudly if files are still locked
            var (syncMsg, hasLockedFiles) = SyncAppXFromManifest(manifestPath);
            if (hasLockedFiles)
                return $"❌ DEPLOY FAILED — files still locked after kill attempt.\n{syncMsg}\n" +
                       "Recovery: run CloseApp(), verify the process is gone, then retry DeployApp.";

            var aumid = DeployManifest(manifestPath, out var deployError);
            if (aumid == null) return deployError!;

            return $"✅ Package deployed successfully.\n" +
                   (killMsg.Length > 0 ? $"{killMsg}\n" : "") +
                   $"{syncMsg}\n" +
                   $"AUMID: {aumid}\n" +
                   $"Next: call LaunchApp(\"{aumid}\") to launch.";
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
    // 15. WAIT FOR WINDOW READY
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Wait until the current app window has rendered its initial UI. " +
        "Polls the accessibility tree until at least minElements are visible or maxWaitMs elapses. " +
        "Call this after LaunchApp when autoScreenshot=false, or before GetSnapshot to ensure the UI is loaded.")]
    public static string WaitForWindowReady(
        [Description("Maximum milliseconds to wait. Default 3000.")]
        int maxWaitMs = 3000,
        [Description("Minimum number of accessible elements before considering the window ready. Default 3.")]
        int minElements = 3)
    {
        try
        {
            EnsureWindowReady();
            var ready = WaitForWindowReadyInternal(maxWaitMs, minElements);
            return ready
                ? $"✅ Window ready — UI elements visible after ≤{maxWaitMs}ms."
                : $"⏱ Timed out after {maxWaitMs}ms. Window may still be loading. Proceed with GetSnapshot.";
        }
        catch (Exception ex) { return $"ERROR: {ex.Message}"; }
    }

    // -------------------------------------------------------------------------
    // 16. BUILD DEPLOY LAUNCH  (convenience — all three steps in one call)
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Convenience tool: runs BuildApp → DeployApp → LaunchApp in a single call. " +
        "Automatically finds the AppxManifest.xml from the project path, kills any running instance, " +
        "builds (incrementally by default), syncs AppX, deploys, launches, waits for the window to render, " +
        "and optionally captures a screenshot. " +
        "Use this instead of calling BuildApp + DeployApp + LaunchApp separately.")]
    public static IEnumerable<ContentBlock> BuildDeployLaunch(
        [Description("Full path to the .csproj file (e.g. Z:\\source\\Raptor\\Raptor.csproj)")]
        string projectPath,
        [Description("Build configuration: 'Debug' or 'Release' (default: Debug)")]
        string configuration = "Debug",
        [Description("Target platform: 'x64', 'ARM64', or 'x86' (default: x64 for packaged apps)")]
        string platform = "x64",
        [Description("Capture a screenshot after launch. Default true.")]
        bool screenshot = true,
        [Description("Force a full rebuild (/t:Rebuild). Default false = incremental.")]
        bool forceRebuild = false)
    {
        var log = new StringBuilder();
        try
        {
            if (!File.Exists(projectPath))
                return [new TextContentBlock { Text = $"ERROR: Project file not found: {projectPath}" }];

            // ---- Step 1: Build ----
            log.AppendLine("🔨 Building…");
            var buildResult = BuildApp(projectPath, configuration, platform, verbose: false, incremental: !forceRebuild);
            log.AppendLine(buildResult);
            if (buildResult.StartsWith("❌") || buildResult.StartsWith("ERROR"))
                return [new TextContentBlock { Text = log.ToString() }];

            // ---- Step 2: Find manifest ----
            var manifestPath = FindManifestPath(projectPath, configuration, platform);
            if (manifestPath == null)
            {
                log.AppendLine($"❌ Could not find AppxManifest.xml under bin\\{platform}\\{configuration}\\. " +
                               "Ensure the project has been built for a packaged MSIX target.");
                return [new TextContentBlock { Text = log.ToString() }];
            }
            log.AppendLine($"📄 Manifest: {manifestPath}");

            // ---- Step 3: Kill + Sync + Deploy ----
            log.AppendLine("🛑 Killing running instance (if any)…");
            var killMsg = KillPackageProcess(manifestPath);
            if (killMsg.Length > 0) log.AppendLine(killMsg);

            var (syncMsg, hasLocked) = SyncAppXFromManifest(manifestPath);
            log.AppendLine(syncMsg);
            if (hasLocked)
            {
                log.AppendLine("❌ Locked files detected. Aborting deploy.");
                return [new TextContentBlock { Text = log.ToString() }];
            }

            var aumid = DeployManifest(manifestPath, out var deployError);
            if (aumid == null)
            {
                log.AppendLine($"❌ Deploy failed: {deployError}");
                return [new TextContentBlock { Text = log.ToString() }];
            }
            log.AppendLine($"✅ Deployed. AUMID: {aumid}");

            // ---- Step 4: Launch ----
            log.AppendLine("🚀 Launching…");
            _currentWindow    = null;
            _elementRegistry.Clear();
            _elementCounter   = 0;
            _windowCounter    = 0;

            var launchResult = LaunchByAumid(aumid, forceRestart: true);
            log.AppendLine(launchResult);

            // ---- Step 5: Wait for ready + screenshot ----
            if (_currentWindow != null)
            {
                WaitForWindowReadyInternal(3000, 3);
                log.AppendLine($"✅ Window ready: '{_currentWindow.Title}' — handle: {_currentHandle}");

                if (screenshot)
                {
                    try
                    {
                        var hwnd = _currentWindow.Properties.NativeWindowHandle.ValueOrDefault;
                        using var bmp    = CaptureHwnd(hwnd);
                        using var stream = new MemoryStream();
                        bmp.Save(stream, ImageFormat.Png);
                        return
                        [
                            new TextContentBlock  { Text = log.ToString() },
                            new ImageContentBlock { Data = Convert.ToBase64String(stream.ToArray()), MimeType = "image/png" }
                        ];
                    }
                    catch (Exception ex) { log.AppendLine($"⚠️ Screenshot failed: {ex.Message}"); }
                }
            }

            return [new TextContentBlock { Text = log.ToString() }];
        }
        catch (Exception ex)
        {
            log.AppendLine($"ERROR: {ex.Message}");
            return [new TextContentBlock { Text = log.ToString() }];
        }
    }

    // -------------------------------------------------------------------------
    // 17. ANALYZE SCREENSHOT  (dominant colors + region breakdown)
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Analyze the current window screenshot: returns dominant colors (top 5 by %) " +
        "and a 3×3 grid breakdown of the most prominent color in each region. " +
        "Useful when comparing against a reference image or debugging layout/theming.")]
    public static string AnalyzeScreenshot()
    {
        try
        {
            EnsureWindowReady();
            var hwnd = _currentWindow!.Properties.NativeWindowHandle.ValueOrDefault;
            using var bmp = CaptureHwnd(hwnd);

            // Sample every 8th pixel for speed
            const int step = 8;
            const int buckets = 16; // per channel (0-255 → 16 buckets of 16)
            var colorCounts = new Dictionary<(int r, int g, int b), int>();

            // Overall dominant colors
            for (int y = 0; y < bmp.Height; y += step)
            for (int x = 0; x < bmp.Width;  x += step)
            {
                var px = bmp.GetPixel(x, y);
                var key = (px.R / buckets, px.G / buckets, px.B / buckets);
                colorCounts.TryGetValue(key, out int c);
                colorCounts[key] = c + 1;
            }

            var total   = colorCounts.Values.Sum();
            var topColors = colorCounts
                .OrderByDescending(kv => kv.Value)
                .Take(5)
                .Select(kv => {
                    var (r, g, b) = kv.Key;
                    int pct = (int)(kv.Value * 100.0 / total);
                    return $"  #{r*buckets:X2}{g*buckets:X2}{b*buckets:X2} — {pct}%";
                });

            // 3×3 grid dominant color per region
            int cols = 3, rows = 3;
            int cw = bmp.Width / cols, ch = bmp.Height / rows;
            var grid = new StringBuilder();
            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < cols; col++)
                {
                    var regionColors = new Dictionary<(int, int, int), int>();
                    for (int y = row * ch; y < (row + 1) * ch; y += step)
                    for (int x = col * cw; x < (col + 1) * cw; x += step)
                    {
                        var px  = bmp.GetPixel(x, y);
                        var key = (px.R / buckets, px.G / buckets, px.B / buckets);
                        regionColors.TryGetValue(key, out int c);
                        regionColors[key] = c + 1;
                    }
                    var best = regionColors.OrderByDescending(kv => kv.Value).First().Key;
                    grid.Append($"#{best.Item1 * buckets:X2}{best.Item2 * buckets:X2}{best.Item3 * buckets:X2}");
                    if (col < cols - 1) grid.Append(" | ");
                }
                grid.AppendLine(row < rows - 1 ? "\n  ---+---+---" : "");
            }

            var sb = new StringBuilder();
            sb.AppendLine($"=== Screenshot Analysis ({bmp.Width}×{bmp.Height}) ===");
            sb.AppendLine("\nDominant colors (approximate):");
            sb.AppendLine(string.Join("\n", topColors));
            sb.AppendLine("\n3×3 grid (dominant color per region):");
            sb.Append("  ");
            sb.Append(grid);
            return sb.ToString();
        }
        catch (Exception ex) { return $"ERROR analyzing screenshot: {ex.Message}"; }
    }

    // -------------------------------------------------------------------------
    // 18. READ FIGMA DESIGN
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "Read a Figma design link and extract everything needed to implement it in WinUI 3 XAML. " +
        "Returns an inline rendered PNG of the frame (so a vision model can see the design) plus " +
        "extracted design tokens: background/foreground colors, typography (font family, size, weight, " +
        "line height), spacing, padding, corner radius, borders, and layout direction. " +
        "Pass any Figma share URL — the file key and node ID are parsed automatically. " +
        "The Figma API token can be passed here or set once via the FIGMA_TOKEN environment variable.\n" +
        "Typical flow: ReadFigmaDesign(url) → edit XAML files → BuildDeployLaunch → CompareScreenshot.")]
    public static IEnumerable<ContentBlock> ReadFigmaDesign(
        [Description("Figma share URL, e.g. https://www.figma.com/design/abc123/MyApp?node-id=1-2.\n" +
                     "Use ListFigmaFrames first to discover available node IDs.")]
        string figmaUrl,
        [Description("Figma personal access token. If empty, reads from FIGMA_TOKEN environment variable.\n" +
                     "Generate one at https://www.figma.com/settings → Personal access tokens.")]
        string apiToken = "")
    {
        try
        {
            var token = ResolveToken(apiToken);
            if (token == null)
                return [new TextContentBlock { Text =
                    "ERROR: No Figma API token.\n" +
                    "• Pass apiToken= to this tool, OR\n" +
                    "• Set the FIGMA_TOKEN environment variable once (recommended).\n" +
                    "Generate a token at https://www.figma.com/settings → Personal access tokens." }];

            var (fileKey, nodeId) = ParseFigmaUrl(figmaUrl);
            if (fileKey == null)
                return [new TextContentBlock { Text =
                    $"ERROR: Could not parse Figma URL: {figmaUrl}\n" +
                    "Expected: https://www.figma.com/design/{{fileKey}}/..." }];

            using var http = MakeFigmaClient(token);

            // Fetch node or full file
            var apiUrl = nodeId != null
                ? $"https://api.figma.com/v1/files/{fileKey}/nodes?ids={Uri.EscapeDataString(nodeId)}"
                : $"https://api.figma.com/v1/files/{fileKey}";

            var resp = http.GetAsync(apiUrl).GetAwaiter().GetResult();
            if (!resp.IsSuccessStatusCode)
            {
                var body = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                return [new TextContentBlock { Text = $"ERROR: Figma API {(int)resp.StatusCode}: {body}" }];
            }

            var json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            using var doc = JsonDocument.Parse(json);

            // Locate the target node element
            JsonElement? targetNode = FindTargetNode(doc.RootElement, nodeId);
            if (targetNode == null)
                return [new TextContentBlock { Text =
                    $"ERROR: Could not locate node '{nodeId}' in Figma response.\n" +
                    "Call ListFigmaFrames to see available node IDs." }];

            var nodeName = targetNode.Value.TryGetProperty("name", out var nn) ? nn.GetString() ?? "" : "";
            var sb = new StringBuilder();
            sb.AppendLine($"=== Figma Design: \"{nodeName}\" ===");
            sb.AppendLine($"File: {fileKey}  Node: {nodeId ?? "(root)"}");
            sb.AppendLine($"URL:  {figmaUrl}");
            sb.AppendLine();
            ExtractFigmaTokens(targetNode.Value, sb, depth: 0, maxDepth: 4);
            sb.AppendLine();
            sb.AppendLine("─── XAML Implementation Notes ───────────────────────────────");
            sb.AppendLine("• Colors above use WinUI3 hex format: #AARRGGBB (with alpha) or #RRGGBB");
            sb.AppendLine("• FontWeight values map to WinUI3 FontWeights enum names");
            sb.AppendLine("• CornerRadius maps directly to <Border> or <Button> CornerRadius property");
            sb.AppendLine("• Padding values are Left,Top,Right,Bottom");
            sb.AppendLine("• After editing XAML, call BuildDeployLaunch(projectPath) to apply.");

            // Render the node as a PNG via Figma image API
            string? imageBase64 = null;
            try
            {
                var renderNodeId = nodeId ?? GetFirstFrameId(targetNode.Value);
                if (renderNodeId != null)
                {
                    var imgApiUrl = $"https://api.figma.com/v1/images/{fileKey}" +
                                   $"?ids={Uri.EscapeDataString(renderNodeId)}&format=png&scale=2";
                    var imgResp = http.GetAsync(imgApiUrl).GetAwaiter().GetResult();
                    if (imgResp.IsSuccessStatusCode)
                    {
                        var imgJson = imgResp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                        using var imgDoc = JsonDocument.Parse(imgJson);
                        if (imgDoc.RootElement.TryGetProperty("images", out var images))
                        {
                            foreach (var img in images.EnumerateObject())
                            {
                                var src = img.Value.GetString();
                                if (!string.IsNullOrEmpty(src))
                                {
                                    var bytes = http.GetByteArrayAsync(src).GetAwaiter().GetResult();
                                    imageBase64 = Convert.ToBase64String(bytes);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { sb.AppendLine($"\n⚠️ Could not render Figma image: {ex.Message}"); }

            var results = new List<ContentBlock> { new TextContentBlock { Text = sb.ToString() } };
            if (imageBase64 != null)
                results.Add(new ImageContentBlock { Data = imageBase64, MimeType = "image/png" });
            return results;
        }
        catch (Exception ex)
        {
            return [new TextContentBlock { Text = $"ERROR reading Figma design: {ex.Message}" }];
        }
    }

    // -------------------------------------------------------------------------
    // 19. LIST FIGMA FRAMES
    // -------------------------------------------------------------------------

    [McpServerTool, Description(
        "List all pages and top-level frames/components in a Figma file. " +
        "Use this first to discover node IDs before calling ReadFigmaDesign with a specific frame. " +
        "The Figma API token can be passed here or set as FIGMA_TOKEN environment variable.")]
    public static string ListFigmaFrames(
        [Description("Figma share URL, e.g. https://www.figma.com/design/abc123/MyApp")]
        string figmaUrl,
        [Description("Figma personal access token. If empty, reads from FIGMA_TOKEN environment variable.")]
        string apiToken = "")
    {
        try
        {
            var token = ResolveToken(apiToken);
            if (token == null)
                return "ERROR: No Figma API token. Pass apiToken= or set FIGMA_TOKEN environment variable.\n" +
                       "Generate one at https://www.figma.com/settings → Personal access tokens.";

            var (fileKey, _) = ParseFigmaUrl(figmaUrl);
            if (fileKey == null)
                return $"ERROR: Could not parse Figma URL: {figmaUrl}";

            using var http = MakeFigmaClient(token);

            var resp = http.GetAsync($"https://api.figma.com/v1/files/{fileKey}?depth=2")
                           .GetAwaiter().GetResult();
            if (!resp.IsSuccessStatusCode)
                return $"ERROR: Figma API {(int)resp.StatusCode}";

            var json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            using var doc = JsonDocument.Parse(json);

            var fileName = doc.RootElement.TryGetProperty("name", out var fn) ? fn.GetString() : fileKey;
            var sb = new StringBuilder();
            sb.AppendLine($"=== Figma File: \"{fileName}\" ===");
            sb.AppendLine($"Key: {fileKey}");
            sb.AppendLine();

            if (doc.RootElement.TryGetProperty("document", out var document) &&
                document.TryGetProperty("children", out var pages))
            {
                foreach (var page in pages.EnumerateArray())
                {
                    var pageName = page.TryGetProperty("name", out var pn) ? pn.GetString() : "?";
                    var pageId   = page.TryGetProperty("id",   out var pi) ? pi.GetString() : "?";
                    sb.AppendLine($"📄 Page: \"{pageName}\"  (node-id: {pageId})");

                    if (page.TryGetProperty("children", out var frames))
                    {
                        foreach (var frame in frames.EnumerateArray())
                        {
                            var frameName = frame.TryGetProperty("name", out var frn) ? frn.GetString() : "?";
                            var frameId   = frame.TryGetProperty("id",   out var fri) ? fri.GetString() : "?";
                            var frameType = frame.TryGetProperty("type", out var frt) ? frt.GetString() : "?";

                            // Compute size hint
                            var sizeHint = "";
                            if (frame.TryGetProperty("absoluteBoundingBox", out var bbb))
                            {
                                var w = bbb.TryGetProperty("width",  out var bw) ? bw.GetDouble() : 0;
                                var h = bbb.TryGetProperty("height", out var bh) ? bh.GetDouble() : 0;
                                sizeHint = $"  [{w:F0}×{h:F0}]";
                            }

                            sb.AppendLine($"  [{frameType}] \"{frameName}\"{sizeHint}");
                            sb.AppendLine($"    node-id: {frameId}");

                            // Build the URL with node-id so user can click it
                            var frameUrl = $"{figmaUrl.Split('?')[0]}?node-id={Uri.EscapeDataString(frameId)}";
                            sb.AppendLine($"    → ReadFigmaDesign(\"{frameUrl}\")");
                        }
                    }
                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }
        catch (Exception ex) { return $"ERROR listing frames: {ex.Message}"; }
    }

    // -------------------------------------------------------------------------
    // PRIVATE HELPERS
    // -------------------------------------------------------------------------

    // Derive AppxManifest.xml path from project path, platform, and configuration
    private static string? FindManifestPath(string projectPath, string configuration, string platform)
    {
        var projectDir = Path.GetDirectoryName(projectPath)!;
        var binDir     = Path.Combine(projectDir, "bin", platform, configuration);
        if (!Directory.Exists(binDir)) return null;
        return Directory.GetFiles(binDir, "AppxManifest.xml", SearchOption.AllDirectories)
                        .FirstOrDefault();
    }

    // Poll accessibility tree until minElements visible or timeout
    private static bool WaitForWindowReadyInternal(int maxWaitMs, int minElements)
    {
        if (_currentWindow == null) return false;
        var deadline = Environment.TickCount64 + maxWaitMs;
        while (Environment.TickCount64 < deadline)
        {
            try
            {
                var count = _currentWindow.FindAllDescendants().Length;
                if (count >= minElements) return true;
            }
            catch { }
            Thread.Sleep(200);
        }
        return false;
    }

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

    // Kill the running process for a package (by matching install location to process EXE path)
    private static string KillPackageProcess(string manifestPath)
    {
        try
        {
            var appxDir = Path.GetDirectoryName(manifestPath)!;
            var killed  = new List<string>();

            foreach (var proc in System.Diagnostics.Process.GetProcesses())
            {
                try
                {
                    var exe = proc.MainModule?.FileName ?? "";
                    if (exe.StartsWith(appxDir, StringComparison.OrdinalIgnoreCase))
                    {
                        proc.Kill();
                        proc.WaitForExit(3000);
                        killed.Add($"{proc.ProcessName} (pid {proc.Id})");
                    }
                }
                catch { }
            }

            if (killed.Count > 0)
            {
                Thread.Sleep(1000); // Let the OS release file handles
                return $"🛑 Killed: {string.Join(", ", killed)}";
            }
            return "";
        }
        catch { return ""; }
    }

    // Sync ALL changed output files from the directory containing AppX into AppX itself.
    // Returns (message, hasLockedFiles). Locked files = partial failure.
    private static (string message, bool hasLockedFiles) SyncAppXFromManifest(string manifestPath)
    {
        try
        {
            var appxDir   = Path.GetDirectoryName(manifestPath)!;
            var outputDir = Path.GetDirectoryName(appxDir)!;
            var synced    = new List<string>();
            var locked    = new List<string>();

            var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { ".dll", ".exe", ".xbf", ".pri", ".runtimeconfig.json", ".deps.json" };

            foreach (var srcFile in Directory.GetFiles(outputDir))
            {
                if (!exts.Contains(Path.GetExtension(srcFile))) continue;
                var fname   = Path.GetFileName(srcFile);
                var dstFile = Path.Combine(appxDir, fname);
                if (!File.Exists(dstFile)) continue;
                try
                {
                    if (File.GetLastWriteTimeUtc(srcFile) > File.GetLastWriteTimeUtc(dstFile))
                    {
                        File.Copy(srcFile, dstFile, overwrite: true);
                        synced.Add(fname);
                    }
                }
                catch { locked.Add(fname); }
            }

            if (locked.Count > 0)
                return ($"⚠️ {locked.Count} file(s) locked — app is still running: {string.Join(", ", locked)}\n" +
                        "Fix: call CloseApp() first, or use BuildDeployLaunch which handles this automatically.", true);
            if (synced.Count > 0)
                return ($"📦 Synced {synced.Count} file(s): {string.Join(", ", synced)}", false);
            return ("📦 AppX already up to date", false);
        }
        catch (Exception ex) { return ($"⚠️ Sync warning: {ex.Message}", false); }
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
            return File.Exists(manifestPath) ? SyncAppXFromManifest(manifestPath).message : "";
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

    // ─── Figma helpers ────────────────────────────────────────────────────────

    private static string? ResolveToken(string apiToken) =>
        string.IsNullOrWhiteSpace(apiToken)
            ? (Environment.GetEnvironmentVariable("FIGMA_TOKEN") is { Length: > 0 } t ? t : null)
            : apiToken;

    private static HttpClient MakeFigmaClient(string token)
    {
        var http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        http.DefaultRequestHeaders.Add("X-Figma-Token", token);
        return http;
    }

    // Parse https://www.figma.com/design/KEY/Title?node-id=1-2
    private static (string? fileKey, string? nodeId) ParseFigmaUrl(string url)
    {
        try
        {
            var uri      = new Uri(url);
            var segments = uri.Segments;
            string? fileKey = null;

            for (int i = 0; i < segments.Length - 1; i++)
            {
                var seg = segments[i].Trim('/');
                if (seg is "file" or "design" or "proto" or "board")
                {
                    fileKey = segments[i + 1].Trim('/');
                    break;
                }
            }

            // Parse query string for node-id
            string? nodeId = null;
            foreach (var part in uri.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
            {
                var kv = part.Split('=', 2);
                if (kv.Length == 2 && kv[0] == "node-id")
                {
                    nodeId = Uri.UnescapeDataString(kv[1]);
                    break;
                }
            }

            return (fileKey, nodeId);
        }
        catch { return (null, null); }
    }

    // Locate the relevant JsonElement from Figma API response
    private static JsonElement? FindTargetNode(JsonElement root, string? nodeId)
    {
        // GET /files/{key}/nodes response → { "nodes": { "1:2": { "document": {...} } } }
        if (nodeId != null && root.TryGetProperty("nodes", out var nodes))
        {
            // Try both "1-2" and "1:2" format
            foreach (var candidate in new[] { nodeId, nodeId.Replace("-", ":"), nodeId.Replace(":", "-") })
            {
                if (nodes.TryGetProperty(candidate, out var wrapper) &&
                    wrapper.TryGetProperty("document", out var doc))
                    return doc;
            }
        }
        // GET /files/{key} response → { "document": {...} }
        if (root.TryGetProperty("document", out var rootDoc))
            return rootDoc;
        return null;
    }

    private static string? GetFirstFrameId(JsonElement node)
    {
        if (node.TryGetProperty("id", out var id)) return id.GetString();
        if (node.TryGetProperty("children", out var ch))
            foreach (var child in ch.EnumerateArray())
                if (child.TryGetProperty("id", out var cid)) return cid.GetString();
        return null;
    }

    // Recursively extract WinUI3-relevant tokens from a Figma node tree
    private static void ExtractFigmaTokens(JsonElement node, StringBuilder sb, int depth, int maxDepth)
    {
        if (depth > maxDepth) return;

        var indent   = new string(' ', depth * 2);
        var name     = node.TryGetProperty("name", out var n)   ? n.GetString() ?? "" : "";
        var type     = node.TryGetProperty("type", out var t)   ? t.GetString() ?? "" : "";
        var nodeId   = node.TryGetProperty("id",   out var nid) ? nid.GetString() ?? "" : "";

        if (depth == 0)
            sb.AppendLine("─── Node Tree & Tokens ──────────────────────────────────────");

        // Size
        var sizeHint = "";
        if (node.TryGetProperty("absoluteBoundingBox", out var bb))
        {
            var w = bb.TryGetProperty("width",  out var bw) ? bw.GetDouble() : 0;
            var h = bb.TryGetProperty("height", out var bh) ? bh.GetDouble() : 0;
            sizeHint = $"  {w:F0}×{h:F0}px";
        }
        sb.AppendLine($"{indent}[{type}] \"{name}\" (id: {nodeId}){sizeHint}");

        // Fills → Background / Foreground
        if (node.TryGetProperty("fills", out var fills) && fills.GetArrayLength() > 0)
        {
            foreach (var fill in fills.EnumerateArray())
            {
                if (!fill.TryGetProperty("visible", out var vis) || vis.GetBoolean()) // skip hidden
                if (fill.TryGetProperty("type", out var ft) && ft.GetString() == "SOLID" &&
                    fill.TryGetProperty("color", out var color))
                {
                    var hex  = FigmaColorToHex(color, fill);
                    var prop = type == "TEXT" ? "Foreground" : "Background";
                    sb.AppendLine($"{indent}  {prop}=\"{hex}\"");
                }
            }
        }

        // Strokes → BorderBrush + BorderThickness
        if (node.TryGetProperty("strokes", out var strokes) && strokes.GetArrayLength() > 0)
        {
            foreach (var stroke in strokes.EnumerateArray())
            {
                if (stroke.TryGetProperty("type", out var st) && st.GetString() == "SOLID" &&
                    stroke.TryGetProperty("color", out var sc))
                {
                    var hex = FigmaColorToHex(sc, stroke);
                    var sw  = node.TryGetProperty("strokeWeight", out var sw_) ? sw_.GetDouble() : 1;
                    sb.AppendLine($"{indent}  BorderBrush=\"{hex}\"  BorderThickness=\"{sw:F0}\"");
                }
            }
        }

        // Corner radius
        if (node.TryGetProperty("cornerRadius", out var cr) && cr.GetDouble() > 0)
            sb.AppendLine($"{indent}  CornerRadius=\"{cr.GetDouble():F0}\"");
        else if (node.TryGetProperty("rectangleCornerRadii", out var rcr) && rcr.GetArrayLength() == 4)
        {
            var r = rcr.EnumerateArray().Select(v => v.GetDouble().ToString("F0")).ToArray();
            sb.AppendLine($"{indent}  CornerRadius=\"{r[0]},{r[1]},{r[2]},{r[3]}\"");
        }

        // Opacity
        if (node.TryGetProperty("opacity", out var op) && Math.Abs(op.GetDouble() - 1.0) > 0.01)
            sb.AppendLine($"{indent}  Opacity=\"{op.GetDouble():F2}\"");

        // Auto-layout → StackPanel / Grid
        if (node.TryGetProperty("layoutMode", out var lm) && lm.GetString() is "HORIZONTAL" or "VERTICAL")
        {
            var dir = lm.GetString() == "HORIZONTAL" ? "Horizontal" : "Vertical";
            sb.AppendLine($"{indent}  Layout: StackPanel Orientation=\"{dir}\"");

            var gap = node.TryGetProperty("itemSpacing", out var g) ? g.GetDouble() : 0;
            if (gap > 0) sb.AppendLine($"{indent}  Spacing=\"{gap:F0}\"");

            var pl = node.TryGetProperty("paddingLeft",   out var p1) ? p1.GetDouble() : 0;
            var pr = node.TryGetProperty("paddingRight",  out var p2) ? p2.GetDouble() : 0;
            var pt = node.TryGetProperty("paddingTop",    out var p3) ? p3.GetDouble() : 0;
            var pb = node.TryGetProperty("paddingBottom", out var p4) ? p4.GetDouble() : 0;
            if (pl + pr + pt + pb > 0)
                sb.AppendLine($"{indent}  Padding=\"{pl:F0},{pt:F0},{pr:F0},{pb:F0}\"");

            var primary = node.TryGetProperty("primaryAxisAlignItems",  out var pa) ? pa.GetString() : null;
            var counter = node.TryGetProperty("counterAxisAlignItems",  out var ca) ? ca.GetString() : null;
            if (primary != null) sb.AppendLine($"{indent}  HorizontalAlignment hint: {primary}");
            if (counter != null) sb.AppendLine($"{indent}  VerticalAlignment hint:   {counter}");
        }

        // Typography (TEXT nodes)
        if (type == "TEXT")
        {
            if (node.TryGetProperty("characters", out var chars))
                sb.AppendLine($"{indent}  Text=\"{chars.GetString()}\"");

            if (node.TryGetProperty("style", out var style))
            {
                if (style.TryGetProperty("fontFamily",   out var ff))  sb.AppendLine($"{indent}  FontFamily=\"{ff.GetString()}\"");
                if (style.TryGetProperty("fontSize",     out var fs))  sb.AppendLine($"{indent}  FontSize=\"{fs.GetDouble():F0}\"");
                if (style.TryGetProperty("fontWeight",   out var fw))
                {
                    var w = fw.GetInt32();
                    sb.AppendLine($"{indent}  FontWeight=\"{FigmaWeightToXaml(w)}\"  ({w})");
                }
                if (style.TryGetProperty("lineHeightPx",  out var lh) && lh.GetDouble() > 0)
                    sb.AppendLine($"{indent}  LineHeight=\"{lh.GetDouble():F1}\"");
                if (style.TryGetProperty("letterSpacing",  out var ls) && Math.Abs(ls.GetDouble()) > 0.01)
                    sb.AppendLine($"{indent}  CharacterSpacing=\"{ls.GetDouble():F1}\"");
                if (style.TryGetProperty("textAlignHorizontal", out var ta))
                {
                    var align = ta.GetString() switch
                    {
                        "LEFT"     => "Left",
                        "CENTER"   => "Center",
                        "RIGHT"    => "Right",
                        "JUSTIFIED"=> "Justify",
                        _          => ta.GetString()
                    };
                    sb.AppendLine($"{indent}  TextAlignment=\"{align}\"");
                }
                if (style.TryGetProperty("italic", out var it) && it.GetBoolean())
                    sb.AppendLine($"{indent}  FontStyle=\"Italic\"");
                if (style.TryGetProperty("textDecoration", out var td) && td.GetString() != "NONE")
                    sb.AppendLine($"{indent}  TextDecorations=\"{td.GetString()}\"");
            }
        }

        // Recurse into children (cap at 12 per level to avoid noise)
        if (depth < maxDepth && node.TryGetProperty("children", out var children))
        {
            var total = children.GetArrayLength();
            int count = 0;
            foreach (var child in children.EnumerateArray())
            {
                if (++count > 12)
                {
                    sb.AppendLine($"{indent}  … ({total - 12} more children not shown)");
                    break;
                }
                ExtractFigmaTokens(child, sb, depth + 1, maxDepth);
            }
        }
    }

    // Figma color {r,g,b,a} (0-1 range) → WinUI3 hex string (#RRGGBB or #AARRGGBB)
    private static string FigmaColorToHex(JsonElement color, JsonElement paint)
    {
        var r  = color.TryGetProperty("r", out var rv) ? rv.GetDouble() : 0;
        var g  = color.TryGetProperty("g", out var gv) ? gv.GetDouble() : 0;
        var b  = color.TryGetProperty("b", out var bv) ? bv.GetDouble() : 0;
        var a  = color.TryGetProperty("a", out var av) ? av.GetDouble() : 1.0;
        if (paint.TryGetProperty("opacity", out var po)) a *= po.GetDouble();

        int ri = (int)Math.Round(r * 255);
        int gi = (int)Math.Round(g * 255);
        int bi = (int)Math.Round(b * 255);
        int ai = (int)Math.Round(a * 255);

        return a < 0.999 ? $"#{ai:X2}{ri:X2}{gi:X2}{bi:X2}" : $"#{ri:X2}{gi:X2}{bi:X2}";
    }

    // Figma numeric weight → WinUI3 FontWeights name
    private static string FigmaWeightToXaml(int weight) => weight switch
    {
        <= 150 => "Thin",
        <= 250 => "ExtraLight",
        <= 325 => "Light",
        <= 375 => "SemiLight",
        <= 450 => "Normal",
        <= 550 => "Medium",
        <= 650 => "SemiBold",
        <= 750 => "Bold",
        <= 850 => "ExtraBold",
        _      => "Black"
    };
}
