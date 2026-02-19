# WinUI3 MCP Server

A **Model Context Protocol (MCP) server** written in C# (.NET 8) that gives AI coding agents
(like OpenCode) the ability to **launch, capture, inspect, and interact with WinUI 3 Windows
applications** — the same way Playwright does for web browsers.

---

## What This Is

When you are developing a WinUI 3 app and run into UI issues (broken layout, wrong colors,
elements not showing, wrong sizing, etc.), normally you have to:

1. Run the app manually
2. Look at the screen yourself
3. Guess what XAML change fixes it
4. Rebuild and repeat

With this MCP server, your AI agent (OpenCode) can do all of that for you:

1. Launch your app automatically
2. Take a screenshot and **see** the UI
3. Inspect the full UI element tree
4. Find the broken element
5. Read your XAML source code
6. Apply the fix directly

This is the **same concept as Playwright MCP** — but for native Windows desktop apps.

---

## How It Works (Architecture)

```
OpenCode (AI Agent)
       |
       | MCP Protocol (stdio / JSON-RPC)
       |
WinUI3McpServer.exe  (this project)
       |
       | WebDriver protocol (HTTP)
       |
WinAppDriver.exe  (Microsoft - runs on port 4723)
       |
       | Windows UI Automation
       |
Your WinUI 3 App
```

**Key technologies:**

| Component | Role |
|-----------|------|
| [Model Context Protocol](https://modelcontextprotocol.io) | Lets AI agents call tools |
| [WinAppDriver](https://github.com/microsoft/WinAppDriver) | Microsoft's automation driver for Windows apps |
| [Appium.WebDriver](https://github.com/appium/dotnet-client) | .NET client for talking to WinAppDriver |
| [Microsoft UI Automation](https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/ui-automation-overview) | Windows API for reading UI element trees |

---

## Prerequisites

### 1. Enable Developer Mode in Windows

```
Settings > Privacy & Security > For Developers > Developer Mode  ->  ON
```

This is required by WinAppDriver to automate apps.

### 2. Install WinAppDriver

Download and install from the official Microsoft repository:

```
https://github.com/microsoft/WinAppDriver/releases
```

Install version **1.2.1** or later. The installer puts it at:

```
C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe
```

### 3. Install .NET 8 SDK

```
https://dotnet.microsoft.com/download/dotnet/8.0
```

Verify with:

```
dotnet --version
```

### 4. Have OpenCode installed

```
npm install -g opencode-ai
```

---

## Project Structure

```
Z:\openCode contribution\
└── WinUI3-MCP\
    └── WinUI3McpServer\
        ├── Program.cs              # MCP server entry point / host setup
        ├── WinUI3Tools.cs          # All 8 MCP tool implementations
        ├── WinUI3McpServer.csproj  # Project file with NuGet dependencies
        └── WinUI3McpServer.sln     # Solution file
```

---

## MCP Tools Reference

The server exposes **8 tools** to the AI agent:

### `LaunchApp`
Launches your WinUI 3 application.

| Parameter | Type | Description |
|-----------|------|-------------|
| `appPath` | string | Full path to `.exe` OR AUMID of a packaged app |
| `workingDirectory` | string | Optional working directory |

**Example prompts:**
```
use winui3 to launch my app at C:\Projects\MyApp\bin\Debug\net8.0-windows\MyApp.exe
```
```
launch MyApp_1.0.0.0_x64__abc123!App using the winui3 tool
```

---

### `CaptureScreenshot`
Takes a screenshot of the running app window. Returns a Base64 PNG the AI can analyze.

| Parameter | Type | Description |
|-----------|------|-------------|
| `saveToPath` | string | Optional file path to save the PNG |

**Example prompts:**
```
use winui3 to take a screenshot and describe what UI issues you see
```
```
capture a screenshot and save it to C:\screenshots\before.png using winui3
```

---

### `GetUiTree`
Returns the full UI element tree — every control, its type, name, visibility, and enabled state.

| Parameter | Type | Description |
|-----------|------|-------------|
| `maxDepth` | int | Tree depth (1–10, default 5) |

**Example prompts:**
```
use winui3 to get the UI tree and find why the Save button is not visible
```

---

### `FindElement`
Finds a specific UI element and returns its properties.

| Parameter | Type | Description |
|-----------|------|-------------|
| `strategy` | string | `id`, `name`, `xpath`, or `class` |
| `value` | string | The identifier value |

**Example prompts:**
```
use winui3 to find the element with id 'btnSave' and tell me its size and position
```

---

### `ClickElement`
Clicks a UI element. Useful for navigating to different pages to test different UI states.

| Parameter | Type | Description |
|-----------|------|-------------|
| `strategy` | string | `id`, `name`, `xpath`, or `class` |
| `value` | string | The identifier value |

**Example prompts:**
```
use winui3 to click the Settings button then take a screenshot
```

---

### `TypeText`
Types text into a TextBox or input field.

| Parameter | Type | Description |
|-----------|------|-------------|
| `strategy` | string | `id`, `name`, `xpath`, or `class` |
| `value` | string | The identifier value |
| `text` | string | Text to type |

**Example prompts:**
```
use winui3 to type 'hello world' into the element with id 'txtInput'
```

---

### `GetWindowInfo`
Returns window title, size, and position.

**Example prompts:**
```
use winui3 to get the current window info
```

---

### `CloseApp`
Closes the app and ends the WinAppDriver session.

**Example prompts:**
```
use winui3 to close the app
```

---

## Setup & Build

### Step 1 — Clone or open the project

```
cd "Z:\openCode contribution\WinUI3-MCP\WinUI3McpServer"
```

### Step 2 — Restore and build

```
dotnet build --configuration Release
```

Expected output:
```
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

### Step 3 — Register with OpenCode

The `opencode.jsonc` config file at `C:\Users\<YourUsername>\opencode.jsonc` should contain:

```jsonc
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "winui3": {
      "type": "local",
      "command": [
        "dotnet",
        "run",
        "--project",
        "Z:\\openCode contribution\\WinUI3-MCP\\WinUI3McpServer\\WinUI3McpServer.csproj",
        "--configuration",
        "Release"
      ],
      "enabled": true
    }
  }
}
```

> This file was already created for you at `C:\Users\Kathi\opencode.jsonc`.

---

## Running the Full Workflow

### Step 1 — Start WinAppDriver as Administrator

Open an elevated terminal and run:

```
"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe"
```

You should see:

```
Windows Application Driver listening for requests at: http://127.0.0.1:4723/
```

Leave this terminal open.

### Step 2 — Start OpenCode

In your project directory:

```
opencode
```

### Step 3 — Ask OpenCode to fix your UI

```
Use the winui3 tool to:
1. Launch my app at C:\Projects\MyApp\bin\Debug\net8.0-windows\MyApp.exe
2. Take a screenshot
3. Get the UI element tree
4. Look at my XAML files and fix whatever layout issues you see
```

OpenCode will use the MCP tools to see your app, diagnose the problem, and edit your XAML.

---

## Example End-to-End Scenario

**Problem:** Your WinUI 3 app has a Button that is off-screen on smaller window sizes.

**What you say to OpenCode:**
```
My Save button disappears when the window is small.
Use winui3 to launch the app, resize the window to 800x600,
take a screenshot, find the Save button element, and fix the XAML layout.
```

**What OpenCode does:**
1. Calls `LaunchApp` → opens your app
2. Calls `CaptureScreenshot` → sees the current state
3. Calls `GetUiTree` → finds all elements including the missing button
4. Calls `FindElement` with `id=btnSave` → checks its bounds
5. Reads your XAML file → identifies the layout issue
6. Edits the XAML → fixes the `Grid` row/column definitions or adds `ScrollViewer`

---

## Finding Your App's AUMID (Packaged Apps)

If your WinUI 3 app is packaged (MSIX), use PowerShell to find the AUMID:

```powershell
Get-StartApps | Where-Object { $_.Name -like "*YourAppName*" }
```

The `AppID` column is the AUMID. Use that as the `appPath` parameter in `LaunchApp`.

---

## NuGet Packages Used

| Package | Version | Purpose |
|---------|---------|---------|
| `ModelContextProtocol` | 0.8.0-preview.1 | MCP server SDK |
| `Microsoft.Extensions.Hosting` | 10.0.3 | .NET generic host / DI |
| `Appium.WebDriver` | 8.1.0 | WinAppDriver client |
| `Newtonsoft.Json` | 13.0.4 | JSON serialization |

---

## Troubleshooting

### "ERROR launching app: A exception with a null response was thrown..."
WinAppDriver is not running. Start it as Administrator first.

### "ERROR launching app: Developer Mode is not enabled"
Go to `Settings > Privacy & Security > For Developers` and turn on Developer Mode.

### "Element not found"
Use `GetUiTree` first to see all available elements and find the correct `AutomationId` or `Name`.

### App launches but screenshot is blank/black
Some WinUI 3 apps use hardware-accelerated rendering which can produce black screenshots.
Try adding `x:Name` attributes to your XAML elements and use `GetUiTree` instead.

### OpenCode does not show winui3 tools
Make sure `opencode.jsonc` is in your home directory (`C:\Users\<YourUsername>\`) and the
path to the `.csproj` file is correct with double backslashes.

---

## Future Improvements / Discussion Points

This project was built as a foundation. Below are areas open for contribution and debate:

### Possible Enhancements

- **Visual diff tool** — Compare two screenshots (before/after fix) and highlight changes
- **XAML Hot Reload integration** — Apply XAML edits without restarting the app
- **Accessibility audit tool** — Check for missing `AutomationProperties.Name` on controls
- **Multi-window support** — Handle apps with multiple windows or dialogs
- **Video recording** — Record a session as a `.mp4` for bug reports
- **Element highlight** — Draw a red border around a found element in the screenshot
- **Theme testing** — Switch between Light/Dark/High Contrast themes and capture each
- **Responsive testing** — Resize the window to multiple sizes and screenshot each

### Open Questions for Debate

1. **Should this be a standalone NuGet package** so other developers can add it to their
   own MCP setups without cloning this repo?

2. **WinAppDriver vs Windows UI Automation API directly** — WinAppDriver is officially
   archived (last release 2020). Should this be rewritten using the
   `UIAutomationClient` COM API directly for better WinUI 3 support?

3. **Should screenshot analysis use vision models specifically** (GPT-4o, Claude) rather
   than letting the general agent decide? A dedicated vision step might give better results.

4. **Integration with Visual Studio's XAML Hot Reload** — Is it possible to trigger a
   hot reload from the MCP server after editing XAML, so the agent can verify its fix
   visually without a full rebuild?

5. **AUMID discovery tool** — Should there be a `ListInstalledApps` tool that returns
   all installed WinUI 3 apps and their AUMIDs automatically?

---

## License

MIT — free to use, modify, and contribute.

---

## Related Resources

- [WinAppDriver on GitHub](https://github.com/microsoft/WinAppDriver)
- [Model Context Protocol Docs](https://modelcontextprotocol.io)
- [OpenCode Docs](https://opencode.ai/docs)
- [WinUI 3 Documentation](https://learn.microsoft.com/en-us/windows/apps/winui/winui3/)
- [Windows UI Automation Overview](https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/ui-automation-overview)
- [Appium .NET Client](https://github.com/appium/dotnet-client)
"# WinUI_MCP" 
