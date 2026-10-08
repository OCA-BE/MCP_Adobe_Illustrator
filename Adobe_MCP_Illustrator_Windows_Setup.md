# Adobe MCP + Illustrator Setup Guide for Windows
## Edge AI Agent → Adobe Illustrator (Windows)

> **Platform:** Windows 10/11 with Adobe Illustrator 2024/2025/2026  
> **Tested on:** Windows 11, Adobe Illustrator 2026 (v30.0.0)  
> **Author:** IBM Consulting — documented from live setup session

---

## Overview

This guide explains how to connect Edge's AI agent to Adobe Illustrator on Windows so you can control Illustrator via natural language — drawing shapes, adding text, exporting PDFs, and more.

### Architecture

```
Edge AI Agent (Linux container)
        │
        │ HTTP (MCP Streamable HTTP)
        ▼
supergateway (Node.js) ← port 8787 on Windows host
        │
        │ stdio (MCP protocol)
        ▼
Python MCP Server (illustrator-mcp)
        │
        │ File-based IPC (C:\Temp\ai_input.jsx / ai_output.txt)
        ▼
Python Worker (ai_worker.py) ← runs as interactive foreground process
        │
        │ COM Automation (win32com)
        ▼
Adobe Illustrator (running on Windows desktop)
```

### Why the Worker Process?

Windows COM automation of desktop applications (like Illustrator) only works from an **interactive foreground process** — not from a subprocess chain. Since Edge runs inside a Linux container and spawns the MCP server via supergateway, a direct COM call from that chain is blocked by Windows session isolation.

The solution: a lightweight **worker script** runs directly in a terminal with full interactive access, bridging COM calls via temp files in `C:\Temp`.

---

## Prerequisites

| Requirement | Version |
|---|---|
| Windows | 10 or 11 (64-bit) |
| Adobe Illustrator | 2024, 2025, or 2026 |
| Python | 3.10+ |
| Node.js + npm | 18+ |
| Git | Any recent version |

Verify installs:
```powershell
python --version
node --version
npm --version
git --version
```

---

## Step 1 — Clone the Illustrator MCP Server

Open **PowerShell** (regular, not Admin) and run:

```powershell
cd C:\Users\$env:USERNAME
git clone https://github.com/krVatsal/illustrator-mcp.git
cd illustrator-mcp
```

> ⚠️ Do NOT clone into `C:\Windows\System32` — use your user home directory.

---

## Step 2 — Create the Python Virtual Environment

```powershell
cd C:\Users\$env:USERNAME\illustrator-mcp

# Create venv
python -m venv .venv

# Install dependencies using python -m pip (NOT pip.exe — avoids launcher path issues)
.venv\Scripts\python.exe -m pip install "mcp<1.0"
.venv\Scripts\python.exe -m pip install pywin32
.venv\Scripts\python.exe -m pip install -e .

# Run pywin32 post-install to register COM components
.venv\Scripts\python.exe .venv\Scripts\pywin32_postinstall.py -install
```

> **Why `mcp<1.0`?** The illustrator-mcp server uses the old `@server.list_tools()` API which was removed in mcp 1.0+. Pinning to 0.9.x keeps it working.

---

## Step 3 — Patch platform_backend.py (Windows COM Fix)

The default server uses `DoJavaScriptFile` via direct Python COM, which fails when called from a subprocess chain (Edge container → supergateway → Python). We replace it with a **file-based IPC** pattern that delegates COM calls to a separate interactive worker process.

Open **PowerShell (Admin)** and run:

```powershell
# Take ownership of the file
$file = "C:\Users\$env:USERNAME\illustrator-mcp\illustrator\platform_backend.py"
takeown /f $file
icacls $file /grant "AzureAD\${env:USERNAME}:(F)" 2>$null
icacls $file /grant "${env:USERNAME}:(F)" 2>$null
```

Then write the patched backend:

```powershell
$content = @'
import abc, base64, io, logging, os, subprocess, sys, tempfile, time
from PIL import Image
logger = logging.getLogger(__name__)

class IllustratorBackend(abc.ABC):
    @abc.abstractmethod
    def focus_app(self): pass
    @abc.abstractmethod
    def capture_screenshot(self) -> str: pass
    @abc.abstractmethod
    def run_script(self, code: str) -> str: pass
    @staticmethod
    def _image_to_base64_jpeg(img, quality=50):
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        return base64.b64encode(buf.getvalue()).decode("utf-8")

class WindowsBackend(IllustratorBackend):
    def focus_app(self): pass
    def capture_screenshot(self):
        time.sleep(0.5)
        from PIL import ImageGrab
        return self._image_to_base64_jpeg(ImageGrab.grab())
    def run_script(self, code: str) -> str:
        for f in ["C:/Temp/ai_input.jsx","C:/Temp/ai_output.txt","C:/Temp/ai_error.txt"]:
            if os.path.exists(f): os.unlink(f)
        with open("C:/Temp/ai_input.jsx", "w", encoding="utf-8") as f:
            f.write(code)
        for _ in range(300):
            time.sleep(0.1)
            if os.path.exists("C:/Temp/ai_error.txt"):
                msg = open("C:/Temp/ai_error.txt", encoding="utf-8").read()
                os.unlink("C:/Temp/ai_error.txt")
                raise RuntimeError(msg)
            if os.path.exists("C:/Temp/ai_output.txt"):
                result = open("C:/Temp/ai_output.txt", encoding="utf-8").read().strip()
                os.unlink("C:/Temp/ai_output.txt")
                return result or "Script executed (no return value)"
        raise RuntimeError("Timeout: worker did not respond in 30s. Is ai_worker.py running?")

class MacBackend(IllustratorBackend):
    _APP_NAME = "Adobe Illustrator"
    def focus_app(self):
        subprocess.run(["osascript", "-e", f'tell application "{self._APP_NAME}" to activate'])
    def capture_screenshot(self):
        self.focus_app(); time.sleep(1)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as f: tmp = f.name
        try:
            subprocess.run(["screencapture", "-x", "-t", "jpg", tmp], check=True, timeout=10)
            return self._image_to_base64_jpeg(Image.open(tmp))
        finally:
            if os.path.exists(tmp): os.unlink(tmp)
    def run_script(self, code: str) -> str:
        with tempfile.NamedTemporaryFile(suffix=".jsx", delete=False, mode="w", encoding="utf-8") as f:
            f.write(code); jsx_path = f.name
        try:
            r = subprocess.run(["osascript", "-e", f'tell application "{self._APP_NAME}" to do javascript (read POSIX file "{jsx_path}") as string'], capture_output=True, text=True, timeout=30)
            if r.returncode != 0: raise RuntimeError(r.stderr.strip())
            return r.stdout.strip() or "Script executed (no return value)"
        finally: os.unlink(jsx_path)

def get_backend():
    if sys.platform == "darwin": return MacBackend()
    elif sys.platform == "win32": return WindowsBackend()
    else: raise RuntimeError(f"Unsupported: {sys.platform}")
'@
$lines = $content -split "`n"
[System.IO.File]::WriteAllLines(
    "C:\Users\$env:USERNAME\illustrator-mcp\illustrator\platform_backend.py",
    $lines,
    [System.Text.Encoding]::UTF8
)
Write-Host "platform_backend.py patched successfully!"
```

---

## Step 4 — Create the Worker Script

```powershell
# Create C:\Temp folder
New-Item -ItemType Directory -Force -Path "C:\Temp"

# Write the worker script
$lines = @(
    'import win32com.client, os, time',
    'print("Worker ready - connecting to Illustrator...")',
    'ai = win32com.client.Dispatch("Illustrator.Application")',
    'print("Connected to: " + ai.Name)',
    'print("Watching C:\\Temp for scripts...")',
    'while True:',
    '    if os.path.exists("C:/Temp/ai_input.jsx"):',
    '        try:',
    '            code = open("C:/Temp/ai_input.jsx", encoding="utf-8").read()',
    '            os.unlink("C:/Temp/ai_input.jsx")',
    '            result = ai.DoJavaScript(code)',
    '            open("C:/Temp/ai_output.txt", "w", encoding="utf-8").write(str(result) if result else "")',
    '        except Exception as e:',
    '            open("C:/Temp/ai_error.txt", "w", encoding="utf-8").write(str(e))',
    '    time.sleep(0.1)'
)
[System.IO.File]::WriteAllLines("C:\Temp\ai_worker.py", $lines, [System.Text.Encoding]::UTF8)
Write-Host "Worker created at C:\Temp\ai_worker.py"
```

---

## Step 5 — Create the mcp-proxy Virtual Environment

`supergateway` (npm) wraps the stdio MCP server and exposes it as HTTP on port 8787. It requires Node.js only — no separate Python venv needed.

Verify supergateway works:
```powershell
npx -y supergateway --help
```

> supergateway is downloaded automatically via `npx -y` on first run. No separate install needed.

---

## Step 6 — Open Windows Firewall Port 8787

Run in **PowerShell (Admin)**:

```powershell
# Remove any old conflicting rules
Remove-NetFirewallRule -DisplayName "MCP Illustrator*" -ErrorAction SilentlyContinue

# Allow Edge container (192.168.127.x) to reach port 8787
New-NetFirewallRule `
    -DisplayName "MCP Illustrator Edge VM" `
    -Direction Inbound `
    -Protocol TCP `
    -LocalPort 8787 `
    -Action Allow `
    -RemoteAddress 192.168.127.0/24 `
    -Profile Any

Write-Host "Firewall rule created."
```

> The Edge AI container connects via `host.edge.internal` which resolves to `192.168.127.1` (the Lima/QEMU virtual network gateway).

---

## Step 7 — Set PowerShell Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
# Type A (Yes to All) when prompted
```

---

## Step 8 — Configure Edge MCP Settings

1. Open **Edge → Settings → MCP Servers**
2. Click **+ Add MCP Server**
3. Fill in:

| Field | Value |
|---|---|
| **Name** | `Adobe Illustrator` |
| **Type** | `HTTP (remote URL)` |
| **URL** | `http://host.edge.internal:8787/mcp` |
| **Environment Variables** | *(leave empty)* |

4. Click **Save**

---

## Every Session — Startup Procedure

You need **two PowerShell windows** open whenever you want to use Illustrator with Edge.

### Step A — Start Adobe Illustrator

Open Illustrator and load the `.ai` document you want to work with.

### Step B — Worker Window (Regular PowerShell — NOT Admin)

```powershell
C:\Users\OlivierCatry\illustrator-mcp\.venv\Scripts\python.exe C:\Temp\ai_worker.py
```

Wait for:
```
Worker ready - connecting to Illustrator...
Connected to: Adobe Illustrator
Watching C:\Temp for scripts...
```

> ⚠️ **Critical:** Run from a **regular (non-Admin)** PowerShell. COM automation requires matching elevation with Illustrator's process.

### Step C — Proxy Window (Admin PowerShell)

```powershell
$env:TEMP = "C:\Temp"
$env:TMP = "C:\Temp"
npx -y supergateway `
    --stdio "C:/Users/OlivierCatry/illustrator-mcp/.venv/Scripts/python.exe -m illustrator" `
    --port 8787 `
    --outputTransport streamableHttp
```

Wait for:
```
[supergateway] Listening on port 8787
[supergateway] - outputTransport: streamableHttp
```

### Step D — Activate in Edge

1. **Settings → MCP** → click **Refresh Tools** on the Adobe Illustrator entry
2. Confirm 🟢 **green** with **7 tools**

---

## Verification Test

In Edge chat, type:
> *"What is the name of the active Illustrator document?"*

Expected: the name of your `.ai` file returned immediately.

---

## What You Can Do

Once connected, ask Edge naturally:

| Request | What happens |
|---|---|
| *"Take a screenshot of Illustrator"* | Returns live view of the canvas |
| *"Add text 'Hello World' in blue at the center"* | Runs ExtendScript to add styled text |
| *"Draw a red rectangle 200x100 at the top-left"* | Creates a vector shape |
| *"List all layers in the document"* | Inspects document structure |
| *"Export as PDF to my Desktop"* | Saves PDF to `C:\Users\<you>\Desktop` |
| *"Create a new document A4 size"* | Opens a new Illustrator document |
| *"Add a title 'Service Management' in IBM Blue"* | Adds formatted text |

---

## Troubleshooting

### 🔴 MCP shows red / connection refused

**Cause:** Supergateway not running or firewall blocking port 8787.  
**Fix:**
1. Check the Admin PowerShell — is supergateway running?
2. Run: `netstat -an | findstr :8787` — should show `0.0.0.0:8787 LISTENING`
3. Re-run the firewall rule from Step 6

---

### 🟡 Green but 0 tools or tools missing

**Cause:** Edge lost the tool registration (happens when server restarts).  
**Fix:** Go to **Settings → MCP** → toggle server **OFF** then back **ON**.

---

### ❌ "Timeout: worker did not respond in 30s"

**Cause:** The worker (`ai_worker.py`) is not running.  
**Fix:** Open a regular (non-Admin) PowerShell and start the worker.

---

### ❌ Worker crashes with COM error on startup

**Cause:** Worker is running as Admin but Illustrator is running as regular user.  
**Fix:** Stop the worker. Start a **regular (non-Admin)** PowerShell and run the worker from there.

---

### ❌ "spawn osascript ENOENT"

**Cause:** You are using an older Adobe MCP server designed for macOS only.  
**Fix:** This guide replaces that server. Ensure you are using `krVatsal/illustrator-mcp` with the patched `platform_backend.py` from Step 3.

---

### ❌ "mcp<1.0 conflicts with mcp-proxy"

**Cause:** Trying to install both in the same venv.  
**Fix:** This setup uses `supergateway` (npm) instead of `mcp-proxy` (Python) — no Python version conflicts.

---

### ❌ File permission denied when patching platform_backend.py

**Cause:** File is owned by system or another user (common when cloned to `system32`).  
**Fix:** Run `takeown /f <filepath>` then `icacls <filepath> /grant "<username>:(F)"` in Admin PowerShell before editing.

---

## Architecture Notes

### Why two separate processes (server + worker)?

Windows COM automation of interactive UI applications requires the calling process to share the same Windows session and desktop as the target app. When Node.js (supergateway) spawns Python as a subprocess, that subprocess is in a non-interactive context and cannot make COM calls to Illustrator.

The worker process runs **directly in a user terminal** (interactive session), giving it full COM access. The MCP server communicates with the worker via temp files in `C:\Temp` — a simple, reliable IPC mechanism.

### Why supergateway instead of mcp-proxy?

`mcp-proxy` exposes the server at `/sse` (SSE transport). Edge's egress proxy requires `/mcp` (StreamableHTTP transport). `supergateway` with `--outputTransport streamableHttp` serves at `/mcp`.

### Why `host.edge.internal` instead of `localhost`?

Edge runs inside a Linux container. `localhost` inside the container refers to the container, not your Windows machine. `host.edge.internal` resolves to `192.168.127.1` — the Windows host as seen from inside the container.

### Version Matrix

| Component | Version | Why |
|---|---|---|
| `mcp` (illustrator venv) | `< 1.0` (0.9.x) | illustrator-mcp uses old `@server.list_tools()` API removed in 1.0 |
| `supergateway` | latest | Serves `/mcp` StreamableHTTP endpoint |
| `pywin32` | latest | Windows COM automation |
| `Pillow` | latest | Screenshot capture via ImageGrab |

---

## File Reference

| File | Location | Purpose |
|---|---|---|
| MCP Server | `C:\Users\<you>\illustrator-mcp\` | Python MCP server (krVatsal) |
| Patched backend | `...\illustrator-mcp\illustrator\platform_backend.py` | Windows IPC fix (replaces direct COM) |
| Worker script | `C:\Temp\ai_worker.py` | Interactive COM bridge |
| IPC input | `C:\Temp\ai_input.jsx` | Script written by MCP server, read by worker |
| IPC output | `C:\Temp\ai_output.txt` | Result written by worker, read by MCP server |
| IPC error | `C:\Temp\ai_error.txt` | Error written by worker if script fails |

---

*Document generated by Edge AI Agent — IBM Consulting*  
*Version 1.0 — October 2026*
