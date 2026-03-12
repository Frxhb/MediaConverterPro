# ❓ Frequently Asked Questions

### 1. Why does it ask to install "WinGet" tools?
Media Converter Pro is a graphical wrapper. It needs the "engines" (FFmpeg, yt-dlp, etc.) to do the heavy lifting. WinGet is the safest way to keep these tools updated on Windows.

### 2. Is my data sent to any server?
**No.** All processing—including AI Transcription and AI Upscaling—happens **100% locally** on your computer. Your files never leave your machine.

### 3. Why is my download failing with a "403 Forbidden" error?
YouTube frequently updates its code to block downloaders. 
- Ensure you have **Node.js** installed (check the Update Tools menu).
- Try using the **PO Token** or **Cookie** features in the Download tab.

### 4. Windows says I cannot run scripts. What do I do?
Windows has a security feature called "Execution Policy." You can fix this by running this command in PowerShell as Admin:
`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force`

### 5. Why does my Antivirus block the .exe file?
Because the `.exe` is a compressed PowerShell script, some scanners think it is suspicious. This is a "False Positive." The source code is open, so you can always run the `.ps1` directly to be 100% safe.

### 6. How do I use my GPU for faster conversion?
In the Video tab, use the **Hardware Acceleration** dropdown and select your card (NVENC for Nvidia, AMF for AMD, QSV for Intel).