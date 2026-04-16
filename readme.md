# Media Converter Pro

Media Converter Pro is a lightweight, high-performance media management suite built entirely in PowerShell and WPF. It acts as a unified graphical frontend for industry-standard CLI tools (FFmpeg, yt-dlp, HandBrakeCLI, and Whisper AI). 

Designed for power users, it handles batch processing, web extraction, and AI-driven enhancements locally, without the overhead of heavy Electron apps and without requiring you to touch a terminal.

---

## ⚡ Core Features

### Video Engineering
* **Hardware Acceleration:** Native support for NVIDIA (NVENC), AMD (AMF), and Intel (QSV) encoding.
* **Advanced Codecs:** Support for H.264, H.265 (HEVC), and AV1.
* **HandBrakeCLI Integration:** Dedicated toggle to use the Handbrake engine—ideal for fixing Variable Framerate (VFR) audio sync issues.
* **Target Size Logic:** Automatically calculates bitrates to hit strict file size limits (e.g., 24.5MB for Discord).
* **Visual Trimming:** Generates a frame-based timeline preview for precise cutting.

### Audio & Muxing
* **Batch Conversion:** MP3, M4A, WAV, and FLAC support.
* **Loudness Normalization:** Integrated EBU R128 loudness correction.
* **Instant Muxing:** A dedicated tab to merge video and audio streams instantly without re-encoding.
* **Audio Extraction:** Pull high-quality audio directly from video containers with one click.

### Advanced Downloader (yt-dlp)
* **Universal Extraction:** Download from YouTube, SoundCloud, Twitter/X, Reddit, and 1000+ other sites.
* **Anti-Bot Bypass:** Native support for browser cookie extraction and manual/auto PO Token injection.
* **SponsorBlock:** Automatically strip sponsored segments from downloaded videos.
* **Post-Processing:** Auto-embed metadata, subtitles, and thumbnails directly into the container.

### Local AI & Special Tools
* **AI Transcriber:** Local voice-to-text generation using [OpenAI Whisper](https://github.com/openai/whisper). Can export as `.srt`, `.vtt`, or hardcode/burn subtitles directly into a new video.
* **AI Upscaling:** Neural network image enhancement integration via [Upscayl](https://www.upscayl.org/).
* **Video Stabilizer:** Utilizes the 2-pass `vidstab` algorithm to smooth shaky handheld footage.
* **Audio Visualizer:** Generates waveforms, frequency bars, or vectorscopes from audio files to create ready-to-publish videos.

### Image Optimization
* **Format Conversion:** JPG, PNG, WEBP, ICO, HEIC, and BMP.
* **Privacy Shield:** Complete removal of EXIF and metadata.
* **Smart Scaling:** Batch resize images while preserving aspect ratios.

---

## 🛠️ Prerequisites & Auto-Installer

Media Converter Pro features a **Smart Auto-Installer**. Upon first launch, the script checks your system `PATH`. If tools are missing, the GUI will prompt you to automatically install them via Windows Package Manager (`winget`).

The suite relies on:
* **FFmpeg / FFprobe** (Core media engine)
* **yt-dlp** (Web extraction)
* **HandBrakeCLI** (Alternative video encoder)
* **Node.js** (Required by yt-dlp for signature decryption)
* **Python** (Required for Whisper AI)
* **Upscayl** (Required for AI image upscaling)

---

## 🚀 Installation & Usage

### Option 1: Running the Script (Recommended)
Running the raw `.ps1` file ensures you are running the exact open-source code without false-positive antivirus flags.

1. Download the latest `MediaConverterPro.ps1`.
2. Right-click the file -> **Run with PowerShell**.

**Execution Policy Note:**
If Windows blocks the script, open a PowerShell terminal as Administrator and run:

    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

Alternatively, you can right-click the `.ps1` file, select **Properties**, check **Unblock**, and click Apply.

### Option 2: Compiled Executable (.exe)
If you prefer a portable executable, download the latest `MediaConverterPro.exe` from the [Releases](https://github.com/Frxhb/MediaConverterPro/releases/latest) page. 

> **⚠️ Antivirus Notice:** The `.exe` is compiled using `PS2EXE`. Because it wraps a PowerShell script inside an executable, Windows Defender or other AV engines may occasionally flag it as a generic threat. This is a false positive. If this occurs, add an exclusion or use Option 1.

---

## ⚙️ Under the Hood

* **Queue Persistence:** If the app is closed or crashes, the current batch queue is serialized to `mcp_queue.json` and can be resumed on the next launch.
* **Live Logs:** Real-time `STDOUT`/`STDERR` monitoring. The UI reads the active buffer asynchronously, keeping the app responsive even during heavy FFmpeg tasks.
* **Custom Parameter Injection:** Every tab includes a "Custom Params" override, allowing you to inject raw FFmpeg or yt-dlp flags while previewing the final compiled command in real-time.
* **System Tray:** The app can be minimized directly to the system tray for background processing.

---

## 📸 Screenshots

![MainGui](/pictures/pic1.jpg)
<br>
![DownloadScreen](/pictures/pic2.jpg)
<br>
![SpecialTab](/pictures/pic3.jpg)
<br>
![UpdateDependencies](/pictures/pic4.jpg)
<br>
![Settings](/pictures/pic5.jpg)

---

## ⚖️ License

Distributed under the **MIT License**. See `LICENSE` for more information.

> **Disclaimer:** This tool is intended for personal use, archiving, and editing your own media. Please respect copyright laws and the terms of service of the platforms you interact with.