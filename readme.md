# 🎬 Media Converter Pro v1.0

Media Converter Pro is a high-performance, all-in-one media management suite 🚀. Built with **PowerShell** and **WPF**, it provides a sleek graphical interface for powerful command-line engines. Handle batch processing, web downloads, and AI-driven enhancements without ever touching a terminal.

---

## ✨ Core Features

### 🎧 Audio Mastery
- **Batch Processing:** Convert between MP3, M4A, WAV, and FLAC seamlessly.
- **Precision Trimming:** Cut audio files with millisecond accuracy.
- **Smart Normalization:** Integrated EBU R128 loudness correction.
- **Extract & Save:** Pull high-quality audio from any video container.

### 📹 Video Engineering
- **Next-Gen Codecs:** Support for H.264, H.265 (HEVC), and AV1.
- **⚡ GPU Acceleration:** Native support for NVIDIA (NVENC), AMD (AMF), and Intel (QSV).
- **Target Size Logic:** Auto-calculate bitrates for Discord (25MB) or WhatsApp targets.
- **Visual Preview:** Generate frame-based timelines to inspect files before processing.

### 🖼️ Image Optimization
- **Modern Formats:** Convert to JPG, PNG, WEBP, HEIC, or BMP.
- **Privacy Shield:** One-click removal of all EXIF and metadata.
- **Smart Scaling:** Batch resize images while preserving aspect ratio.

### 🌐 Advanced Downloader
- **Powered by [yt-dlp](https://github.com/yt-dlp/yt-dlp):** Download from YouTube, SoundCloud, TikTok, and 1000+ sites. [List of Supported Sites](https://github.com/yt-dlp/yt-dlp/blob/master/supportedsites.md)
- **4K/8K Ready:** Fetch the highest available quality automatically.
- **Ad-Free:** Integrated **SponsorBlock** to strip ads and sponsors from videos.
- **Auth Support:** Easy PO Token and Cookie integration for restricted content.

### 🤖 AI & Special Tools
- **AI Transcriber:** Local voice-to-text using [OpenAI Whisper](https://github.com/openai/whisper).
- **AI Upscaling:** Neural network image enhancement via [Upscayl](https://www.upscayl.org/).
- **Video Stabilizer:** 2-pass "vidstab" algorithm to fix shaky handheld footage.
- **Instant Muxing:** Merge video and audio streams without re-encoding.

---

## Screenshots

![1](/pictures/pic1.jpg)

<br>

![2](/pictures/pic2.jpg)

<br>

![3](/pictures/pic3.jpg)

<br>

![4](/pictures/pic4.jpg)

<br>

![5](/pictures/pic5.jpg)

<br>


---

## 🛠️ Requirements

Media Converter Pro features a **Smart Auto-Installer**. If tools are missing, the app will offer to set them up via `winget`.

* **[FFmpeg](https://ffmpeg.org/):** The core media engine.
* **[yt-dlp](https://github.com/yt-dlp/yt-dlp):** For web extraction.
* **[Node.js](https://nodejs.org/):** For YouTube signature decryption.
* **[Python](https://www.python.org/):** Required for Whisper AI features.

---

## 🚀 Installation & Usage

You can run Media Converter Pro in two ways:

### Option 1: Running the Script (.ps1) - Recommended
1. **Download** the `MediaConverterPro.ps1` file.
2. **Right-click** the file and select **Run with PowerShell**.
3. **Execution Policy Note:** If Windows prevents the script from running, you can bypass this by opening PowerShell and typing:
   `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
   *Alternatively:* Open the script in a text editor (like Notepad), copy everything, paste it into a new PowerShell window, and hit Enter.

### Option 2: Running the Executable (.exe)
If you have a compiled version:
1. Simply double-click `MediaConverterPro.exe`. You can retrieve the newest .exe from the [releases](https://github.com/Frxhb/MediaConverterPro/releases/latest) page.
2. **⚠️ Antivirus Note:** Some antivirus engines (like Windows Defender) may falsely flag compiled PowerShell scripts as a "Positive" threat. This is a common issue with `PS2EXE` compilers. If this happens, you can either add an exclusion or stick to the `.ps1` script version.

---

## 🛡️ Technical Highlights

- **Queue Persistence:** If the app closes, your job queue is saved and can be resumed.
- **Live Logs:** Real-time STDOUT/STDERR monitoring for every task.
- **Isolation:** Every process runs in the background; the UI stays fluid and responsive.
- **Custom Flags:** Inject your own FFmpeg or yt-dlp arguments for ultimate control.

---

## ⚖️ License
Distributed under the **MIT License**. See `LICENSE` for more information.

> **Note:** Please respect copyright laws and the terms of service of any platforms you interact with.