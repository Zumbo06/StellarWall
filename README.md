StellarWall is a lightweight and customizable live wallpaper engine for Windows that allows you to set animated GIFs and MP4 videos as your desktop background. Personalize your workspace with dynamic playlists that change based on time intervals, the time of day, or even the day of the week.

## ✨ Features

*   **Animated Backgrounds:** Supports both **GIF** and **MP4** video files.
*   **Dynamic Playlists:**
    *   **Interval Mode:** Cycle through a folder of wallpapers or a custom list at your chosen time interval (minutes or hours). Supports manual order, sequential playback (with initial shuffle), or fully random selection.
    *   **Time of Day Mode:** Automatically switch wallpapers to match the mood of the morning, afternoon, evening, and night.
    *   **Day of Week Mode:** Curate unique wallpaper playlists for each day of the week.
*   **Background Audio:** Enhance your live wallpaper with an accompanying MP3 audio track.
*   **Performance Optimization:**
    *   **Low Spec PC Mode:** Option to limit MP4 video playback to 1080p or lower to conserve resources.
    *   **Focus-Aware Pausing:** Automatically pauses visuals (and optionally audio) when the desktop is not active, saving system resources.
    *   **Aggressive GPU Reduction:** Optionally hides wallpaper content entirely when the desktop loses focus for maximum GPU savings (may cause a slight flicker on focus change).
    *   **Configurable Preview Quality:** Choose between smooth (higher quality) or fast (lower resource) video preview generation in the UI.
*   **User-Friendly Interface:** Easily select files, manage playlists, and configure settings through an intuitive tabbed interface.
*   **System Tray Integration:** Runs conveniently in the system tray with quick access to:
    *   Show/Hide the main application window.
    *   Pause/Resume the entire engine.
    *   Quickly play recent wallpapers.
    *   Quit the application.
*   **Start with Windows:** Optionally configure StellarWall to launch automatically when you log in to Windows.
*   **Recent Wallpapers:** Quickly access your last few used wallpapers directly from the tray menu.

## 🚀 Getting Started

### Prerequisites

*   Windows Operating System (primarily tested on Windows 10/11)
*   [Python](https://www.python.org/downloads/) (if running from source, version 3.9+ recommended)
*   Necessary video codecs for MP4 playback (usually included with Windows, but Media Feature Pack might be needed for N/KN editions).

### Installation

**Option 1: Download Pre-built Release (Recommended for most users)**

1.  Go to the [Releases page](https://github.com/Zumbo06/StellarWall/releases) of this repository.
2.  Download the latest `StellarWall_vX.Y.Z.zip` (or `.exe` if a one-file build is provided).
3.  Extract the ZIP file to a folder of your choice (if applicable).
4.  Run `StellarWall.exe`.

**Option 2: Running from Source**

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/Zumbo06/StellarWall.git
    cd StellarWall
    ```
2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv .venv
    # Windows
    .venv\Scripts\activate
    # macOS/Linux
    source .venv/bin/activate
    ```
3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
4.  **Run the application:**
    ```bash
    python live_wallpaper_qt6.py
    ```

## 📖 How to Use

1.  **Launch StellarWall.**
2.  **Select a Wallpaper Mode:**
    *   **Single Wallpaper:** Choose a single GIF or MP4 file. Enable sound for MP4s if desired.
    *   **Playlist - Interval:** Add files or a folder to create a playlist. Set the change interval and playback order.
    *   **Playlist - Time of Day:** Assign different wallpapers for Morning, Afternoon, Evening, and Night.
    *   **Playlist - Day of Week:** Create separate playlists for each day.
3.  **Background Audio (Optional):** Select an MP3 file to play in the background. Control volume with the slider.
4.  **Click "Apply Wallpaper"** to start the live wallpaper.
5.  **Controls:**
    *   **Pause Visual / Resume Visual:** Temporarily pause or resume the visual animation.
    *   **Stop & Clear Visual:** Stops the current wallpaper and clears it from the desktop.
6.  **Application Settings Tab:**
    *   Configure "Start with Windows."
    *   Adjust performance optimization settings.
7.  **System Tray Icon:** Right-click the StellarWall icon in your system tray for quick actions
