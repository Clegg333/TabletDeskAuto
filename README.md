# TabletDesk

> **Notice:** This program was created with the assistance of ChatGPT-4 and Windscribe. All credit goes to these tools. The program is not malicious, but it is untested on machines other than my own. **Use at your own risk.**

TabletDesk is a Windows utility designed to seamlessly launch Deskthing web apps in Microsoft Edge Kiosk Mode on an Android tablet used as an additional display via SpaceDesk. The app is ideal for turning an old Android tablet into a dedicated dashboard, such as a weather station or smart home control panel, which automatically appears on boot.

## Purpose
- **Automatic Deskthing App Launch:** On system startup, TabletDesk waits for a SpaceDesk-connected Android tablet to be available as a secondary display, then launches your chosen Deskthing web app (such as WeatherWaves) in Edge's full-screen kiosk mode on that display.
- **Hands-Free Experience:** No manual setup needed after each reboot—just power on your PC and tablet, and your dashboard appears automatically.

## Features
- **Edge Kiosk Mode:** Launches Microsoft Edge in kiosk mode for a true full-screen, touch-friendly experience.
- **Display Detection:** Waits for the SpaceDesk virtual display (Android tablet) to connect before launching the app.
- **Configurable App & Display:** Easily choose which Deskthing app and which display to use.
- **Hotkey Support:** Instantly relaunch the kiosk with a customizable hotkey (default: Ctrl+Alt+W).
- **System Tray Icon:** Control and configure the app from the Windows system tray.
- **Robust Logging:** All actions and errors are logged for easy troubleshooting.
- **Installer with EULA:** Professional, user-friendly installer with EULA acceptance and shortcut creation.
- **No Personal Data:** All user paths and settings are generic and privacy-safe.

## Typical Use Case
- Repurpose an Android tablet as a dedicated weather dashboard, smart home control, or any Deskthing web app.
- On PC boot, the dashboard appears automatically on the tablet (connected via SpaceDesk as a second monitor).

## Installation Guide

### Prerequisites
- **Windows 10/11** PC
- **Microsoft Edge** browser installed
- **SpaceDesk** server installed on your PC
- **SpaceDesk** app installed on your Android tablet
- **DeskThing** web app (your target dashboard)
- **AutoHotkey v2** installed (for hotkey and automation support)

### Steps
1. **Connect your Android tablet as a display:**
   - Launch SpaceDesk server on your PC.
   - Open SpaceDesk app on your Android tablet and connect to your PC.
   - Your tablet should now appear as an extra monitor in Windows Display Settings.
2. **Download and Run the Installer:**
   - Go to the [GitHub Releases page](https://github.com/Clegg333/TabletDeskAuto/releases) and download the latest `TabletDesk-Installer.exe`.
   - Double-click the installer to start. (You may see a Windows SmartScreen warning about an unknown publisher—click 'More info' > 'Run anyway'.)
3. **Accept the EULA:**
   - Read and accept the End User License Agreement to continue.
4. **Choose Install Location:**
   - Select a folder or use the default location for installation.
5. **Complete Installation:**
   - The installer will extract all required files and create shortcuts.
   - Optionally, launch TabletDesk immediately after install.
6. **First Use:**
   - Use the Desktop shortcut to start TabletDesk.
   - On boot, TabletDesk will wait for your SpaceDesk-connected tablet and launch the Deskthing app in Edge Kiosk Mode on that display.
7. **Configuring TabletDesk:**
   - Use the system tray icon to access settings, change the Deskthing app URL, hotkey, or display.

## Troubleshooting
- **Edge or SpaceDesk not detected:** Ensure both are installed and your tablet is connected as an extra display.
- **SmartScreen warning:** This is normal for self-signed certificates. See Security Note below.
- **Nothing appears on tablet:** Make sure the tablet is connected and set as an extended display, not duplicated.

## Security Note
This installer is signed with a self-signed certificate for integrity. If you wish to trust the certificate, import `TabletDeskInstaller.cer` into your Trusted Root Certification Authorities. See the release notes for details.

## License
See `TabletDesk-EULA.txt` for license terms.
