# Weekday Wallpaper Changer for Windows 11

**Change your desktop wallpaper automatically, based on the weekday — no manual fiddling required.**

## 🚀 Features

1. **Graphical Folder Selection**  
   - Easy GUI dialogs for picking image and script locations—no path-typing needed.
2. **Flexible Script Installation**  
   - Installer copies and patches the script for you; destination folder is user-chosen.
3. **Automatic Path Patching**  
   - The installer automatically updates image paths in your script.
4. **Task Scheduler Automation**  
   - Sets up a daily scheduled task at midnight to change wallpaper.
5. **Error Handling and Cleanup**  
   - Clear error messages, uninstall option included.

## 🖼️ Image File Setup

Place 7 images (e.g. `1-Sun.jpg`, `2-Mon.png`, `3-Tue.bmp`, ..., `7-Sat.webp`) in a folder of your choice.

**Supported image formats:** JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP

The script will automatically detect the file extension for each weekday image.

## 🛠️ PowerShell Script Setup

Your `set_wallpaper.ps1` should use this line:
```powershell
$imgPath = "{IMG_FOLDER}\$filename"
