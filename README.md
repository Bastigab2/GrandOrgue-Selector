# GrandOrgue Selector

Visual manager for GrandOrgue virtual organs.

## Description

GrandOrgue Selector is a desktop application that allows you to organize, preview and launch your GrandOrgue virtual organ collection. It provides a modern interface with image previews, categories, statistics and automatic launch features.

## Features

- **Organ Management**: Load, import, edit and delete organs
- **Visual Preview**: Display screenshots/images for each organ
- **Categories**: Organize organs by type (Baroque, Romantic, Modern, Contemporary, Other)
- **Search & Filter**: Quick search by name or description
- **Priority System**: Reorder organs by priority
- **Favorites**: Set a favorite organ for auto-launch
- **Auto-Launch**: Automatically launch favorite organ on startup with countdown
- **Statistics**: Track launch count and last used date
- **Screenshot Capture**: Capture GrandOrgue window screenshots directly (JPG format for smaller files)
- **Multi-Monitor Support**: All dialogs and screenshots work on the same monitor as the application
- **Export/Import**: Backup and restore configuration with images (ZIP format)
- **Multi-language**: English, Italian, French, German, Spanish
- **Themes**: 3 visual themes available
  - Classic Dark
  - Light Wood
  - Walnut

## Requirements

- Python 3.x
- Pillow (PIL)

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   pip install Pillow
   ```
3. Run the application:
   ```bash
   python GrandOrgueSelector.py
   ```

## Usage

### Adding Organs

- **Load Organ**: Click "Load Organ" button or use `Ctrl+O` to add a single organ file (.orgue/.organ)
- **Import Folder**: Click "Import Folder" or use `Ctrl+I` to scan a directory and import all organ files

### Managing Organs

- **Edit**: Select an organ and click "Edit" to modify name, category and description
- **Delete**: Select an organ and click "Delete" or press `Delete` key
- **Priority**: Use "Priority +" and "Priority -" buttons to reorder organs
- **Favorite**: Right-click on an organ to set/remove as favorite (marked with *)

### Launching Organs

- Double-click on an organ in the list
- Press `Enter` with an organ selected
- Click the "LAUNCH ORGAN" button

### Adding Images

Click on the preview area to:
- Select an existing image file
- Capture a screenshot of GrandOrgue window (5 second countdown)

### Auto-Launch

1. Set a favorite organ (right-click -> "Set as favorite")
2. Go to View -> Settings
3. Enable "Auto-launch favorite organ on startup"
4. Set delay (minimum 3 seconds)
5. On next startup, press any key to cancel auto-launch

### Changing Theme

Go to View -> Theme and select:
- **Classic Dark**: Default dark theme
- **Light Wood**: Warm light wood tones
- **Walnut**: Dark walnut wood with piano-style buttons

### Changing Language

Go to View -> Language and select your preferred language.

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| `Ctrl+O` | Load organ |
| `Ctrl+I` | Import folder |
| `Enter` | Launch selected organ |
| `Delete` | Delete selected organ |
| `F1` | Show help |
| `F5` | Refresh |
| `Esc` | Exit / Cancel auto-launch |

## Files

- `organs.json` - Organ configuration database
- `settings.json` - Application settings
- `images/` - Stored organ images/screenshots

## Export/Import

- **Export**: Creates a ZIP file containing configuration, settings and all images
- **Import**: Restores from ZIP or merges from JSON file

## Author

**Gabriele Bastianelli**
Urbino, Italy

## License

This project is provided as-is for personal use.

## Version

1.2.0
