# Live Server Setup Instructions

## How to Use Live Server with Your HTML File:

### Method 1: Right-click on index.html
1. Open VS Code
2. Navigate to: `src/ui/index.html`
3. Right-click on the `index.html` file
4. Select "Open with Live Server"

### Method 2: Use Status Bar
1. Open `src/ui/index.html` in VS Code
2. Look for "Go Live" button in the bottom status bar
3. Click "Go Live"

### Method 3: Command Palette
1. Press `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac)
2. Type "Live Server: Open with Live Server"
3. Press Enter

## What Happens:
- Live Server will start on port 5500 (configured in .vscode/settings.json)
- Your browser will automatically open to: http://localhost:5500
- The page will auto-refresh when you make changes to your files

## Configuration Applied:
- Root directory: `/src/ui` (so Live Server serves from the correct location)
- Port: 5500
- Host: localhost
- Browser: Default browser (will use Brave if it's your default browser)

## If it doesn't work:
1. Make sure Live Server extension is enabled
2. Check that you're opening the HTML file from the correct location (src/ui/index.html)
3. Ensure no other service is using port 5500

## Current File Structure:
```
FinDeck(Excel to PPT Project)/
├── .vscode/
│   └── settings.json (Live Server configuration)
└── src/
    └── ui/
        ├── index.html (main file)
        ├── assets/
        │   ├── css/
        │   ├── js/
        │   └── vendor/
        └── ...
```