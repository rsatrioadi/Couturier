# Couturier

A small macOS utility for creating and editing Microsoft Office font themes. Vibe-coded by Claude.

## What it does

Lists, previews, and edits the XML font theme files that Office for Mac reads from your user content folder (`Theme Fonts`). Each theme sets a **heading font** and a **body font**, which appear in the theme font picker inside Word, Excel, and PowerPoint.

## Requirements

- macOS 11 (Big Sur) or later
- Microsoft Office for Mac installed (provides the Themes folder)
- [Rust](https://rustup.rs) + [Node.js](https://nodejs.org) v18+

## Dev setup

```sh
npm install
npm run dev     # builds and launches the app
npm run build   # produces a .app bundle in src-tauri/target/release/bundle
```

## Themes folder

Couturier auto-discovers the Themes folder on first launch using heuristics
(`User Content`, `Themes`, etc.). If it can't find it,
a banner appears letting you point to the correct folder manually. The chosen
path is saved to `~/Library/Application Support/com.couturier.app/config.json`.

The font theme files themselves live at a path like:

```
~/Library/Group Containers/UBF8T346G9.Office/User Content/Themes/Theme Fonts/
```

## File format

Each theme is a plain XML file:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:fontScheme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="My Theme">
  <a:majorFont><a:latin typeface="Helvetica Neue"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
```

## Planned

- Theme Colors editor (`Theme Colors/`)
- Theme Effects editor (`Theme Effects/`)
