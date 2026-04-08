use serde::{Deserialize, Serialize};
use std::collections::BTreeSet;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use tauri::{AppHandle, Manager, State};

// ---------------------------------------------------------------------------
// App state — holds the resolved Themes root for the session
// ---------------------------------------------------------------------------

pub struct ThemesRootState(pub Mutex<Option<PathBuf>>);

// ---------------------------------------------------------------------------
// Model
// ---------------------------------------------------------------------------

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct FontTheme {
    pub name: String,
    pub heading_font: String,
    pub body_font: String,
    pub file_path: String,
}

#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct ColorTheme {
    pub name: String,
    /// Hex strings without '#', e.g. "4472C4"
    pub dk1: String,
    pub lt1: String,
    pub dk2: String,
    pub lt2: String,
    pub accent1: String,
    pub accent2: String,
    pub accent3: String,
    pub accent4: String,
    pub accent5: String,
    pub accent6: String,
    pub hlink: String,
    pub fol_hlink: String,
    pub file_path: String,
}

// ---------------------------------------------------------------------------
// Config persistence
// ---------------------------------------------------------------------------

#[derive(Serialize, Deserialize, Default)]
struct Config {
    themes_root: Option<String>,
}

fn config_path(app: &AppHandle) -> Option<PathBuf> {
    app.path().app_data_dir().ok().map(|d| d.join("config.json"))
}

fn load_config(app: &AppHandle) -> Config {
    let Some(path) = config_path(app) else { return Config::default() };
    std::fs::read_to_string(&path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

fn save_config(app: &AppHandle, config: &Config) {
    let Some(path) = config_path(app) else { return };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(config) {
        let _ = std::fs::write(path, json);
    }
}

// ---------------------------------------------------------------------------
// Heuristic folder discovery
// ---------------------------------------------------------------------------

/// Try to find `name` inside `parent`, tolerating a `.localized` suffix or
/// any prefix match (case-insensitive). Returns the first existing directory
/// that matches, or `None`.
fn fuzzy_find(parent: &Path, name: &str) -> Option<PathBuf> {
    // 1. Exact
    let exact = parent.join(name);
    if exact.is_dir() {
        return Some(exact);
    }
    // 2. name.localized
    let localized = parent.join(format!("{}.localized", name));
    if localized.is_dir() {
        return Some(localized);
    }
    // 3. First entry whose name starts with `name` (case-insensitive)
    let name_lower = name.to_lowercase();
    let entries = std::fs::read_dir(parent).ok()?;
    let mut matches: Vec<PathBuf> = entries
        .flatten()
        .map(|e| e.path())
        .filter(|p| {
            p.is_dir()
                && p.file_name()
                    .and_then(|n| n.to_str())
                    .map(|n| n.to_lowercase().starts_with(&name_lower))
                    .unwrap_or(false)
        })
        .collect();
    matches.sort(); // deterministic
    matches.into_iter().next()
}

/// Discover the `Themes` root directory using heuristics.
/// Returns the path to the folder that directly contains
/// `Theme Fonts`, `Theme Colors`, `Theme Effects`.
fn discover_themes_root_path() -> Option<PathBuf> {
    let home = std::env::var("HOME").ok()?;
    let gc = PathBuf::from(&home).join("Library/Group Containers/UBF8T346G9.Office");
    if !gc.is_dir() {
        return None;
    }
    let user_content = fuzzy_find(&gc, "User Content")?;
    let themes = fuzzy_find(&user_content, "Themes")?;
    Some(themes)
}

/// Given the Themes root, resolve a specific subfolder (e.g. "Theme Fonts")
/// also with fuzzy matching.
fn resolve_subfolder(themes_root: &Path, subfolder: &str) -> Option<PathBuf> {
    fuzzy_find(themes_root, subfolder)
}

// ---------------------------------------------------------------------------
// Tauri commands — discovery & configuration
// ---------------------------------------------------------------------------

/// Called on startup. Returns the Themes root path (not the subfolder).
/// Load order: persisted config → heuristic. Saves to config if heuristic
/// succeeds and nothing was saved before.
#[tauri::command]
fn get_themes_root(
    app: AppHandle,
    state: State<'_, ThemesRootState>,
) -> Option<String> {
    let mut lock = state.0.lock().unwrap();

    // Already resolved this session
    if let Some(ref p) = *lock {
        return Some(p.to_string_lossy().into_owned());
    }

    // Try persisted config
    let config = load_config(&app);
    if let Some(ref saved) = config.themes_root {
        let p = PathBuf::from(saved);
        if p.is_dir() {
            *lock = Some(p.clone());
            return Some(p.to_string_lossy().into_owned());
        }
    }

    // Try heuristic
    if let Some(p) = discover_themes_root_path() {
        // Persist for next launch
        save_config(&app, &Config { themes_root: Some(p.to_string_lossy().into_owned()) });
        *lock = Some(p.clone());
        return Some(p.to_string_lossy().into_owned());
    }

    None
}

/// Validate and persist a user-supplied Themes root path.
#[tauri::command]
fn set_themes_root(
    app: AppHandle,
    state: State<'_, ThemesRootState>,
    path: String,
) -> Result<(), String> {
    let p = PathBuf::from(&path);
    if !p.is_dir() {
        return Err(format!("\"{}\" is not a valid directory", path));
    }
    save_config(&app, &Config { themes_root: Some(path.clone()) });
    *state.0.lock().unwrap() = Some(p);
    Ok(())
}

/// Open a native folder-picker dialog and return the selected path.
/// Runs blocking — that's fine on Tauri's command thread pool.
#[tauri::command]
fn pick_themes_folder() -> Option<String> {
    rfd::FileDialog::new()
        .set_title("Select the \"Themes\" folder")
        .pick_folder()
        .map(|p| p.to_string_lossy().into_owned())
}

/// Resolve a subfolder name (e.g. "Theme Fonts") under the given themes root,
/// using the same fuzzy matching. Returns the full path if found.
#[tauri::command]
fn resolve_theme_subfolder(themes_root: String, subfolder: String) -> Option<String> {
    let root = PathBuf::from(&themes_root);
    resolve_subfolder(&root, &subfolder).map(|p| p.to_string_lossy().into_owned())
}

// ---------------------------------------------------------------------------
// Minimal XML helpers
// ---------------------------------------------------------------------------

fn attr_value<'a>(tag: &'a str, attr: &str) -> Option<&'a str> {
    let needle = format!("{}=\"", attr);
    let start = tag.find(needle.as_str())? + needle.len();
    let end = start + tag[start..].find('"')?;
    Some(&tag[start..end])
}

fn parse_theme(xml: &str, file_path: &str) -> Option<FontTheme> {
    let scheme_start = xml.find("<a:fontScheme")?;
    let scheme_tag_end = xml[scheme_start..].find('>')?;
    let scheme_tag = &xml[scheme_start..scheme_start + scheme_tag_end + 1];
    let name = attr_value(scheme_tag, "name")?.to_string();

    let maj_s = xml.find("<a:majorFont>")?;
    let maj_e = xml.find("</a:majorFont>")?;
    let heading_font = find_latin_typeface(&xml[maj_s..maj_e]).unwrap_or_default();

    let min_s = xml.find("<a:minorFont>")?;
    let min_e = xml.find("</a:minorFont>")?;
    let body_font = find_latin_typeface(&xml[min_s..min_e]).unwrap_or_default();

    Some(FontTheme { name, heading_font, body_font, file_path: file_path.to_string() })
}

fn find_latin_typeface(section: &str) -> Option<String> {
    let tag_start = section.find("<a:latin")?;
    let tag_end = section[tag_start..].find('>')?;
    let tag = &section[tag_start..tag_start + tag_end + 1];
    Some(attr_value(tag, "typeface")?.to_string())
}

fn make_xml(name: &str, heading: &str, body: &str) -> String {
    let e = |s: &str| {
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
    };
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <a:fontScheme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"{name}\">\n\
         \x20 <a:majorFont>\n\
         \x20   <a:latin typeface=\"{heading}\"/>\n\
         \x20   <a:ea typeface=\"\"/>\n\
         \x20   <a:cs typeface=\"\"/>\n\
         \x20 </a:majorFont>\n\
         \x20 <a:minorFont>\n\
         \x20   <a:latin typeface=\"{body}\"/>\n\
         \x20   <a:ea typeface=\"\"/>\n\
         \x20   <a:cs typeface=\"\"/>\n\
         \x20 </a:minorFont>\n\
         </a:fontScheme>",
        name = e(name),
        heading = e(heading),
        body = e(body),
    )
}

// ---------------------------------------------------------------------------
// Color XML helpers
// ---------------------------------------------------------------------------

/// Extract the hex color value from a named slot element.
/// Handles both `<a:srgbClr val="RRGGBB"/>` and `<a:sysClr lastClr="RRGGBB" …/>`.
fn parse_color_slot(xml: &str, slot: &str) -> String {
    let open  = format!("<a:{}>", slot);
    let close = format!("</a:{}>", slot);
    let start = match xml.find(open.as_str()) {
        Some(s) => s + open.len(),
        None    => return "000000".to_string(),
    };
    let end = match xml[start..].find(close.as_str()) {
        Some(e) => start + e,
        None    => return "000000".to_string(),
    };
    let inner = &xml[start..end];
    if inner.contains("<a:srgbClr") {
        if let Some(v) = attr_value(inner, "val") {
            return v.to_uppercase();
        }
    }
    if inner.contains("<a:sysClr") {
        if let Some(v) = attr_value(inner, "lastClr") {
            return v.to_uppercase();
        }
    }
    "000000".to_string()
}

fn parse_color_theme(xml: &str, file_path: &str) -> Option<ColorTheme> {
    let scheme_start = xml.find("<a:clrScheme")?;
    let tag_end      = xml[scheme_start..].find('>')?;
    let scheme_tag   = &xml[scheme_start..scheme_start + tag_end + 1];
    let name = attr_value(scheme_tag, "name")?.to_string();

    Some(ColorTheme {
        name,
        dk1:      parse_color_slot(xml, "dk1"),
        lt1:      parse_color_slot(xml, "lt1"),
        dk2:      parse_color_slot(xml, "dk2"),
        lt2:      parse_color_slot(xml, "lt2"),
        accent1:  parse_color_slot(xml, "accent1"),
        accent2:  parse_color_slot(xml, "accent2"),
        accent3:  parse_color_slot(xml, "accent3"),
        accent4:  parse_color_slot(xml, "accent4"),
        accent5:  parse_color_slot(xml, "accent5"),
        accent6:  parse_color_slot(xml, "accent6"),
        hlink:    parse_color_slot(xml, "hlink"),
        fol_hlink: parse_color_slot(xml, "folHlink"),
        file_path: file_path.to_string(),
    })
}

fn make_color_xml(theme: &ColorTheme) -> String {
    let e = |s: &str| {
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
    };
    // Store all colors as srgbClr (uppercase hex, no '#').
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <a:clrScheme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"{name}\">\n\
         \x20 <a:dk1><a:srgbClr val=\"{dk1}\"/></a:dk1>\n\
         \x20 <a:lt1><a:srgbClr val=\"{lt1}\"/></a:lt1>\n\
         \x20 <a:dk2><a:srgbClr val=\"{dk2}\"/></a:dk2>\n\
         \x20 <a:lt2><a:srgbClr val=\"{lt2}\"/></a:lt2>\n\
         \x20 <a:accent1><a:srgbClr val=\"{a1}\"/></a:accent1>\n\
         \x20 <a:accent2><a:srgbClr val=\"{a2}\"/></a:accent2>\n\
         \x20 <a:accent3><a:srgbClr val=\"{a3}\"/></a:accent3>\n\
         \x20 <a:accent4><a:srgbClr val=\"{a4}\"/></a:accent4>\n\
         \x20 <a:accent5><a:srgbClr val=\"{a5}\"/></a:accent5>\n\
         \x20 <a:accent6><a:srgbClr val=\"{a6}\"/></a:accent6>\n\
         \x20 <a:hlink><a:srgbClr val=\"{hl}\"/></a:hlink>\n\
         \x20 <a:folHlink><a:srgbClr val=\"{fh}\"/></a:folHlink>\n\
         </a:clrScheme>",
        name = e(&theme.name),
        dk1  = theme.dk1,  lt1 = theme.lt1,
        dk2  = theme.dk2,  lt2 = theme.lt2,
        a1   = theme.accent1, a2 = theme.accent2,
        a3   = theme.accent3, a4 = theme.accent4,
        a5   = theme.accent5, a6 = theme.accent6,
        hl   = theme.hlink, fh  = theme.fol_hlink,
    )
}

// ---------------------------------------------------------------------------
// Color theme CRUD commands
// ---------------------------------------------------------------------------

#[tauri::command]
fn load_color_themes(folder: String) -> Vec<ColorTheme> {
    let Ok(entries) = std::fs::read_dir(&folder) else { return vec![] };
    let mut themes: Vec<ColorTheme> = entries
        .flatten()
        .filter(|e| {
            e.path()
                .extension()
                .and_then(|x| x.to_str())
                .map(|x| x.eq_ignore_ascii_case("xml"))
                .unwrap_or(false)
        })
        .filter_map(|e| {
            let path = e.path();
            let xml = std::fs::read_to_string(&path).ok()?;
            parse_color_theme(&xml, path.to_str()?)
        })
        .collect();
    themes.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
    themes
}

#[tauri::command]
fn save_color_theme(folder: String, theme: ColorTheme) -> Result<String, String> {
    let dir = PathBuf::from(&folder);
    if !dir.is_dir() {
        return Err(format!("Folder does not exist: {}", folder));
    }
    let path = if theme.file_path.is_empty() {
        dir.join(format!("{}.xml", sanitize_filename(&theme.name)))
    } else {
        PathBuf::from(&theme.file_path)
    };
    let xml = make_color_xml(&theme);
    std::fs::write(&path, xml).map_err(|e| e.to_string())?;
    Ok(path.to_string_lossy().into_owned())
}

#[tauri::command]
fn rename_color_theme(
    folder:   String,
    old_path: String,
    new_name: String,
    mut theme: ColorTheme,
) -> Result<String, String> {
    let dir      = PathBuf::from(&folder);
    let new_path = dir.join(format!("{}.xml", sanitize_filename(&new_name)));
    theme.name      = new_name;
    theme.file_path = new_path.to_string_lossy().into_owned();
    let xml = make_color_xml(&theme);
    std::fs::write(&new_path, xml).map_err(|e| e.to_string())?;
    if !old_path.is_empty() && old_path != new_path.to_string_lossy() {
        let _ = std::fs::remove_file(&old_path);
    }
    Ok(new_path.to_string_lossy().into_owned())
}

// ---------------------------------------------------------------------------
// Font enumeration (ttf-parser)
// ---------------------------------------------------------------------------

fn family_name_from_data(data: &[u8], face_index: u32) -> Option<String> {
    let face = ttf_parser::Face::parse(data, face_index).ok()?;
    let names = face.names();
    for target_id in [16u16, 1u16] {
        if let Some(s) = names
            .into_iter()
            .find(|n| n.name_id == target_id && n.platform_id == ttf_parser::PlatformId::Windows)
            .and_then(|n| n.to_string())
        {
            return Some(s);
        }
        if let Some(s) = names
            .into_iter()
            .find(|n| n.name_id == target_id && n.platform_id == ttf_parser::PlatformId::Macintosh)
            .and_then(|n| n.to_string())
        {
            return Some(s);
        }
    }
    None
}

#[tauri::command]
async fn list_fonts() -> Vec<String> {
    tauri::async_runtime::spawn_blocking(|| {
        let home = std::env::var("HOME").unwrap_or_default();
        let dirs = [
            "/System/Library/Fonts".to_string(),
            "/System/Library/Fonts/Supplemental".to_string(),
            "/Library/Fonts".to_string(),
            format!("{}/Library/Fonts", home),
        ];
        let mut families: BTreeSet<String> = BTreeSet::new();
        for dir in &dirs {
            let Ok(entries) = std::fs::read_dir(dir) else { continue };
            for entry in entries.flatten() {
                let path = entry.path();
                let ext = path
                    .extension()
                    .and_then(|e| e.to_str())
                    .map(|e| e.to_lowercase())
                    .unwrap_or_default();
                if !matches!(ext.as_str(), "ttf" | "otf" | "ttc") { continue }
                let Ok(data) = std::fs::read(&path) else { continue };
                let count = if ext == "ttc" {
                    ttf_parser::fonts_in_collection(&data).unwrap_or(1)
                } else { 1 };
                for i in 0..count {
                    if let Some(name) = family_name_from_data(&data, i) {
                        families.insert(name);
                    }
                }
            }
        }
        families.into_iter().collect()
    }).await.unwrap_or_default()
}

/// Directories where Microsoft Office stores its private fonts on macOS.
const OFFICE_FONT_DIRS: &[&str] = &[
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts",
    "/Applications/Microsoft Excel.app/Contents/Resources/DFonts",
    "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts",
];

/// Returns `{family, path, index}` records for fonts in Office-private dirs.
/// The frontend uses this for the datalist and for on-demand font loading.
#[tauri::command]
async fn list_extra_font_faces() -> Vec<serde_json::Value> {
    tauri::async_runtime::spawn_blocking(|| {
        let mut out: Vec<serde_json::Value> = Vec::new();
        let mut seen_paths: std::collections::HashSet<String> = std::collections::HashSet::new();

        for dir in OFFICE_FONT_DIRS {
            let Ok(entries) = std::fs::read_dir(dir) else { continue };
            for entry in entries.flatten() {
                let path = entry.path();
                let ext = path.extension()
                    .and_then(|e| e.to_str())
                    .map(|e| e.to_lowercase())
                    .unwrap_or_default();
                if !matches!(ext.as_str(), "ttf" | "otf" | "ttc") { continue }
                let path_str = path.to_string_lossy().into_owned();
                if !seen_paths.insert(path_str.clone()) { continue }

                let Ok(data) = std::fs::read(&path) else { continue };
                let count = if ext == "ttc" {
                    ttf_parser::fonts_in_collection(&data).unwrap_or(1)
                } else { 1 };
                for i in 0..count {
                    if let Some(family) = family_name_from_data(&data, i) {
                        out.push(serde_json::json!({
                            "family": family,
                            "path": path_str,
                            "index": i,
                        }));
                    }
                }
            }
        }
        out
    }).await.unwrap_or_default()
}

/// Read a font file and return its bytes as a base64 string so the frontend
/// can construct a `data:` URI for an `@font-face` rule.  This sidesteps
/// WKWebView's block on `file://` URLs from `tauri://localhost` origin.
#[tauri::command]
fn get_font_data_b64(path: String) -> Option<String> {
    let data = std::fs::read(&path).ok()?;
    // Simple base64 encoding without an extra crate
    Some(base64_encode(&data))
}

fn base64_encode(data: &[u8]) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity((data.len() + 2) / 3 * 4);
    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let n = (b0 << 16) | (b1 << 8) | b2;
        out.push(CHARS[((n >> 18) & 63) as usize] as char);
        out.push(CHARS[((n >> 12) & 63) as usize] as char);
        out.push(if chunk.len() > 1 { CHARS[((n >> 6) & 63) as usize] as char } else { '=' });
        out.push(if chunk.len() > 2 { CHARS[(n & 63) as usize] as char } else { '=' });
    }
    out
}

// ---------------------------------------------------------------------------
// Theme CRUD commands — all take `folder` (full path to e.g. "Theme Fonts")
// ---------------------------------------------------------------------------

#[tauri::command]
fn load_themes(folder: String) -> Vec<FontTheme> {
    let Ok(entries) = std::fs::read_dir(&folder) else { return vec![] };
    let mut themes: Vec<FontTheme> = entries
        .flatten()
        .filter(|e| {
            e.path()
                .extension()
                .and_then(|x| x.to_str())
                .map(|x| x.eq_ignore_ascii_case("xml"))
                .unwrap_or(false)
        })
        .filter_map(|e| {
            let path = e.path();
            let xml = std::fs::read_to_string(&path).ok()?;
            parse_theme(&xml, path.to_str()?)
        })
        .collect();
    themes.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
    themes
}

#[tauri::command]
fn save_theme(folder: String, theme: FontTheme) -> Result<String, String> {
    let dir = PathBuf::from(&folder);
    if !dir.is_dir() {
        return Err(format!("Folder does not exist: {}", folder));
    }
    let path = if theme.file_path.is_empty() {
        dir.join(format!("{}.xml", sanitize_filename(&theme.name)))
    } else {
        PathBuf::from(&theme.file_path)
    };
    let xml = make_xml(&theme.name, &theme.heading_font, &theme.body_font);
    std::fs::write(&path, xml).map_err(|e| e.to_string())?;
    Ok(path.to_string_lossy().into_owned())
}

#[tauri::command]
fn delete_theme(file_path: String) -> Result<(), String> {
    std::fs::remove_file(&file_path).map_err(|e| e.to_string())
}

#[tauri::command]
fn rename_theme(
    folder: String,
    old_path: String,
    new_name: String,
    heading_font: String,
    body_font: String,
) -> Result<String, String> {
    let dir = PathBuf::from(&folder);
    let new_path = dir.join(format!("{}.xml", sanitize_filename(&new_name)));
    let xml = make_xml(&new_name, &heading_font, &body_font);
    std::fs::write(&new_path, xml).map_err(|e| e.to_string())?;
    if !old_path.is_empty() && old_path != new_path.to_string_lossy() {
        let _ = std::fs::remove_file(&old_path);
    }
    Ok(new_path.to_string_lossy().into_owned())
}

fn sanitize_filename(name: &str) -> String {
    name.chars()
        .map(|c| if c == '/' || c == '\0' { '_' } else { c })
        .collect()
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

pub fn run() {
    tauri::Builder::default()
        .manage(ThemesRootState(Mutex::new(None)))
        .invoke_handler(tauri::generate_handler![
            get_themes_root,
            set_themes_root,
            pick_themes_folder,
            resolve_theme_subfolder,
            list_fonts,
            list_extra_font_faces,
            get_font_data_b64,
            load_themes,
            save_theme,
            delete_theme,
            rename_theme,
            load_color_themes,
            save_color_theme,
            rename_color_theme,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
