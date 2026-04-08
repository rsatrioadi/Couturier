const { invoke } = window.__TAURI__.core;

// ── Dialog helpers (window.confirm/alert don't work in Tauri) ─

function showDialog({ message, confirmLabel = 'OK', showCancel = false, destructive = false }) {
  return new Promise(resolve => {
    const dialog   = document.getElementById('appDialog');
    const msg      = document.getElementById('appDialogMsg');
    const okBtn    = document.getElementById('appDialogOk');
    const cancelBtn = document.getElementById('appDialogCancel');

    msg.textContent = message;
    okBtn.textContent = confirmLabel;
    okBtn.className = destructive ? 'destructive' : '';
    cancelBtn.hidden = !showCancel;

    const finish = (result) => {
      dialog.close();
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      resolve(result);
    };
    const onOk     = () => finish(true);
    const onCancel = () => finish(false);

    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', onCancel);
    dialog.showModal();
  });
}

const appAlert   = (msg) => showDialog({ message: msg });
const appConfirm = (msg) => showDialog({ message: msg, confirmLabel: 'Delete',
                                         showCancel: true, destructive: true });

// ── Font variant helpers ──────────────────────────────────────
const STYLE_SUFFIXES = [
  'Bold Italic', 'Light Italic', 'Medium Italic', 'Semibold Italic', 'Black Italic',
  'Bold', 'Light', 'Medium', 'Semibold', 'Black', 'Thin', 'Italic',
];

function splitFont(full) {
  const s = (full || '').trim();
  for (const suffix of STYLE_SUFFIXES) {
    if (s.toLowerCase().endsWith((' ' + suffix).toLowerCase())) {
      return { family: s.slice(0, -(suffix.length + 1)).trimEnd(), style: suffix };
    }
  }
  return { family: s, style: 'Regular' };
}

function buildFont(family, style) {
  return (!style || style === 'Regular') ? family : `${family} ${style}`;
}

function syncPills(pillsId, style) {
  document.querySelectorAll(`#${pillsId} .variant-pill`).forEach(btn => {
    btn.classList.toggle('active', btn.dataset.variant === style);
  });
}

function setPillsDisabled(pillsId, disabled) {
  document.querySelectorAll(`#${pillsId} .variant-pill`).forEach(btn => {
    btn.disabled = disabled;
  });
}

// ── Color slot definitions ────────────────────────────────────
const COLOR_SLOTS = [
  { field: 'dk1',       label: 'Dark 1',       id: 'clr-dk1' },
  { field: 'lt1',       label: 'Light 1',      id: 'clr-lt1' },
  { field: 'dk2',       label: 'Dark 2',       id: 'clr-dk2' },
  { field: 'lt2',       label: 'Light 2',      id: 'clr-lt2' },
  { field: 'accent1',   label: 'Accent 1',     id: 'clr-accent1' },
  { field: 'accent2',   label: 'Accent 2',     id: 'clr-accent2' },
  { field: 'accent3',   label: 'Accent 3',     id: 'clr-accent3' },
  { field: 'accent4',   label: 'Accent 4',     id: 'clr-accent4' },
  { field: 'accent5',   label: 'Accent 5',     id: 'clr-accent5' },
  { field: 'accent6',   label: 'Accent 6',     id: 'clr-accent6' },
  { field: 'hlink',     label: 'Hyperlink',    id: 'clr-hlink' },
  { field: 'fol_hlink', label: 'Followed Link',id: 'clr-fol_hlink' },
];

// Colors for a freshly-created theme (Office default palette)
const DEFAULT_COLORS = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A9D18E',
  accent4: 'FFC000', accent5: '5A96C8', accent6: '70AD47',
  hlink: '0563C1', fol_hlink: '954F72',
};

// ── Session state ────────────────────────────────────────────
let themesRoot  = null;
let fontsFolder = null;
let colorsFolder = null;

let themes        = [];
let selectedIndex = -1;

let colorThemes        = [];
let colorSelectedIndex = -1;

// ── Boot ─────────────────────────────────────────────────────

async function init() {
  loadFontsAsync(); // fire and forget

  // Tab switching
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const panelId = btn.dataset.panel;
      document.querySelectorAll('.app').forEach(p => { p.hidden = (p.id !== panelId); });
    });
  });

  await resolveThemesRoot();

  // Font panel listeners
  document.getElementById('addBtn').addEventListener('click', addTheme);
  document.getElementById('removeBtn').addEventListener('click', removeTheme);
  document.getElementById('selectFolderBtn').addEventListener('click', pickFolder);
  document.getElementById('revealFontsBtn').addEventListener('click', () => {
    if (fontsFolder) invoke('reveal_folder', { path: fontsFolder });
  });

  document.getElementById('nameField').addEventListener('change', nameCommitted);
  document.getElementById('headingFont').addEventListener('change', headingChanged);
  document.getElementById('headingFont').addEventListener('blur',   headingChanged);
  document.getElementById('bodyFont').addEventListener('change', bodyChanged);
  document.getElementById('bodyFont').addEventListener('blur',   bodyChanged);

  document.getElementById('headingFont').addEventListener('input', () => {
    syncPills('headingVariants', splitFont(document.getElementById('headingFont').value).style);
  });
  document.getElementById('bodyFont').addEventListener('input', () => {
    syncPills('bodyVariants', splitFont(document.getElementById('bodyFont').value).style);
  });

  document.querySelectorAll('#headingVariants .variant-pill').forEach(btn => {
    btn.addEventListener('click', async () => {
      const input = document.getElementById('headingFont');
      input.value = buildFont(splitFont(input.value).family, btn.dataset.variant);
      syncPills('headingVariants', btn.dataset.variant);
      await headingChanged();
    });
  });
  document.querySelectorAll('#bodyVariants .variant-pill').forEach(btn => {
    btn.addEventListener('click', async () => {
      const input = document.getElementById('bodyFont');
      input.value = buildFont(splitFont(input.value).family, btn.dataset.variant);
      syncPills('bodyVariants', btn.dataset.variant);
      await bodyChanged();
    });
  });

  // Color panel listeners
  document.getElementById('colorAddBtn').addEventListener('click', addColorTheme);
  document.getElementById('colorRemoveBtn').addEventListener('click', removeColorTheme);
  document.getElementById('revealColorsBtn').addEventListener('click', () => {
    if (colorsFolder) invoke('reveal_folder', { path: colorsFolder });
  });
  document.getElementById('colorNameField').addEventListener('change', colorNameCommitted);

  COLOR_SLOTS.forEach(slot => {
    document.getElementById(slot.id).addEventListener('change', async () => {
      if (colorSelectedIndex < 0) return;
      const raw = document.getElementById(slot.id).value; // '#rrggbb'
      colorThemes[colorSelectedIndex][slot.field] = raw.slice(1).toUpperCase();
      await saveColor();
      updateColorPreview(colorThemes[colorSelectedIndex]);
      refreshColorRow(colorSelectedIndex);
    });
  });
}

async function resolveThemesRoot() {
  const root = await invoke('get_themes_root');
  if (root) {
    await applyThemesRoot(root);
  } else {
    showBanner(true);
    setAppEnabled(false);
  }
}

async function applyThemesRoot(root) {
  themesRoot = root;

  const fontsResolved = await invoke('resolve_theme_subfolder', {
    themesRoot: root,
    subfolder: 'Theme Fonts',
  });
  const colorsResolved = await invoke('resolve_theme_subfolder', {
    themesRoot: root,
    subfolder: 'Theme Colors',
  });

  if (!fontsResolved) {
    showBanner(true, `"Theme Fonts" subfolder not found inside the selected folder.`);
    setAppEnabled(false);
    return;
  }

  fontsFolder  = fontsResolved;
  colorsFolder = colorsResolved || null;
  showBanner(false);
  setAppEnabled(true);

  await loadThemes();
  if (colorsFolder) await loadColorThemes();
  else setColorDetailEnabled(false);
}

async function pickFolder() {
  const picked = await invoke('pick_themes_folder');
  if (!picked) return;
  try {
    await invoke('set_themes_root', { path: picked });
    await applyThemesRoot(picked);
  } catch (e) {
    showBanner(true, `Invalid folder: ${e}`);
  }
}

function showBanner(visible, msg) {
  const banner = document.getElementById('folderBanner');
  banner.hidden = !visible;
  document.querySelector('.folder-banner-msg').textContent =
    msg || 'Themes folder not found.';
}

function setAppEnabled(on) {
  document.querySelectorAll('.app').forEach(el => {
    el.style.opacity       = on ? '' : '0.4';
    el.style.pointerEvents = on ? '' : 'none';
  });
  document.getElementById('addBtn').disabled      = !on;
  document.getElementById('colorAddBtn').disabled = !on;
}

// Map of font family name (lowercase) -> {path, index} for Office-private fonts.
// Populated once by loadFontsAsync; used by ensureFontLoaded() on demand.
const extraFontMap = new Map();   // key: family.toLowerCase()
const loadedFonts  = new Set();   // families whose @font-face is already injected

async function loadFontsAsync() {
  try {
    const [fonts, extras] = await Promise.all([
      invoke('list_fonts'),
      invoke('list_extra_font_faces'),
    ]);

    // Build the lookup map
    extras.forEach(e => {
      extraFontMap.set(e.family.toLowerCase(), e);
    });

    // Merge into datalist (deduplicated)
    const allFamilies = new Set([...fonts, ...extras.map(e => e.family)]);
    const dl   = document.getElementById('fontList');
    const frag = document.createDocumentFragment();
    allFamilies.forEach(f => {
      const opt = document.createElement('option');
      opt.value = f;
      frag.appendChild(opt);
    });
    dl.appendChild(frag);
  } catch (e) {
    console.warn('Font enumeration failed:', e);
  }
}

/**
 * If `family` is an Office-private font that hasn't been injected yet,
 * fetch its bytes from Rust as base64 and create a data: @font-face rule.
 * Returns when the font is ready (or immediately if already loaded / not extra).
 */
async function ensureFontLoaded(family) {
  if (!family) return;
  const key = family.toLowerCase();
  if (loadedFonts.has(key)) return;
  const entry = extraFontMap.get(key);
  if (!entry) return;  // system font — no action needed

  try {
    const b64 = await invoke('get_font_data_b64', { path: entry.path });
    if (!b64) return;
    const mime = entry.path.toLowerCase().endsWith('.otf') ? 'font/otf' : 'font/truetype';
    const rule = `@font-face { font-family: ${JSON.stringify(family)}; src: url("data:${mime};base64,${b64}"); }`;
    const style = document.createElement('style');
    style.textContent = rule;
    document.head.appendChild(style);
    loadedFonts.add(key);
  } catch (e) {
    console.warn('Could not load font data for', family, e);
  }
}

// ── Font theme list ───────────────────────────────────────────

async function loadThemes() {
  themes = await invoke('load_themes', { folder: fontsFolder });
  renderList();
  if (themes.length > 0) select(0);
  else setDetailEnabled(false);
}

function renderList() {
  const list  = document.getElementById('themeList');
  const empty = document.getElementById('emptyState');
  list.innerHTML = '';
  if (themes.length === 0) { list.appendChild(empty); return; }
  themes.forEach((t, i) => list.appendChild(makeRow(t, i)));
}

function makeRow(theme, i) {
  const row = document.createElement('div');
  row.className = 'theme-row' + (i === selectedIndex ? ' sel' : '');

  const aaBox = document.createElement('div');
  aaBox.className = 'aa-box';
  const aaUpper = document.createElement('span');
  aaUpper.textContent = 'A';
  const aaLower = document.createElement('span');
  aaLower.textContent = 'a';
  aaBox.append(aaUpper, aaLower);

  // Load Office-private fonts asynchronously then apply family names
  Promise.all([
    ensureFontLoaded(theme.heading_font),
    ensureFontLoaded(theme.body_font),
  ]).then(() => {
    aaUpper.style.fontFamily = theme.heading_font || 'inherit';
    aaLower.style.fontFamily = theme.body_font    || 'inherit';
  });

  const info = document.createElement('div');
  info.className = 'theme-info';

  const name = document.createElement('div');
  name.className = 'theme-name';
  name.textContent = theme.name;

  const sub1 = document.createElement('div');
  sub1.className = 'theme-sub';
  sub1.textContent = theme.heading_font || '(system)';

  const sub2 = document.createElement('div');
  sub2.className = 'theme-sub';
  sub2.textContent = theme.body_font || '(system)';

  info.append(name, sub1, sub2);
  row.append(aaBox, info);
  row.addEventListener('click', () => select(i));
  return row;
}

function select(i) {
  selectedIndex = i;
  renderList();
  document.getElementById('removeBtn').disabled = i < 0;
  if (i >= 0 && i < themes.length) {
    const t = themes[i];
    document.getElementById('nameField').value   = t.name;
    document.getElementById('headingFont').value = t.heading_font;
    document.getElementById('bodyFont').value    = t.body_font;
    syncPills('headingVariants', splitFont(t.heading_font).style);
    syncPills('bodyVariants',    splitFont(t.body_font).style);
    updateSample(t.heading_font, t.body_font);
    setDetailEnabled(true);
  }
}

function setDetailEnabled(on) {
  ['nameField', 'headingFont', 'bodyFont'].forEach(id => {
    document.getElementById(id).disabled = !on;
  });
  setPillsDisabled('headingVariants', !on);
  setPillsDisabled('bodyVariants',    !on);
  if (!on) {
    document.getElementById('nameField').value   = '';
    document.getElementById('headingFont').value = '';
    document.getElementById('bodyFont').value    = '';
    syncPills('headingVariants', 'Regular');
    syncPills('bodyVariants',    'Regular');
    updateSample('', '');
  }
}

async function updateSample(heading, body) {
  // Ensure Office-private fonts are loaded before setting font-family
  await Promise.all([ensureFontLoaded(heading), ensureFontLoaded(body)]);
  document.getElementById('sampleHeading').style.fontFamily = heading || 'inherit';
  document.getElementById('sampleBody').style.fontFamily    = body    || 'inherit';
}

// ── Font field handlers ───────────────────────────────────────

async function nameCommitted() {
  if (selectedIndex < 0) return;
  const newName = document.getElementById('nameField').value.trim();
  if (!newName || newName === themes[selectedIndex].name) return;

  if (themes.some((t, i) => i !== selectedIndex && t.name === newName)) {
    await appAlert(`A theme named "${newName}" already exists.`);
    document.getElementById('nameField').value = themes[selectedIndex].name;
    return;
  }

  const t = themes[selectedIndex];
  try {
    const newPath = await invoke('rename_theme', {
      folder:      fontsFolder,
      oldPath:     t.file_path,
      newName,
      headingFont: t.heading_font,
      bodyFont:    t.body_font,
    });
    themes[selectedIndex] = { ...t, name: newName, file_path: newPath };
    resort();
  } catch (e) { await appAlert('Rename failed: ' + e); }
}

async function headingChanged() {
  if (selectedIndex < 0) return;
  const font = document.getElementById('headingFont').value.trim();
  if (font === themes[selectedIndex].heading_font) return;
  themes[selectedIndex].heading_font = font;
  await save();
  updateSample(font, themes[selectedIndex].body_font);
  refreshRow(selectedIndex);
}

async function bodyChanged() {
  if (selectedIndex < 0) return;
  const font = document.getElementById('bodyFont').value.trim();
  if (font === themes[selectedIndex].body_font) return;
  themes[selectedIndex].body_font = font;
  await save();
  updateSample(themes[selectedIndex].heading_font, font);
  refreshRow(selectedIndex);
}

async function save() {
  if (selectedIndex < 0) return;
  try {
    const path = await invoke('save_theme', {
      folder: fontsFolder,
      theme:  themes[selectedIndex],
    });
    themes[selectedIndex].file_path = path;
  } catch (e) { console.error('Save failed:', e); }
}

// ── Font Add / Remove ─────────────────────────────────────────

async function addTheme() {
  let name = 'Custom', n = 2;
  while (themes.some(t => t.name === name)) name = `Custom ${n++}`;

  const newTheme = {
    name,
    heading_font: 'Helvetica Neue',
    body_font:    'Helvetica Neue',
    file_path:    '',
  };

  try {
    const path = await invoke('save_theme', { folder: fontsFolder, theme: newTheme });
    newTheme.file_path = path;
    themes.push(newTheme);
    themes.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
    selectedIndex = themes.findIndex(t => t.file_path === path);
    renderList();
    select(selectedIndex);
    const f = document.getElementById('nameField');
    f.focus(); f.select();
  } catch (e) { await appAlert('Could not create theme: ' + e); }
}

async function removeTheme() {
  if (selectedIndex < 0) return;
  const t = themes[selectedIndex];
  const confirmed = await appConfirm(`Delete "${t.name}"?\n\nThis will permanently remove the file.`);
  if (!confirmed) return;

  try {
    await invoke('delete_theme', { filePath: t.file_path });
    themes.splice(selectedIndex, 1);
    selectedIndex = Math.min(selectedIndex, themes.length - 1);
    renderList();
    if (themes.length > 0) select(selectedIndex);
    else {
      selectedIndex = -1;
      setDetailEnabled(false);
      document.getElementById('removeBtn').disabled = true;
    }
  } catch (e) { await appAlert('Delete failed: ' + e); }
}

// ── Font helpers ──────────────────────────────────────────────

function refreshRow(i) {
  const list = document.getElementById('themeList');
  const old  = list.children[i];
  if (old) list.replaceChild(makeRow(themes[i], i), old);
}

function resort() {
  const fp = themes[selectedIndex]?.file_path;
  themes.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
  selectedIndex = themes.findIndex(t => t.file_path === fp);
  renderList();
  if (selectedIndex >= 0) select(selectedIndex);
}

// ── Color theme list ──────────────────────────────────────────

async function loadColorThemes() {
  colorThemes = await invoke('load_color_themes', { folder: colorsFolder });
  renderColorList();
  if (colorThemes.length > 0) selectColor(0);
  else setColorDetailEnabled(false);
}

function renderColorList() {
  const list  = document.getElementById('colorThemeList');
  const empty = document.getElementById('colorEmptyState');
  list.innerHTML = '';
  if (colorThemes.length === 0) { list.appendChild(empty); return; }
  colorThemes.forEach((t, i) => list.appendChild(makeColorRow(t, i)));
}

function makeColorRow(theme, i) {
  const row = document.createElement('div');
  row.className = 'theme-row color-theme-row' + (i === colorSelectedIndex ? ' sel' : '');

  const name = document.createElement('div');
  name.className = 'theme-name';
  name.textContent = theme.name;

  // A single strip of 8 color squares: dk1, lt1, dk2, lt2, accent1-4
  const strip = document.createElement('div');
  strip.className = 'clr-strip';
  const keys = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6'];
  keys.forEach(k => {
    const cell = document.createElement('div');
    cell.className = 'clr-strip-cell';
    cell.style.background = '#' + theme[k];
    strip.appendChild(cell);
  });

  row.append(name, strip);
  row.addEventListener('click', () => selectColor(i));
  return row;
}

function selectColor(i) {
  colorSelectedIndex = i;
  renderColorList();
  document.getElementById('colorRemoveBtn').disabled = i < 0;
  if (i >= 0 && i < colorThemes.length) {
    const t = colorThemes[i];
    document.getElementById('colorNameField').value = t.name;
    COLOR_SLOTS.forEach(slot => {
      document.getElementById(slot.id).value = '#' + t[slot.field].toLowerCase();
    });
    updateColorPreview(t);
    setColorDetailEnabled(true);
  }
}

function setColorDetailEnabled(on) {
  document.getElementById('colorNameField').disabled = !on;
  COLOR_SLOTS.forEach(slot => {
    document.getElementById(slot.id).disabled = !on;
  });
  if (!on) {
    document.getElementById('colorNameField').value = '';
    COLOR_SLOTS.forEach(slot => {
      document.getElementById(slot.id).value = '#000000';
    });
    clearColorPreview();
  }
}

function updateColorPreview(theme) {
  if (!theme) return;
  // Dark slide: dk1 background, lt1 heading, lt2 subtext
  const dark = document.getElementById('prevSlideDark');
  dark.style.background = '#' + theme.dk1;
  document.getElementById('prevDkBig').style.color  = '#' + theme.lt1;
  document.getElementById('prevDkText').style.color = '#' + theme.lt2;

  // Light slide: lt1 background, dk1 heading, dk2 subtext
  const light = document.getElementById('prevSlideLight');
  light.style.background = '#' + theme.lt1;
  document.getElementById('prevLtBig').style.color  = '#' + theme.dk1;
  document.getElementById('prevLtText').style.color = '#' + theme.dk2;

  // Accent swatches
  const accentsEl = document.getElementById('prevAccents');
  accentsEl.innerHTML = '';
  ['accent1','accent2','accent3','accent4','accent5','accent6'].forEach(key => {
    const sw = document.createElement('div');
    sw.className = 'preview-accent-swatch';
    sw.style.background = '#' + theme[key];
    accentsEl.appendChild(sw);
  });

  // Links
  document.getElementById('prevHlink').style.color    = '#' + theme.hlink;
  document.getElementById('prevFolhlink').style.color = '#' + theme.fol_hlink;
}

function clearColorPreview() {
  ['prevSlideDark','prevSlideLight'].forEach(id => {
    document.getElementById(id).style.background = '';
  });
  ['prevDkBig','prevDkText','prevLtBig','prevLtText'].forEach(id => {
    document.getElementById(id).style.color = '';
  });
  document.getElementById('prevAccents').innerHTML = '';
  document.getElementById('prevHlink').style.color    = '';
  document.getElementById('prevFolhlink').style.color = '';
}

// ── Color field handlers ──────────────────────────────────────

async function colorNameCommitted() {
  if (colorSelectedIndex < 0) return;
  const newName = document.getElementById('colorNameField').value.trim();
  if (!newName || newName === colorThemes[colorSelectedIndex].name) return;

  if (colorThemes.some((t, i) => i !== colorSelectedIndex && t.name === newName)) {
    await appAlert(`A color theme named "${newName}" already exists.`);
    document.getElementById('colorNameField').value = colorThemes[colorSelectedIndex].name;
    return;
  }

  const t = colorThemes[colorSelectedIndex];
  try {
    const newPath = await invoke('rename_color_theme', {
      folder:  colorsFolder,
      oldPath: t.file_path,
      newName,
      theme:   t,
    });
    colorThemes[colorSelectedIndex] = { ...t, name: newName, file_path: newPath };
    resortColors();
  } catch (e) { await appAlert('Rename failed: ' + e); }
}

async function saveColor() {
  if (colorSelectedIndex < 0) return;
  try {
    const path = await invoke('save_color_theme', {
      folder: colorsFolder,
      theme:  colorThemes[colorSelectedIndex],
    });
    colorThemes[colorSelectedIndex].file_path = path;
  } catch (e) { console.error('Color save failed:', e); }
}

// ── Color Add / Remove ────────────────────────────────────────

async function addColorTheme() {
  if (!colorsFolder) {
    await appAlert('The "Theme Colors" subfolder was not found.\nPlease ensure it exists inside your Themes folder.');
    return;
  }
  let name = 'Custom', n = 2;
  while (colorThemes.some(t => t.name === name)) name = `Custom ${n++}`;

  const newTheme = { name, file_path: '', ...DEFAULT_COLORS };

  try {
    const path = await invoke('save_color_theme', { folder: colorsFolder, theme: newTheme });
    newTheme.file_path = path;
    colorThemes.push(newTheme);
    colorThemes.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
    colorSelectedIndex = colorThemes.findIndex(t => t.file_path === path);
    renderColorList();
    selectColor(colorSelectedIndex);
    const f = document.getElementById('colorNameField');
    f.focus(); f.select();
  } catch (e) { await appAlert('Could not create color theme: ' + e); }
}

async function removeColorTheme() {
  if (colorSelectedIndex < 0) return;
  const t = colorThemes[colorSelectedIndex];
  const confirmed = await appConfirm(`Delete "${t.name}"?\n\nThis will permanently remove the file.`);
  if (!confirmed) return;

  try {
    await invoke('delete_theme', { filePath: t.file_path });
    colorThemes.splice(colorSelectedIndex, 1);
    colorSelectedIndex = Math.min(colorSelectedIndex, colorThemes.length - 1);
    renderColorList();
    if (colorThemes.length > 0) selectColor(colorSelectedIndex);
    else {
      colorSelectedIndex = -1;
      setColorDetailEnabled(false);
      document.getElementById('colorRemoveBtn').disabled = true;
    }
  } catch (e) { await appAlert('Delete failed: ' + e); }
}

// ── Color helpers ─────────────────────────────────────────────

function refreshColorRow(i) {
  const list = document.getElementById('colorThemeList');
  const old  = list.children[i];
  if (old) list.replaceChild(makeColorRow(colorThemes[i], i), old);
}

function resortColors() {
  const fp = colorThemes[colorSelectedIndex]?.file_path;
  colorThemes.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
  colorSelectedIndex = colorThemes.findIndex(t => t.file_path === fp);
  renderColorList();
  if (colorSelectedIndex >= 0) selectColor(colorSelectedIndex);
}

init();
