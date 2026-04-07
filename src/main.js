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

const appAlert   = (msg)  => showDialog({ message: msg });
const appConfirm = (msg)  => showDialog({ message: msg, confirmLabel: 'Delete',
                                          showCancel: true, destructive: true });

// ── Font variant helpers ──────────────────────────────────────
// Ordered longest-first so "Bold Italic" matches before "Bold"
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

// ── Session state ────────────────────────────────────────────
let themesRoot  = null;   // path to the Themes root folder
let fontsFolder = null;   // = themesRoot + fuzzy("Theme Fonts")

let themes        = [];
let selectedIndex = -1;

// ── Boot ─────────────────────────────────────────────────────

async function init() {
  loadFontsAsync(); // fire and forget

  await resolveThemesRoot();

  document.getElementById('addBtn').addEventListener('click', addTheme);
  document.getElementById('removeBtn').addEventListener('click', removeTheme);
  document.getElementById('selectFolderBtn').addEventListener('click', pickFolder);

  document.getElementById('nameField').addEventListener('change', nameCommitted);
  document.getElementById('headingFont').addEventListener('change', headingChanged);
  document.getElementById('headingFont').addEventListener('blur',   headingChanged);
  document.getElementById('bodyFont').addEventListener('change', bodyChanged);
  document.getElementById('bodyFont').addEventListener('blur',   bodyChanged);

  // Sync pills when the user types directly into a font input
  document.getElementById('headingFont').addEventListener('input', () => {
    syncPills('headingVariants', splitFont(document.getElementById('headingFont').value).style);
  });
  document.getElementById('bodyFont').addEventListener('input', () => {
    syncPills('bodyVariants', splitFont(document.getElementById('bodyFont').value).style);
  });

  // Variant pill clicks
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

  // Fuzzy-resolve "Theme Fonts" under the root
  const resolved = await invoke('resolve_theme_subfolder', {
    themesRoot: root,
    subfolder: 'Theme Fonts',
  });

  if (!resolved) {
    showBanner(true, `"Theme Fonts" subfolder not found inside the selected folder.`);
    setAppEnabled(false);
    return;
  }

  fontsFolder = resolved;
  showBanner(false);
  setAppEnabled(true);
  await loadThemes();
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
  if (msg) {
    document.querySelector('.folder-banner-msg').textContent = msg;
  } else {
    document.querySelector('.folder-banner-msg').textContent = 'Themes folder not found.';
  }
}

function setAppEnabled(on) {
  document.querySelector('.app').style.opacity = on ? '' : '0.4';
  document.querySelector('.app').style.pointerEvents = on ? '' : 'none';
  document.getElementById('addBtn').disabled = !on;
}

async function loadFontsAsync() {
  try {
    const fonts = await invoke('list_fonts');
    const dl   = document.getElementById('fontList');
    const frag = document.createDocumentFragment();
    fonts.forEach(f => {
      const opt = document.createElement('option');
      opt.value = f;
      frag.appendChild(opt);
    });
    dl.appendChild(frag);
  } catch (e) {
    console.warn('Font enumeration failed:', e);
  }
}

// ── Theme list ────────────────────────────────────────────────

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
  aaUpper.style.fontFamily = theme.heading_font || 'inherit';
  const aaLower = document.createElement('span');
  aaLower.textContent = 'a';
  aaLower.style.fontFamily = theme.body_font || 'inherit';
  aaBox.append(aaUpper, aaLower);

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

function updateSample(heading, body) {
  document.getElementById('sampleHeading').style.fontFamily = heading || 'inherit';
  document.getElementById('sampleBody').style.fontFamily    = body    || 'inherit';
}

// ── Field handlers ────────────────────────────────────────────

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

// ── Add / Remove ──────────────────────────────────────────────

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

// ── Helpers ───────────────────────────────────────────────────

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

init();
