import {
  createDefaultContext,
  createFxRateProvider,
  evaluateDocument,
  formatResult,
  parseDocument,
  renderGhostHtml,
} from './lib/engine.js?v=9';
import { StorageAdapter } from './lib/storage.js?v=2';

const editor = document.getElementById('editor');
const numiWindow = document.querySelector('.numi-window');
const resultsPane = document.getElementById('results');
const gutter = document.getElementById('gutter');
const ghost = document.getElementById('ghost');
const statusText = document.getElementById('statusText');
const rateInfo = document.getElementById('rateInfo');
const notesList = document.getElementById('notesList');
const activeNoteTitle = document.getElementById('activeNoteTitle');
const newNoteBtn = document.getElementById('newNoteBtn');
const panelToggleBtn = document.getElementById('panelToggleBtn');
const sidePanel = document.getElementById('sidePanel');
const duplicateNoteBtn = document.getElementById('duplicateNoteBtn');
const importBtn = document.getElementById('importBtn');
const fileInput = document.getElementById('fileInput');
const exportBtn = document.getElementById('exportBtn');
const copyResultsBtn = document.getElementById('copyResultsBtn');
const copyFormattedBtn = document.getElementById('copyFormattedBtn');
const libraryBtn = document.getElementById('libraryBtn');
const themeToggleBtn = document.getElementById('themeToggleBtn');
const precisionInput = document.getElementById('precisionInput');
const resizeHandle = document.getElementById('resizeHandle');
const libraryOverlay = document.getElementById('libraryOverlay');
const libraryContent = document.getElementById('libraryContent');
const libraryCloseBtn = document.getElementById('libraryCloseBtn');

const state = {
  notes: StorageAdapter.loadNotes(),
  activeId: null,
  fx: StorageAdapter.loadFx() || { base: 'USD', fetchedAt: 0, date: null, rates: { USD: 1 } },
  precision: StorageAdapter.loadPrecision(),
  windowSize: StorageAdapter.loadWindowSize(),
  context: null,
  isRenamingTitle: false,
};

const fxProvider = createFxRateProvider();

const LIBRARY_SECTIONS = [
  {
    title: 'Core Math',
    summary: 'Basic arithmetic, parentheses, powers, modulo, binary/octal/hex numbers, and natural wording like plus/minus/times/divide.',
    items: ['1 + 2 * 3', '(4 + 5) ^ 2', '10 mod 3', '0xff + 0b10', '20 plus 5'],
  },
  {
    title: 'Variables',
    summary: 'Save values with assignments, reuse them later, and rename notes separately from content.',
    items: ['tax = 21%', 'price = 120 EUR', 'price + tax on price'],
  },
  {
    title: 'Percentages',
    summary: 'Percent of / on / off flows, plus percent comparisons.',
    items: ['20% of 45', '5% on 30 EUR', '6% off 120 USD', '50 as a % of 200'],
  },
  {
    title: 'Rollups',
    summary: 'Use previous result and sum/average over the contiguous numeric block right above.',
    items: ['item1 = 10 USD', 'item2 = 20 USD', 'sum', 'average', 'prev + 5 USD'],
  },
  {
    title: 'Currencies',
    summary: 'ISO codes and symbols work. Use to / in / into for conversion.',
    items: ['10 USD to EUR', '€5 + 3', '1 USD into UAH', 'travel = 450 USD', 'travel in EUR'],
  },
  {
    title: 'Date & Time',
    summary: 'Ask for today/now/time, convert timezones, and use natural city-time phrasing.',
    items: ['today', 'now in Tokyo', 'time in New York', 'New York time', 'fromunix(1700000000)'],
  },
  {
    title: 'Lengths & Areas',
    summary: 'Metric and imperial lengths plus many area aliases.',
    items: ['3 km in m', '12 ft to cm', '1 acre to sq ft', '1 hectare to sq m', '100 sq ft to m2'],
  },
  {
    title: 'Volume & Weight',
    summary: 'Cooking volumes, cubic units, and common mass units.',
    items: ['2 cups to ml', '1 gal to l', '3 stone to kg', '5 lb to oz', '1 m3 to ft3'],
  },
  {
    title: 'Data & CSS Units',
    summary: 'Data sizes and CSS unit conversion with px/pt/em/rem/pc.',
    items: ['1000 MB in GB', '1 GiB to MiB', '24 pt to px', '2 rem to px', '48 px to pt'],
  },
  {
    title: 'Supported Keywords',
    summary: 'Useful built-ins and synonyms the parser understands.',
    items: ['sum, total, average, avg, prev', 'now, time, today, fromunix()', 'plus, minus, times, divide, with, without', 'to, in, into, of, on, off'],
  },
  {
    title: 'Timezone Aliases',
    summary: 'Built-in city shortcuts in addition to standard IANA zones.',
    items: ['New York, NYC, Boston, Miami, Toronto', 'Chicago, Dallas, Denver, Los Angeles, Seattle, Vancouver', 'London, Paris, Madrid, Berlin, Rome, Amsterdam, Zurich', 'Tokyo, Osaka, Seoul, Singapore, Dubai, Delhi, Mumbai, Hong Kong', 'Sydney, Melbourne, Brisbane, Auckland'],
  },
];

function setTheme(theme) {
  document.documentElement.classList.toggle('theme-light', theme === 'light');
  StorageAdapter.saveTheme(theme);
}

function currentTheme() {
  return document.documentElement.classList.contains('theme-light') ? 'light' : 'dark';
}

function applyStoredTheme() {
  setTheme(StorageAdapter.loadTheme());
}

function getActiveNote() {
  return state.notes.find((note) => note.id === state.activeId) || state.notes[0];
}

function saveNotes() {
  StorageAdapter.saveNotes(state.notes);
}

function syncActiveNoteTitle() {
  const active = getActiveNote();
  if (!active || state.isRenamingTitle) return;
  activeNoteTitle.textContent = active.title || 'Untitled note';
}

function renderNotesList() {
  notesList.innerHTML = '';
  for (const note of state.notes.sort((a, b) => b.updatedAt - a.updatedAt)) {
    const row = document.createElement('div');
    row.className = `note-item${note.id === state.activeId ? ' active' : ''}`;

    const selectBtn = document.createElement('button');
    selectBtn.type = 'button';
    selectBtn.className = 'note-select';
    selectBtn.innerHTML = `<div class="note-title">${escapeHtml(note.title)}</div><div class="note-meta">${new Date(note.updatedAt).toLocaleString()}</div>`;
    selectBtn.addEventListener('click', () => {
      state.activeId = note.id;
      loadActiveNote();
    });

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'note-delete';
    deleteBtn.setAttribute('aria-label', `Delete ${note.title}`);
    deleteBtn.textContent = '×';
    deleteBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      deleteNote(note.id);
    });

    row.appendChild(selectBtn);
    row.appendChild(deleteBtn);
    notesList.appendChild(row);
  }
}

function escapeHtml(input) {
  return input.replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch]));
}

function renderLibrary() {
  libraryContent.innerHTML = LIBRARY_SECTIONS.map((section) => `
    <section class="library-section">
      <div class="panel-heading">${escapeHtml(section.title)}</div>
      <p class="library-summary">${escapeHtml(section.summary)}</p>
      <div class="library-examples">
        ${section.items.map((item) => `
          <div class="library-example">
            <code class="library-code">${escapeHtml(item)}</code>
            <button class="snippet-btn library-try-btn" type="button" data-snippet="${escapeHtml(item)}">Try</button>
          </div>
        `).join('')}
      </div>
    </section>
  `).join('');
}

function setLibraryOpen(isOpen) {
  libraryOverlay.classList.toggle('is-hidden', !isOpen);
  libraryOverlay.setAttribute('aria-hidden', String(!isOpen));
}

function setPanelOpen(isOpen) {
  sidePanel.classList.toggle('is-hidden', !isOpen);
  panelToggleBtn.setAttribute('aria-expanded', String(isOpen));
}

function clampWindowSize(size) {
  const maxWidth = Math.min(window.innerWidth - 48, 1400);
  const maxHeight = window.innerHeight - 40;
  return {
    width: Math.max(760, Math.min(maxWidth, size.width)),
    height: Math.max(560, Math.min(maxHeight, size.height)),
  };
}

function applyWindowSize() {
  if (window.innerWidth <= 900) {
    numiWindow.style.width = '';
    numiWindow.style.height = '';
    return;
  }
  if (!state.windowSize) return;
  const next = clampWindowSize(state.windowSize);
  state.windowSize = next;
  numiWindow.style.width = `${next.width}px`;
  numiWindow.style.height = `${next.height}px`;
}

function deleteNote(noteId) {
  const deletingActive = state.activeId === noteId;
  state.notes = state.notes.filter((note) => note.id !== noteId);

  if (!state.notes.length) {
    const fallback = StorageAdapter.createNote('');
    state.notes = [fallback];
  }

  if (deletingActive || !state.notes.some((note) => note.id === state.activeId)) {
    state.activeId = state.notes[0].id;
  }

  saveNotes();
  loadActiveNote({ keepPanelOpen: true });
}

function renderGutter(count) {
  gutter.textContent = Array.from({ length: Math.max(count, 1) }, (_, index) => `${index + 1}`).join('\n');
}

function renderResults(evaluated) {
  resultsPane.innerHTML = '';
  for (const line of evaluated.lines) {
    const row = document.createElement('div');
    row.className = `result-line ${line.error ? 'error' : ''} ${line.kind}`;
    row.textContent = line.error || line.displayText || '';
    resultsPane.appendChild(row);
  }
}

function syncScroll() {
  gutter.scrollTop = editor.scrollTop;
  ghost.scrollTop = editor.scrollTop;
  resultsPane.scrollTop = editor.scrollTop;
}

function setStatus(text, options = {}) {
  statusText.textContent = text;
  statusText.classList.toggle('status-stale', !!options.stale);
}

function updateRateInfo() {
  if (!state.fx?.rates) {
    rateInfo.textContent = '';
    return;
  }
  const age = Date.now() - (state.fx.fetchedAt || 0);
  const stale = age > 1000 * 60 * 60 * 12;
  const date = state.fx.date || new Date(state.fx.fetchedAt || Date.now()).toLocaleDateString();
  rateInfo.textContent = `FX ${state.fx.base} · ${date}${stale ? ' · stale' : ''}`;
}

function evaluateAndRender() {
  const active = getActiveNote();
  if (!active) return;

  const documentModel = parseDocument(editor.value);
  state.context = createDefaultContext({
    fx: state.fx,
    fxProvider,
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    precision: state.precision,
  });

  const evaluated = evaluateDocument(documentModel, state.context);
  ghost.innerHTML = renderGhostHtml(documentModel, evaluated);
  renderGutter(documentModel.lines.length);
  renderResults(evaluated);
  syncActiveNoteTitle();
  updateRateInfo();
  setStatus('Ready');
}

function persistEditor() {
  const active = getActiveNote();
  if (!active) return;
  const updated = StorageAdapter.updateNote(active, editor.value);
  state.notes = state.notes.map((note) => (note.id === active.id ? updated : note));
  saveNotes();
  renderNotesList();
}

function loadActiveNote(options = {}) {
  const active = getActiveNote();
  if (!active) return;
  state.activeId = active.id;
  editor.value = active.content;
  state.isRenamingTitle = false;
  syncActiveNoteTitle();
  renderNotesList();
  evaluateAndRender();
  setPanelOpen(!!options.keepPanelOpen);
}

function commitTitleRename(nextTitle) {
  const active = getActiveNote();
  if (!active) return;
  const updated = StorageAdapter.renameNote(active, nextTitle);
  state.notes = state.notes.map((note) => (note.id === active.id ? updated : note));
  saveNotes();
  renderNotesList();
  state.isRenamingTitle = false;
  syncActiveNoteTitle();
}

function cancelTitleRename() {
  state.isRenamingTitle = false;
  syncActiveNoteTitle();
}

function beginTitleRename() {
  const active = getActiveNote();
  if (!active || state.isRenamingTitle) return;
  state.isRenamingTitle = true;

  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'window-title window-title-input';
  input.value = active.title || 'Untitled note';
  input.setAttribute('aria-label', 'Note title');

  const finish = (save) => {
    input.removeEventListener('blur', onBlur);
    input.removeEventListener('keydown', onKeyDown);
    input.replaceWith(activeNoteTitle);
    if (save) {
      commitTitleRename(input.value);
    } else {
      cancelTitleRename();
    }
  };

  const onBlur = () => finish(true);
  const onKeyDown = (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      finish(true);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      finish(false);
    }
  };

  activeNoteTitle.replaceWith(input);
  input.addEventListener('blur', onBlur);
  input.addEventListener('keydown', onKeyDown);
  input.focus();
  input.select();
}

async function refreshFx() {
  try {
    setStatus('Updating FX…');
    const snapshot = await fxProvider.getRates('USD');
    state.fx = snapshot;
    StorageAdapter.saveFx(snapshot);
    updateRateInfo();
    evaluateAndRender();
  } catch (_) {
    setStatus('Using cached FX', { stale: true });
    updateRateInfo();
  }
}

function appendSnippet(text) {
  const start = editor.selectionStart;
  const end = editor.selectionEnd;
  const value = editor.value;
  const insertion = (start && !value.endsWith('\n') ? '\n' : '') + text;
  editor.value = value.slice(0, start) + insertion + value.slice(end);
  editor.selectionStart = editor.selectionEnd = start + insertion.length;
  persistEditor();
  evaluateAndRender();
  editor.focus();
}

newNoteBtn.addEventListener('click', () => {
  const note = StorageAdapter.createNote('');
  state.notes.unshift(note);
  state.activeId = note.id;
  saveNotes();
  loadActiveNote();
  beginTitleRename();
});

activeNoteTitle.addEventListener('click', beginTitleRename);

panelToggleBtn.addEventListener('click', (event) => {
  event.stopPropagation();
  const isHidden = sidePanel.classList.contains('is-hidden');
  setPanelOpen(isHidden);
});

duplicateNoteBtn.addEventListener('click', () => {
  const active = getActiveNote();
  if (!active) return;
  const note = StorageAdapter.duplicateNote(active);
  state.notes.unshift(note);
  state.activeId = note.id;
  saveNotes();
  loadActiveNote();
});

importBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  const content = await file.text();
  const note = StorageAdapter.createNote(content);
  state.notes.unshift(note);
  state.activeId = note.id;
  saveNotes();
  loadActiveNote();
  fileInput.value = '';
});

exportBtn.addEventListener('click', () => {
  const active = getActiveNote();
  if (!active) return;
  const blob = new Blob([active.content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${active.title || 'note'}.numi`;
  link.click();
  URL.revokeObjectURL(url);
});

copyResultsBtn.addEventListener('click', async () => {
  const documentModel = parseDocument(editor.value);
  const evaluated = evaluateDocument(documentModel, createDefaultContext({ fx: state.fx, fxProvider, precision: state.precision }));
  const output = evaluated.lines.map((line) => line.error || line.displayText || '').join('\n');
  await navigator.clipboard.writeText(output);
  setStatus('Results copied');
});

copyFormattedBtn.addEventListener('click', async () => {
  const documentModel = parseDocument(editor.value);
  const evaluated = evaluateDocument(documentModel, createDefaultContext({ fx: state.fx, fxProvider, precision: state.precision }));
  const output = documentModel.lines.map((line, index) => {
    const result = evaluated.lines[index];
    if (!result || !result.displayText || ['blank', 'header', 'comment', 'text'].includes(line.kind)) return line.raw;
    return `${line.raw} = ${result.error || result.displayText}`;
  }).join('\n');
  await navigator.clipboard.writeText(output);
  setStatus('Formatted note copied');
});

libraryBtn.addEventListener('click', () => setLibraryOpen(true));
libraryCloseBtn.addEventListener('click', () => setLibraryOpen(false));

themeToggleBtn.addEventListener('click', () => setTheme(currentTheme() === 'dark' ? 'light' : 'dark'));

precisionInput.addEventListener('input', () => {
  const next = Math.max(0, Math.min(8, Number.parseInt(precisionInput.value || '0', 10) || 0));
  precisionInput.value = String(next);
  state.precision = next;
  StorageAdapter.savePrecision(next);
  evaluateAndRender();
});

resizeHandle.addEventListener('pointerdown', (event) => {
  if (window.innerWidth <= 900) return;
  event.preventDefault();
  event.stopPropagation();

  const rect = numiWindow.getBoundingClientRect();
  const startX = event.clientX;
  const startY = event.clientY;
  const startWidth = rect.width;
  const startHeight = rect.height;

  resizeHandle.setPointerCapture(event.pointerId);
  document.body.classList.add('is-resizing');

  function onMove(moveEvent) {
    const width = startWidth + (moveEvent.clientX - startX);
    const height = startHeight + (moveEvent.clientY - startY);
    const next = clampWindowSize({ width, height });
    state.windowSize = next;
    numiWindow.style.width = `${next.width}px`;
    numiWindow.style.height = `${next.height}px`;
  }

  function onEnd(endEvent) {
    resizeHandle.releasePointerCapture(endEvent.pointerId);
    resizeHandle.removeEventListener('pointermove', onMove);
    resizeHandle.removeEventListener('pointerup', onEnd);
    resizeHandle.removeEventListener('pointercancel', onEnd);
    document.body.classList.remove('is-resizing');
    if (state.windowSize) StorageAdapter.saveWindowSize(state.windowSize);
  }

  resizeHandle.addEventListener('pointermove', onMove);
  resizeHandle.addEventListener('pointerup', onEnd);
  resizeHandle.addEventListener('pointercancel', onEnd);
});

editor.addEventListener('input', () => {
  persistEditor();
  evaluateAndRender();
});

editor.addEventListener('scroll', syncScroll);
editor.addEventListener('keydown', (event) => {
  if (event.key === 'Tab') {
    event.preventDefault();
    const start = editor.selectionStart;
    const end = editor.selectionEnd;
    const value = editor.value;
    editor.value = `${value.slice(0, start)}\t${value.slice(end)}`;
    editor.selectionStart = editor.selectionEnd = start + 1;
    persistEditor();
    evaluateAndRender();
  }
});

for (const button of document.querySelectorAll('.snippet-btn')) {
  button.addEventListener('click', () => appendSnippet(button.dataset.snippet.replaceAll('&#10;', '\n')));
}

libraryContent.addEventListener('click', (event) => {
  const button = event.target.closest('.library-try-btn');
  if (!button) return;
  appendSnippet(button.dataset.snippet);
  setLibraryOpen(false);
});

document.addEventListener('click', (event) => {
  if (sidePanel.classList.contains('is-hidden')) return;
  if (sidePanel.contains(event.target) || panelToggleBtn.contains(event.target)) return;
  setPanelOpen(false);
});

libraryOverlay.addEventListener('click', (event) => {
  if (event.target === libraryOverlay) setLibraryOpen(false);
});

sidePanel.addEventListener('click', (event) => {
  event.stopPropagation();
});

window.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && !libraryOverlay.classList.contains('is-hidden')) setLibraryOpen(false);
});

window.addEventListener('resize', applyWindowSize);

applyStoredTheme();
renderLibrary();
state.activeId = state.notes[0]?.id || null;
precisionInput.value = String(state.precision);
setPanelOpen(false);
applyWindowSize();
loadActiveNote();
refreshFx();
