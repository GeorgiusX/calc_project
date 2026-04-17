import {
  createDefaultContext,
  createFxRateProvider,
  evaluateDocument,
  formatResult,
  parseDocument,
  renderGhostHtml,
} from './lib/engine.js?v=5';
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
const themeToggleBtn = document.getElementById('themeToggleBtn');
const precisionInput = document.getElementById('precisionInput');
const resizeHandle = document.getElementById('resizeHandle');

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
  loadActiveNote();
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

function loadActiveNote() {
  const active = getActiveNote();
  if (!active) return;
  state.activeId = active.id;
  editor.value = active.content;
  state.isRenamingTitle = false;
  syncActiveNoteTitle();
  renderNotesList();
  evaluateAndRender();
  setPanelOpen(false);
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

document.addEventListener('click', (event) => {
  if (sidePanel.classList.contains('is-hidden')) return;
  if (sidePanel.contains(event.target) || panelToggleBtn.contains(event.target)) return;
  setPanelOpen(false);
});

sidePanel.addEventListener('click', (event) => {
  event.stopPropagation();
});

window.addEventListener('resize', applyWindowSize);

applyStoredTheme();
state.activeId = state.notes[0]?.id || null;
precisionInput.value = String(state.precision);
setPanelOpen(false);
applyWindowSize();
loadActiveNote();
refreshFx();
