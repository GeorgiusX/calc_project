const STORAGE_KEY = 'calc_notes_v2';
const THEME_KEY = 'calc_theme_v2';
const FX_KEY = 'calc_fx_v2';
const PRECISION_KEY = 'calc_precision_v1';
const WINDOW_SIZE_KEY = 'calc_window_size_v1';

function uid() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`;
}

function deriveTitle(content) {
  const lines = content.split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    if (trimmed.startsWith('#')) return trimmed.replace(/^#\s*/, '').trim();
    const label = trimmed.match(/^([^:]{1,60}):/);
    if (label) return label[1].trim();
    return trimmed.slice(0, 60);
  }
  return 'Untitled note';
}

function seedNotes() {
  const content = [
    '# Pricing',
    'subtotal = 120 EUR',
    'vat = 21%',
    'Total: subtotal + vat on subtotal',
    '',
    '# Timing',
    'today',
    'now in Tokyo',
    '',
    '# Units',
    '3 ft in cm',
    'sum',
  ].join('\n');

  return [
    {
      id: uid(),
      content,
      title: deriveTitle(content),
      updatedAt: Date.now(),
    },
  ];
}

export const StorageAdapter = {
  loadNotes() {
    try {
      const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
      if (Array.isArray(raw) && raw.length) return raw;
    } catch (_) {
      // Ignore parse errors and reseed.
    }
    const seeded = seedNotes();
    this.saveNotes(seeded);
    return seeded;
  },

  saveNotes(notes) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(notes));
  },

  createNote(content = '') {
    return {
      id: uid(),
      content,
      title: deriveTitle(content),
      updatedAt: Date.now(),
    };
  },

  updateNote(note, content) {
    return {
      ...note,
      content,
      title: deriveTitle(content),
      updatedAt: Date.now(),
    };
  },

  duplicateNote(note) {
    return this.createNote(note.content);
  },

  loadTheme() {
    return localStorage.getItem(THEME_KEY) || 'dark';
  },

  saveTheme(theme) {
    localStorage.setItem(THEME_KEY, theme);
  },

  loadFx() {
    try {
      return JSON.parse(localStorage.getItem(FX_KEY) || 'null');
    } catch (_) {
      return null;
    }
  },

  saveFx(snapshot) {
    localStorage.setItem(FX_KEY, JSON.stringify(snapshot));
  },

  loadPrecision() {
    const raw = Number(localStorage.getItem(PRECISION_KEY));
    if (Number.isInteger(raw) && raw >= 0 && raw <= 8) return raw;
    return 3;
  },

  savePrecision(precision) {
    localStorage.setItem(PRECISION_KEY, String(precision));
  },

  loadWindowSize() {
    try {
      const raw = JSON.parse(localStorage.getItem(WINDOW_SIZE_KEY) || 'null');
      if (!raw || typeof raw !== 'object') return null;
      const width = Number(raw.width);
      const height = Number(raw.height);
      if (!Number.isFinite(width) || !Number.isFinite(height)) return null;
      return { width, height };
    } catch (_) {
      return null;
    }
  },

  saveWindowSize(size) {
    localStorage.setItem(WINDOW_SIZE_KEY, JSON.stringify(size));
  },
};
