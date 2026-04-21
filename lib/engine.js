const OPERATOR_WORDS = new Map([
  ['plus', '+'],
  ['and', '+'],
  ['minus', '-'],
  ['times', '*'],
  ['multiplied', '*'],
  ['multiply', '*'],
  ['divide', '/'],
  ['divided', '/'],
  ['with', '+'],
  ['without', '-'],
  ['mod', 'mod'],
  ['xor', 'xor'],
]);

const CURRENCY_SYMBOLS = {
  '$': 'USD',
  '€': 'EUR',
  '£': 'GBP',
  '¥': 'JPY',
  '₹': 'INR',
  '₩': 'KRW',
  'A$': 'AUD',
  'C$': 'CAD',
  'R$': 'BRL',
  '₿': 'BTC',
};

const CRYPTO_CODES = new Set([
  'BTC', 'ETH', 'SOL', 'BNB', 'XRP', 'ADA', 'DOGE', 'DOT',
  'AVAX', 'MATIC', 'LINK', 'UNI', 'USDT', 'USDC', 'LTC', 'TRX',
]);

const COINGECKO_IDS = {
  bitcoin: 'BTC',
  ethereum: 'ETH',
  solana: 'SOL',
  binancecoin: 'BNB',
  ripple: 'XRP',
  cardano: 'ADA',
  dogecoin: 'DOGE',
  polkadot: 'DOT',
  'avalanche-2': 'AVAX',
  'matic-network': 'MATIC',
  chainlink: 'LINK',
  uniswap: 'UNI',
  tether: 'USDT',
  'usd-coin': 'USDC',
  litecoin: 'LTC',
  tron: 'TRX',
};

const BASE_DIMENSIONS = ['length', 'mass', 'duration', 'angle', 'data', 'css'];
const CURRENT_TZ = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';

function zeroDims() {
  return { length: 0, mass: 0, duration: 0, angle: 0, data: 0, css: 0 };
}

function normalizeDims(input) {
  const out = zeroDims();
  if (!input) return out;
  for (const key of BASE_DIMENSIONS) out[key] = input[key] || 0;
  return out;
}

function addDims(left, right, direction = 1) {
  const out = zeroDims();
  for (const key of BASE_DIMENSIONS) out[key] = (left[key] || 0) + direction * (right[key] || 0);
  return out;
}

function dimsEqual(a, b) {
  return BASE_DIMENSIONS.every((key) => (a[key] || 0) === (b[key] || 0));
}

function hasDims(dims) {
  return BASE_DIMENSIONS.some((key) => dims[key]);
}

function dimsToLabel(dims) {
  const parts = [];
  for (const key of BASE_DIMENSIONS) {
    const power = dims[key] || 0;
    if (!power) continue;
    const label = key === 'duration' ? 's' : key === 'mass' ? 'g' : key === 'data' ? 'B' : key === 'angle' ? 'rad' : key === 'css' ? 'px' : 'm';
    parts.push(power === 1 ? label : `${label}^${power}`);
  }
  return parts.join('·');
}

function makeNumber(value, options = {}) {
  return { kind: 'number', value, format: options.format || null };
}

function makePercentage(ratio) {
  return { kind: 'percentage', ratio };
}

function makeCurrency(amount, code) {
  return { kind: 'currency', amount, code: code.toUpperCase() };
}

function makeQuantity(baseValue, dims, preferredUnit = null, format = null) {
  return { kind: 'quantity', baseValue, dims: normalizeDims(dims), preferredUnit, format };
}

function makeTemperature(kelvin, unit = 'C') {
  return { kind: 'temperature', kelvin, unit };
}

function makeDuration(ms, preferredUnit = 'h') {
  return { kind: 'duration', ms, preferredUnit };
}

function makeDateTime(ts, timezone = CURRENT_TZ) {
  return { kind: 'datetime', ts, timezone };
}

function makeText(text) {
  return { kind: 'text', text };
}

function cloneValue(value) {
  return value ? structuredClone(value) : value;
}

function isNumericLike(value) {
  return value && ['number', 'percentage', 'currency', 'quantity', 'temperature', 'duration'].includes(value.kind);
}

function createUnitRegistry() {
  const units = new Map();
  const aliases = new Map();

  function addUnit(def) {
    const entry = {
      ...def,
      dims: normalizeDims(def.dims),
      aliases: def.aliases.map((alias) => alias.toLowerCase()),
    };
    units.set(def.id, entry);
    for (const alias of entry.aliases) aliases.set(alias, entry);
  }

  function addSimple(id, symbol, factor, dims, aliases, options = {}) {
    addUnit({ id, symbol, factor, dims, aliases, ...options });
  }

  addSimple('nm', 'nm', 1e-9, { length: 1 }, ['nm', 'nanometer', 'nanometers', 'nanometre', 'nanometres']);
  addSimple('um', 'um', 1e-6, { length: 1 }, ['um', 'micrometer', 'micrometers', 'micrometre', 'micrometres', 'micron', 'microns']);
  addSimple('mm', 'mm', 0.001, { length: 1 }, ['mm', 'millimeter', 'millimeters', 'millimetre', 'millimetres']);
  addSimple('cm', 'cm', 0.01, { length: 1 }, ['cm', 'centimeter', 'centimeters', 'centimetre', 'centimetres']);
  addSimple('m', 'm', 1, { length: 1 }, ['m', 'meter', 'meters', 'metre', 'metres']);
  addSimple('km', 'km', 1000, { length: 1 }, ['km', 'kilometer', 'kilometers', 'kilometre', 'kilometres']);
  addSimple('in', 'in', 0.0254, { length: 1 }, ['in', 'inch', 'inches']);
  addSimple('ft', 'ft', 0.3048, { length: 1 }, ['ft', 'foot', 'feet']);
  addSimple('yd', 'yd', 0.9144, { length: 1 }, ['yd', 'yard', 'yards']);
  addSimple('mi', 'mi', 1609.344, { length: 1 }, ['mi', 'mile', 'miles']);
  addSimple('nmi', 'nmi', 1852, { length: 1 }, ['nmi', 'nautical mile', 'nautical miles']);

  addSimple('mg', 'mg', 0.001, { mass: 1 }, ['mg', 'milligram', 'milligrams']);
  addSimple('g', 'g', 1, { mass: 1 }, ['g', 'gram', 'grams']);
  addSimple('kg', 'kg', 1000, { mass: 1 }, ['kg', 'kilogram', 'kilograms']);
  addSimple('lb', 'lb', 453.59237, { mass: 1 }, ['lb', 'lbs', 'pound', 'pounds']);
  addSimple('oz', 'oz', 28.349523125, { mass: 1 }, ['oz', 'ounce', 'ounces']);
  addSimple('st', 'st', 6350.29318, { mass: 1 }, ['st', 'stone', 'stones']);
  addSimple('ton', 'ton', 907184.74, { mass: 1 }, ['ton', 'tons', 'short ton', 'short tons']);
  addSimple('tonne', 't', 1000000, { mass: 1 }, ['t', 'tonne', 'tonnes', 'metric ton', 'metric tons']);

  addSimple('ms', 'ms', 1, { duration: 1 }, ['ms', 'millisecond', 'milliseconds'], { durationMs: 1 });
  addSimple('s', 's', 1000, { duration: 1 }, ['s', 'sec', 'secs', 'second', 'seconds'], { durationMs: 1000 });
  addSimple('min', 'min', 60000, { duration: 1 }, ['min', 'mins', 'minute', 'minutes'], { durationMs: 60000 });
  addSimple('h', 'h', 3600000, { duration: 1 }, ['h', 'hr', 'hrs', 'hour', 'hours'], { durationMs: 3600000 });
  addSimple('day', 'd', 86400000, { duration: 1 }, ['day', 'days', 'd'], { durationMs: 86400000 });
  addSimple('week', 'wk', 604800000, { duration: 1 }, ['week', 'weeks', 'wk'], { durationMs: 604800000 });
  addSimple('month', 'mo', 2628000000, { duration: 1 }, ['month', 'months', 'mo'], { durationMs: 2628000000 });
  addSimple('year', 'yr', 31536000000, { duration: 1 }, ['year', 'years', 'yr'], { durationMs: 31536000000 });

  addSimple('rad', 'rad', 1, { angle: 1 }, ['rad', 'radian', 'radians']);
  addSimple('deg', '°', Math.PI / 180, { angle: 1 }, ['deg', 'degree', 'degrees']);
  addSimple('grad', 'grad', Math.PI / 200, { angle: 1 }, ['grad', 'grads', 'gon', 'gons']);
  addSimple('turn', 'turn', Math.PI * 2, { angle: 1 }, ['turn', 'turns', 'rev', 'revs', 'revolution', 'revolutions']);

  addSimple('bit', 'bit', 0.125, { data: 1 }, ['bit', 'bits']);
  addSimple('B', 'B', 1, { data: 1 }, ['b', 'byte', 'bytes']);
  addSimple('KB', 'KB', 1000, { data: 1 }, ['kb', 'kilobyte', 'kilobytes']);
  addSimple('MB', 'MB', 1000 ** 2, { data: 1 }, ['mb', 'megabyte', 'megabytes']);
  addSimple('GB', 'GB', 1000 ** 3, { data: 1 }, ['gb', 'gigabyte', 'gigabytes']);
  addSimple('TB', 'TB', 1000 ** 4, { data: 1 }, ['tb', 'terabyte', 'terabytes']);
  addSimple('PB', 'PB', 1000 ** 5, { data: 1 }, ['pb', 'petabyte', 'petabytes']);
  addSimple('KiB', 'KiB', 1024, { data: 1 }, ['kib', 'kibibyte', 'kibibytes']);
  addSimple('MiB', 'MiB', 1024 ** 2, { data: 1 }, ['mib', 'mebibyte', 'mebibytes']);
  addSimple('GiB', 'GiB', 1024 ** 3, { data: 1 }, ['gib', 'gibibyte', 'gibibytes']);
  addSimple('TiB', 'TiB', 1024 ** 4, { data: 1 }, ['tib', 'tebibyte', 'tebibytes']);

  addSimple('px', 'px', 1, { css: 1 }, ['px', 'pixel', 'pixels']);
  addSimple('pt', 'pt', 96 / 72, { css: 1 }, ['pt', 'point', 'points']);
  addSimple('em', 'em', 16, { css: 1 }, ['em']);
  addSimple('pc', 'pc', 16, { css: 1 }, ['pc', 'pica', 'picas']);
  addSimple('rem', 'rem', 16, { css: 1 }, ['rem']);

  addSimple('mm2', 'mm²', 0.000001, { length: 2 }, ['mm2', 'sq mm', 'square millimeter', 'square millimeters']);
  addSimple('m2', 'm²', 1, { length: 2 }, ['m2', 'sqm', 'sq m', 'square meter', 'square meters', 'square metre', 'square metres']);
  addSimple('cm2', 'cm²', 0.0001, { length: 2 }, ['cm2', 'sq cm', 'square centimeter', 'square centimeters', 'square centimetre', 'square centimetres']);
  addSimple('km2', 'km²', 1000000, { length: 2 }, ['km2', 'sq km', 'square kilometer', 'square kilometers', 'square kilometre', 'square kilometres']);
  addSimple('in2', 'in²', 0.00064516, { length: 2 }, ['in2', 'sq in', 'square inch', 'square inches']);
  addSimple('ft2', 'ft²', 0.09290304, { length: 2 }, ['ft2', 'sq ft', 'square foot', 'square feet']);
  addSimple('yd2', 'yd²', 0.83612736, { length: 2 }, ['yd2', 'sq yd', 'square yard', 'square yards']);
  addSimple('mi2', 'mi²', 2589988.110336, { length: 2 }, ['mi2', 'sq mi', 'square mile', 'square miles']);
  addSimple('ha', 'ha', 10000, { length: 2 }, ['ha', 'hectare', 'hectares']);
  addSimple('acre', 'ac', 4046.8564224, { length: 2 }, ['acre', 'acres', 'ac']);
  addSimple('m3', 'm³', 1, { length: 3 }, ['m3', 'cbm', 'cubic meter', 'cubic meters', 'cubic metre', 'cubic metres']);
  addSimple('cm3', 'cm³', 0.000001, { length: 3 }, ['cm3', 'cc', 'cubic centimeter', 'cubic centimeters', 'cubic centimetre', 'cubic centimetres']);
  addSimple('in3', 'in³', 0.000016387064, { length: 3 }, ['in3', 'cubic inch', 'cubic inches']);
  addSimple('ft3', 'ft³', 0.028316846592, { length: 3 }, ['ft3', 'cubic foot', 'cubic feet']);
  addSimple('yd3', 'yd³', 0.764554857984, { length: 3 }, ['yd3', 'cubic yard', 'cubic yards']);
  addSimple('l', 'L', 0.001, { length: 3 }, ['l', 'liter', 'liters', 'litre', 'litres']);
  addSimple('dl', 'dL', 0.0001, { length: 3 }, ['dl', 'deciliter', 'deciliters', 'decilitre', 'decilitres']);
  addSimple('cl', 'cL', 0.00001, { length: 3 }, ['cl', 'centiliter', 'centiliters', 'centilitre', 'centilitres']);
  addSimple('ml', 'mL', 0.000001, { length: 3 }, ['ml', 'milliliter', 'milliliters', 'millilitre', 'millilitres']);
  addSimple('floz', 'fl oz', 0.0000295735295625, { length: 3 }, ['floz', 'fl oz', 'fluid ounce', 'fluid ounces']);
  addSimple('cup', 'cup', 0.0002365882365, { length: 3 }, ['cup', 'cups']);
  addSimple('pt_us', 'pt', 0.000473176473, { length: 3 }, ['pt us', 'pint', 'pints']);
  addSimple('qt', 'qt', 0.000946352946, { length: 3 }, ['qt', 'quart', 'quarts']);
  addSimple('tbsp', 'tbsp', 0.00001478676478125, { length: 3 }, ['tbsp', 'tablespoon', 'tablespoons']);
  addSimple('tsp', 'tsp', 0.00000492892159375, { length: 3 }, ['tsp', 'teaspoon', 'teaspoons']);
  addSimple('gal', 'gal', 0.003785411784, { length: 3 }, ['gal', 'gallon', 'gallons']);

  const temperatures = new Map([
    ['c', { id: 'C', symbol: '°C', aliases: ['c', 'celsius'], toKelvin: (n) => n + 273.15, fromKelvin: (n) => n - 273.15 }],
    ['f', { id: 'F', symbol: '°F', aliases: ['f', 'fahrenheit'], toKelvin: (n) => ((n - 32) * 5) / 9 + 273.15, fromKelvin: (n) => ((n - 273.15) * 9) / 5 + 32 }],
    ['k', { id: 'K', symbol: 'K', aliases: ['k', 'kelvin'], toKelvin: (n) => n, fromKelvin: (n) => n }],
  ]);

  return {
    get(alias) {
      return aliases.get(alias.toLowerCase()) || null;
    },
    isTemperature(alias) {
      return temperatures.get(alias.toLowerCase()) || null;
    },
    units,
  };
}

const UNIT_REGISTRY = createUnitRegistry();

const BUILT_IN_TIMEZONES = new Map([
  ['utc', 'UTC'],
  ['gmt', 'UTC'],
  ['lisbon', 'Europe/Lisbon'],
  ['madrid', 'Europe/Madrid'],
  ['berlin', 'Europe/Berlin'],
  ['rome', 'Europe/Rome'],
  ['amsterdam', 'Europe/Amsterdam'],
  ['vienna', 'Europe/Vienna'],
  ['zurich', 'Europe/Zurich'],
  ['tokyo', 'Asia/Tokyo'],
  ['osaka', 'Asia/Tokyo'],
  ['seoul', 'Asia/Seoul'],
  ['singapore', 'Asia/Singapore'],
  ['dubai', 'Asia/Dubai'],
  ['delhi', 'Asia/Kolkata'],
  ['mumbai', 'Asia/Kolkata'],
  ['kolkata', 'Asia/Kolkata'],
  ['bangkok', 'Asia/Bangkok'],
  ['sydney', 'Australia/Sydney'],
  ['melbourne', 'Australia/Melbourne'],
  ['brisbane', 'Australia/Brisbane'],
  ['auckland', 'Pacific/Auckland'],
  ['london', 'Europe/London'],
  ['paris', 'Europe/Paris'],
  ['new york', 'America/New_York'],
  ['nyc', 'America/New_York'],
  ['boston', 'America/New_York'],
  ['miami', 'America/New_York'],
  ['toronto', 'America/Toronto'],
  ['chicago', 'America/Chicago'],
  ['dallas', 'America/Chicago'],
  ['denver', 'America/Denver'],
  ['los angeles', 'America/Los_Angeles'],
  ['la', 'America/Los_Angeles'],
  ['san francisco', 'America/Los_Angeles'],
  ['vancouver', 'America/Vancouver'],
  ['seattle', 'America/Los_Angeles'],
  ['pst', 'America/Los_Angeles'],
  ['cst', 'America/Chicago'],
  ['mst', 'America/Denver'],
  ['est', 'America/New_York'],
  ['edt', 'America/New_York'],
  ['cet', 'Europe/Paris'],
  ['cest', 'Europe/Paris'],
  ['hkt', 'Asia/Hong_Kong'],
  ['hong kong', 'Asia/Hong_Kong'],
]);

export function createFxRateProvider(fetchImpl = globalThis.fetch) {
  async function fetchCryptoRates(usdRates) {
    const ids = Object.keys(COINGECKO_IDS).join(',');
    const url = `https://api.coingecko.com/api/v3/simple/price?ids=${ids}&vs_currencies=usd`;
    try {
      const response = await fetchImpl(url);
      if (!response.ok) return;
      const data = await response.json();
      for (const [geckoId, ticker] of Object.entries(COINGECKO_IDS)) {
        const priceUsd = data[geckoId]?.usd;
        if (priceUsd && priceUsd > 0) {
          usdRates[ticker] = 1 / priceUsd;
        }
      }
    } catch (_) {
      // Crypto fetch failed — fiat rates still work fine.
    }
  }

  return {
    async getRates(base = 'USD') {
      const candidates = [
        `https://open.er-api.com/v6/latest/${encodeURIComponent(base)}`,
        `https://api.frankfurter.app/latest?from=${encodeURIComponent(base)}`,
      ];

      for (const url of candidates) {
        try {
          const response = await fetchImpl(url);
          if (!response.ok) continue;
          const data = await response.json();
          const rates = data.rates || {};
          const snapshot = {
            base: (data.base_code || data.base || base).toUpperCase(),
            fetchedAt: Date.now(),
            date: data.time_last_update_utc || data.date || null,
            rates: {},
          };
          for (const [code, value] of Object.entries(rates)) snapshot.rates[code.toUpperCase()] = value;
          snapshot.rates[snapshot.base] = 1;
          await fetchCryptoRates(snapshot.rates);
          return snapshot;
        } catch (_) {
          // Try the next provider.
        }
      }

      throw new Error('Failed to fetch FX rates');
    },
  };
}

export function createDateTimeProvider(now = () => Date.now()) {
  return {
    parseExpression(input, locale = 'en-US', timezone = CURRENT_TZ) {
      const normalized = input.trim().toLowerCase();
      if (normalized === 'now' || normalized === 'time') return makeDateTime(now(), timezone);
      if (normalized === 'today') {
        const date = new Date(now());
        date.setHours(0, 0, 0, 0);
        return makeDateTime(date.getTime(), timezone);
      }
      return new Error(`Unsupported date/time expression: ${input}`);
    },
    fromUnix(value, timezone = CURRENT_TZ) {
      return makeDateTime(value * 1000, timezone);
    },
    resolveTimezone(input) {
      const direct = BUILT_IN_TIMEZONES.get(input.trim().toLowerCase());
      if (direct) return direct;
      try {
        new Intl.DateTimeFormat('en-US', { timeZone: input });
        return input;
      } catch (_) {
        return null;
      }
    },
    now,
    locale: 'en-US',
  };
}

export function createDefaultContext(overrides = {}) {
  const dateTime = overrides.dateTimeProvider || createDateTimeProvider();
  return {
    variables: overrides.variables ? new Map(overrides.variables) : new Map(),
    lineValues: overrides.lineValues ? [...overrides.lineValues] : [],
    timezone: overrides.timezone || CURRENT_TZ,
    locale: overrides.locale || 'en-US',
    fx: overrides.fx || { base: 'USD', fetchedAt: 0, date: null, rates: { USD: 1 } },
    fxProvider: overrides.fxProvider || createFxRateProvider(),
    dateTimeProvider: dateTime,
    settings: {
      ppi: 96,
      emPx: 16,
      precision: overrides.precision ?? 3,
    },
  };
}

function stripQuotedSegments(input) {
  return input.replace(/"[^"]*"/g, ' ');
}

function escapeHtml(input) {
  return input.replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch]));
}

function deriveLabelTitle(raw) {
  return raw.replace(/^#\s*/, '').trim();
}

function looksLikeExpression(input) {
  const probe = stripQuotedSegments(input).trim();
  if (!probe) return false;
  if (/[0-9$€£¥₹₩=()+\-*/%^]/.test(probe)) return true;
  return /\b(sum|total|average|avg|prev|today|now|time|fromunix|sqrt|sin|cos|tan|round|root|as|of|on|off|in|to|usd|eur|gbp|km|cm|hour|hours|min|day)\b/i.test(probe);
}

function classifyLine(raw, index) {
  const trimmed = raw.trim();
  if (!trimmed) return { index, raw, kind: 'blank' };
  if (trimmed.startsWith('#')) return { index, raw, kind: 'header', text: deriveLabelTitle(trimmed) };
  if (trimmed.startsWith('//')) return { index, raw, kind: 'comment', text: trimmed.slice(2).trim() };

  const stripped = stripQuotedSegments(raw);
  const labelMatch = stripped.match(/^(\s*[^:=#/"'][^:]*?)\s*:\s*(.+)$/);
  if (labelMatch) {
    const label = raw.slice(0, raw.indexOf(':')).trim();
    const expression = raw.slice(raw.indexOf(':') + 1).trim();
    if (expression) return { index, raw, kind: 'labelled_expression', label, expression };
  }

  const assignmentMatch = stripped.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$/);
  if (assignmentMatch) {
    return { index, raw, kind: 'assignment', name: assignmentMatch[1], expression: raw.slice(raw.indexOf('=') + 1).trim() };
  }

  if (looksLikeExpression(raw)) return { index, raw, kind: 'expression', expression: raw.trim() };
  return { index, raw, kind: 'text', text: raw };
}

export function parseDocument(source) {
  const lines = source.split(/\r?\n/).map((raw, index) => classifyLine(raw, index));
  return { source, lines };
}

function tokenize(source) {
  const input = source
    .replace(/\bone hundred\b/gi, '100')
    .replace(/\btwo hundred\b/gi, '200')
    .replace(/\bthree hundred\b/gi, '300')
    .replace(/\bfour hundred\b/gi, '400')
    .replace(/\bfive hundred\b/gi, '500')
    .replace(/\bsix hundred\b/gi, '600')
    .replace(/\bseven hundred\b/gi, '700')
    .replace(/\beight hundred\b/gi, '800')
    .replace(/\bnine hundred\b/gi, '900')
    .replace(/\bone\b/gi, '1')
    .replace(/\btwo\b/gi, '2')
    .replace(/\bthree\b/gi, '3')
    .replace(/\bfour\b/gi, '4')
    .replace(/\bfive\b/gi, '5')
    .replace(/\bsix\b/gi, '6')
    .replace(/\bseven\b/gi, '7')
    .replace(/\beight\b/gi, '8')
    .replace(/\bnine\b/gi, '9')
    .replace(/\bten\b/gi, '10')
    .replace(/\bas\s+a\s+%\s+of\b/gi, ' ASPERCENTOF ')
    .replace(/\bmultiplied\s+by\b/gi, ' * ')
    .replace(/\bdivide\s+by\b/gi, ' / ')
    .replace(/\bdivided\s+by\b/gi, ' / ');

  const tokens = [];
  let i = 0;

  while (i < input.length) {
    const char = input[i];
    if (/\s/.test(char)) {
      i += 1;
      continue;
    }

    if (char === '"') {
      let end = i + 1;
      while (end < input.length && input[end] !== '"') end += 1;
      tokens.push({ type: 'string', value: input.slice(i + 1, end) });
      i = Math.min(end + 1, input.length);
      continue;
    }

    const triSymbol = input.slice(i, i + 2);
    if (['A$', 'C$', 'R$'].includes(triSymbol)) {
      tokens.push({ type: 'currency_symbol', value: triSymbol });
      i += 2;
      continue;
    }

    const dollarNumber = input.slice(i).match(/^\$([0-9]+)\b/);
    if (dollarNumber) {
      tokens.push({ type: 'dollar_number', value: Number(dollarNumber[1]) });
      i += dollarNumber[0].length;
      continue;
    }

    if ('()%,;^+-*/'.includes(char)) {
      tokens.push({ type: 'op', value: char });
      i += 1;
      continue;
    }

    if (char === '$' || char === '€' || char === '£' || char === '¥' || char === '₹' || char === '₩' || char === '°') {
      tokens.push({ type: char === '°' ? 'degree' : 'currency_symbol', value: char });
      i += 1;
      continue;
    }

    if (char === '<' && input[i + 1] === '<') {
      tokens.push({ type: 'op', value: '<<' });
      i += 2;
      continue;
    }
    if (char === '>' && input[i + 1] === '>') {
      tokens.push({ type: 'op', value: '>>' });
      i += 2;
      continue;
    }
    if (char === '&' || char === '|') {
      tokens.push({ type: 'op', value: char });
      i += 1;
      continue;
    }

    const hex = input.slice(i).match(/^0x[0-9a-f]+/i);
    if (hex) {
      tokens.push({ type: 'number', value: Number.parseInt(hex[0], 16) });
      i += hex[0].length;
      continue;
    }

    const bin = input.slice(i).match(/^0b[01]+/i);
    if (bin) {
      tokens.push({ type: 'number', value: Number.parseInt(bin[0].slice(2), 2) });
      i += bin[0].length;
      continue;
    }

    const oct = input.slice(i).match(/^0o[0-7]+/i);
    if (oct) {
      tokens.push({ type: 'number', value: Number.parseInt(oct[0].slice(2), 8) });
      i += oct[0].length;
      continue;
    }

    const num = input.slice(i).match(/^\d+(?:[ _]\d+)*(?:\.\d+)?|\.\d+/);
    if (num) {
      tokens.push({ type: 'number', value: Number(num[0].replace(/[ _]/g, '')) });
      i += num[0].length;
      continue;
    }

    const word = input.slice(i).match(/^[A-Za-z_][A-Za-z0-9_]*/);
    if (word) {
      const lowered = word[0].toLowerCase();
      if (OPERATOR_WORDS.has(lowered)) {
        tokens.push({ type: 'op', value: OPERATOR_WORDS.get(lowered) });
      } else if (['of', 'on', 'off', 'in', 'to', 'ASPERCENTOF'].includes(word[0])) {
        tokens.push({ type: 'keyword', value: word[0].toLowerCase() });
      } else {
        tokens.push({ type: 'identifier', value: word[0] });
      }
      i += word[0].length;
      continue;
    }

    throw new Error(`Unexpected token: ${char}`);
  }

  return tokens;
}

class Parser {
  constructor(tokens) {
    this.tokens = tokens;
    this.index = 0;
  }

  peek(offset = 0) {
    return this.tokens[this.index + offset] || null;
  }

  consume() {
    return this.tokens[this.index++] || null;
  }

  match(type, value = null) {
    const token = this.peek();
    return !!token && token.type === type && (value == null || token.value === value);
  }

  expect(type, value = null) {
    const token = this.consume();
    if (!token || token.type !== type || (value != null && token.value !== value)) {
      throw new Error(`Expected ${value || type}`);
    }
    return token;
  }

  parse() {
    const expr = this.parseConversion();
    if (this.peek()) throw new Error(`Unexpected token: ${this.peek().value}`);
    return expr;
  }

  parseConversion() {
    let node = this.parsePercentCompare();

    while (this.match('keyword', 'in') || this.match('keyword', 'to')) {
      this.consume();
      const phrase = this.readPhrase();
      if (!phrase) throw new Error('Expected conversion target');
      node = { kind: 'convert', expr: node, target: phrase };
    }

    return node;
  }

  parsePercentCompare() {
    let node = this.parseBitwise();
    if (this.match('keyword', 'aspercentof')) {
      this.consume();
      const right = this.parseBitwise();
      return { kind: 'percent_compare', left: node, right };
    }
    return node;
  }

  parseBitwise() {
    let node = this.parseAddSub();
    while (this.match('op', '&') || this.match('op', '|') || this.match('op', 'xor') || this.match('op', '<<') || this.match('op', '>>')) {
      const op = this.consume().value;
      node = { kind: 'binary', op, left: node, right: this.parseAddSub() };
    }
    return node;
  }

  parseAddSub() {
    let node = this.parseMulDiv();
    while (this.match('op', '+') || this.match('op', '-')) {
      const op = this.consume().value;
      node = { kind: 'binary', op, left: node, right: this.parseMulDiv() };
    }
    return node;
  }

  parseMulDiv() {
    let node = this.parsePower();

    while (true) {
      if (this.match('op', '*') || this.match('op', '/') || this.match('op', 'mod')) {
        const op = this.consume().value;
        node = { kind: 'binary', op, left: node, right: this.parsePower() };
        continue;
      }
      if (this.match('keyword', 'of') || this.match('keyword', 'on') || this.match('keyword', 'off')) {
        const op = this.consume().value;
        node = { kind: 'binary', op, left: node, right: this.parsePower() };
        continue;
      }
      if (this.startsPrimary(this.peek())) {
        node = { kind: 'binary', op: '*', left: node, right: this.parsePower() };
        continue;
      }
      break;
    }

    return node;
  }

  parsePower() {
    let node = this.parseUnary();
    if (this.match('op', '^')) {
      this.consume();
      node = { kind: 'binary', op: '^', left: node, right: this.parseUnary() };
    }
    return node;
  }

  parseUnary() {
    if (this.match('op', '+') || this.match('op', '-')) {
      return { kind: 'unary', op: this.consume().value, expr: this.parseUnary() };
    }
    return this.parsePostfix();
  }

  parsePostfix() {
    let node = this.parsePrimary();

    while (true) {
      if (this.match('op', '%')) {
        this.consume();
        node = { kind: 'percent', expr: node };
        continue;
      }

      const square = this.match('identifier') && ['square', 'sq', 'cubic', 'cu', 'cb'].includes(this.peek().value.toLowerCase());
      if (square) {
        const prefix = this.consume().value.toLowerCase();
        const phrase = this.readPhrase();
        if (!phrase) throw new Error('Expected unit after square/cubic');
        node = { kind: 'unit_suffix', expr: node, unit: `${prefix} ${phrase}` };
        continue;
      }

      if (this.match('degree')) {
        this.consume();
        node = { kind: 'unit_suffix', expr: node, unit: 'deg' };
        continue;
      }

      if (this.match('currency_symbol')) {
        const symbol = this.consume().value;
        node = { kind: 'binary', op: '*', left: node, right: { kind: 'currency_symbol', symbol } };
        continue;
      }

      if (this.match('identifier')) {
        const phrase = this.readPhrase({ stopIfKeyword: true });
        const lower = phrase.toLowerCase();
        if (isScale(lower)) {
          node = { kind: 'scale', expr: node, scale: lower };
          continue;
        }
        if (UNIT_REGISTRY.get(lower) || UNIT_REGISTRY.isTemperature(lower) || canExpandDimPrefix(lower) || isCurrencyCode(lower)) {
          node = { kind: 'unit_suffix', expr: node, unit: phrase };
          continue;
        }
        this.index -= phrase.split(/\s+/).length;
      }

      break;
    }

    return node;
  }

  parsePrimary() {
    if (this.match('number')) return { kind: 'number', value: this.consume().value };
    if (this.match('dollar_number')) return { kind: 'dollar_number', value: this.consume().value };
    if (this.match('string')) return { kind: 'string', value: this.consume().value };
    if (this.match('currency_symbol')) return { kind: 'currency_symbol', symbol: this.consume().value };

    if (this.match('identifier')) {
      const identifier = this.consume().value;
      if (identifier.toLowerCase() === 'line' && this.match('number')) {
        return { kind: 'line_ref', index: this.consume().value - 1 };
      }
      if (this.match('op', '(')) {
        this.consume();
        const args = [];
        while (!this.match('op', ')')) {
          args.push(this.parseConversion());
          if (this.match('op', ',') || this.match('op', ';')) this.consume();
          else break;
        }
        this.expect('op', ')');
        return { kind: 'call', name: identifier, args };
      }
      return { kind: 'identifier', name: identifier };
    }

    if (this.match('op', '(')) {
      this.consume();
      const expr = this.parseConversion();
      this.expect('op', ')');
      return expr;
    }

    throw new Error('Unexpected end of expression');
  }

  startsPrimary(token) {
    if (!token) return false;
    if (token.type === 'number' || token.type === 'currency_symbol' || token.type === 'degree' || token.type === 'string') return true;
    if (token.type === 'identifier') return true;
    return token.type === 'op' && token.value === '(';
  }

  readPhrase(options = {}) {
    const words = [];
    while (this.match('identifier')) {
      const word = this.peek().value;
      if (options.stopIfKeyword && !words.length && ['in', 'to'].includes(word.toLowerCase())) break;
      words.push(this.consume().value);
      const candidate = words.join(' ').toLowerCase();
      if (BUILT_IN_TIMEZONES.has(candidate)) break;
      if (UNIT_REGISTRY.get(candidate) || UNIT_REGISTRY.isTemperature(candidate) || isCurrencyCode(candidate)) break;
      if (words.length >= 2) break;
    }
    return words.join(' ').trim();
  }
}

function isScale(word) {
  return ['k', 'thousand', 'm', 'million', 'billion'].includes(word);
}

function scaleFor(word) {
  switch (word) {
    case 'k':
    case 'thousand':
      return 1e3;
    case 'm':
    case 'million':
      return 1e6;
    case 'billion':
      return 1e9;
    default:
      return 1;
  }
}

function isCurrencyCode(word) {
  return /^[a-z]{3}$/i.test(word) || CRYPTO_CODES.has(word.toUpperCase());
}

function canExpandDimPrefix(word) {
  return ['square', 'sq', 'cubic', 'cu', 'cb'].includes(word);
}

function resolvePrefixedUnit(input) {
  const normalized = input.trim().toLowerCase();
  const square = normalized.match(/^(square|sq)\s+(.+)$/);
  if (square) {
    const base = UNIT_REGISTRY.get(square[2]);
    if (!base) return null;
    return {
      id: `sq-${base.id}`,
      symbol: `${base.symbol}²`,
      factor: base.factor ** 2,
      dims: addDims(base.dims, base.dims),
      durationMs: null,
    };
  }
  const cubic = normalized.match(/^(cubic|cu|cb)\s+(.+)$/);
  if (cubic) {
    const base = UNIT_REGISTRY.get(cubic[2]);
    if (!base) return null;
    return {
      id: `cb-${base.id}`,
      symbol: `${base.symbol}³`,
      factor: base.factor ** 3,
      dims: addDims(addDims(base.dims, base.dims), base.dims),
      durationMs: null,
    };
  }
  return null;
}

function resolveUnit(input, settings) {
  const normalized = input.trim().toLowerCase();
  if (normalized === 'ppi') return { id: 'ppi', symbol: 'ppi', factor: 1, dims: zeroDims(), setting: 'ppi' };
  if (normalized === 'em') {
    const base = UNIT_REGISTRY.get('em');
    return { ...base, factor: settings.emPx };
  }
  return resolvePrefixedUnit(normalized) || UNIT_REGISTRY.get(normalized) || UNIT_REGISTRY.isTemperature(normalized) || null;
}

function resolveCurrency(symbolOrCode) {
  const key = symbolOrCode.toUpperCase();
  if (CURRENCY_SYMBOLS[symbolOrCode]) return CURRENCY_SYMBOLS[symbolOrCode];
  return isCurrencyCode(key) ? key : null;
}

function formatNumber(value, format = null, precision = 3) {
  if (!Number.isFinite(value)) return String(value);
  if (format === 'scientific') return value.toExponential(4);
  const digits = Math.max(0, Math.min(8, precision));
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: digits,
  }).format(value);
}

function formatDuration(ms, preferredUnit = 'h', precision = 3) {
  const unit = UNIT_REGISTRY.get(preferredUnit) || UNIT_REGISTRY.get('h');
  const value = ms / unit.durationMs;
  return `${formatNumber(value, null, precision)} ${unit.symbol}`;
}

function formatDateTime(value, locale = 'en-US') {
  return new Intl.DateTimeFormat(locale, {
    dateStyle: 'medium',
    timeStyle: 'short',
    timeZone: value.timezone,
  }).format(new Date(value.ts));
}

export function formatResult(value, options = {}) {
  if (!value) return '';
  const precision = options.precision ?? 3;
  switch (value.kind) {
    case 'number':
      return formatNumber(value.value, value.format || options.format || null, precision);
    case 'percentage':
      return `${formatNumber(value.ratio * 100, null, precision)}%`;
    case 'currency':
      return `${currencySymbol(value.code)} ${formatNumber(value.amount, value.format, precision)}`;
    case 'temperature': {
      const scale = value.unit || 'C';
      const numeric = scale === 'K' ? value.kelvin : scale === 'F' ? ((value.kelvin - 273.15) * 9) / 5 + 32 : value.kelvin - 273.15;
      return `${formatNumber(numeric, null, precision)} ${scale === 'K' ? 'K' : `°${scale}`}`;
    }
    case 'duration':
      return formatDuration(value.ms, value.preferredUnit, precision);
    case 'datetime':
      return formatDateTime(value, options.locale || 'en-US');
    case 'quantity': {
      if (value.preferredUnit) {
        const numeric = value.baseValue / value.preferredUnit.factor;
        return `${formatNumber(numeric, value.format, precision)} ${value.preferredUnit.symbol}`;
      }
      return `${formatNumber(value.baseValue, value.format, precision)} ${dimsToLabel(value.dims)}`;
    }
    case 'text':
      return value.text;
    default:
      return '';
  }
}

function currencySymbol(code) {
  for (const [symbol, current] of Object.entries(CURRENCY_SYMBOLS)) {
    if (current === code) return symbol;
  }
  return code;
}

function coerceValue(node, context) {
  switch (node.kind) {
    case 'number':
      return makeNumber(node.value);
    case 'string':
      return makeText(node.value);
    case 'currency_symbol':
      return makeCurrency(1, resolveCurrency(node.symbol));
    case 'dollar_number': {
      const candidate = context.lineValues[node.value - 1]?.value;
      if (candidate) return cloneValue(candidate);
      return makeCurrency(node.value, 'USD');
    }
    case 'identifier':
      return evaluateIdentifier(node.name, context);
    case 'line_ref': {
      const candidate = context.lineValues[node.index]?.value;
      if (!candidate) throw new Error(`Unknown line $${node.index + 1}`);
      return cloneValue(candidate);
    }
    case 'percent':
      return makePercentage(getNumber(evaluateAst(node.expr, context)) / 100);
    case 'scale': {
      const base = evaluateAst(node.expr, context);
      if (base.kind === 'number') return makeNumber(base.value * scaleFor(node.scale), { format: base.format });
      if (base.kind === 'currency') return makeCurrency(base.amount * scaleFor(node.scale), base.code);
      if (base.kind === 'quantity') return makeQuantity(base.baseValue * scaleFor(node.scale), base.dims, base.preferredUnit, base.format);
      throw new Error('Scale can only apply to numeric values');
    }
    case 'unit_suffix': {
      const base = evaluateAst(node.expr, context);
      return attachUnit(base, node.unit, context);
    }
    case 'unary': {
      const value = evaluateAst(node.expr, context);
      if (node.op === '+') return value;
      if (value.kind === 'number') return makeNumber(-value.value, { format: value.format });
      if (value.kind === 'currency') return makeCurrency(-value.amount, value.code);
      if (value.kind === 'quantity') return makeQuantity(-value.baseValue, value.dims, value.preferredUnit, value.format);
      if (value.kind === 'duration') return makeDuration(-value.ms, value.preferredUnit);
      if (value.kind === 'temperature') return makeTemperature(-value.kelvin, value.unit);
      throw new Error('Unsupported unary value');
    }
    case 'binary':
      return evaluateBinary(node, context);
    case 'call':
      return evaluateCall(node, context);
    case 'convert':
      return convertValue(evaluateAst(node.expr, context), node.target, context);
    case 'percent_compare': {
      const left = evaluateAst(node.left, context);
      const right = evaluateAst(node.right, context);
      return makePercentage(getComparableNumber(left, context) / getComparableNumber(right, context));
    }
    default:
      throw new Error(`Unsupported AST node: ${node.kind}`);
  }
}

function evaluateIdentifier(name, context) {
  const lowered = name.toLowerCase();
  if (context.variables.has(name)) return cloneValue(context.variables.get(name));
  if (context.variables.has(lowered)) return cloneValue(context.variables.get(lowered));
  if (lowered === 'prev') {
    const prev = getPrevValue(context);
    if (!prev) throw new Error('No previous result');
    return cloneValue(prev);
  }
  if (lowered === 'sum' || lowered === 'total') return computeRollup(context, 'sum');
  if (lowered === 'average' || lowered === 'avg') return computeRollup(context, 'average');
  if (lowered === 'pi') return makeNumber(Math.PI);
  if (lowered === 'e') return makeNumber(Math.E);
  if (lowered === 'now' || lowered === 'time' || lowered === 'today') {
    const parsed = context.dateTimeProvider.parseExpression(lowered, context.locale, context.timezone);
    if (parsed instanceof Error) throw parsed;
    return parsed;
  }
  if (/^\$[0-9]+$/.test(name)) {
    const index = Number(name.slice(1)) - 1;
    const line = context.lineValues[index];
    if (!line || !line.value) throw new Error(`Unknown line ${name}`);
    return cloneValue(line.value);
  }
  throw new Error(`Unknown identifier: ${name}`);
}

function getPrevValue(context) {
  for (let i = context.lineValues.length - 1; i >= 0; i -= 1) {
    const value = context.lineValues[i]?.value;
    if (value) return value;
  }
  return null;
}

function computeRollup(context, mode) {
  const items = [];
  for (let i = context.lineValues.length - 1; i >= 0; i -= 1) {
    const line = context.lineValues[i];
    if (line?.breaksRollup) break;
    if (!line?.value || !isNumericLike(line.value)) break;
    items.unshift(line.value);
  }
  if (!items.length) throw new Error(`No lines available for ${mode}`);
  let total = cloneValue(items[0]);
  for (let i = 1; i < items.length; i += 1) total = addValues(total, items[i], context);
  if (mode === 'sum') return total;
  return divideValues(total, makeNumber(items.length), context);
}

function attachUnit(value, unitText, context) {
  const unit = resolveUnit(unitText, context.settings);
  if (!unit) {
    const currency = resolveCurrency(unitText);
    if (currency) return multiplyValues(value, makeCurrency(1, currency), context);
    throw new Error(`Unknown unit: ${unitText}`);
  }

  if (unit.toKelvin) {
    const numeric = getNumber(value);
    return makeTemperature(unit.toKelvin(numeric), unit.id);
  }

  if (unit.durationMs) {
    return makeDuration(getNumber(value) * unit.durationMs, unit.id);
  }

  if (unit.setting) {
    return makeNumber(getNumber(value));
  }

  return makeQuantity(getNumber(value) * unit.factor, unit.dims, unit);
}

function getNumber(value) {
  if (value.kind === 'number') return value.value;
  if (value.kind === 'percentage') return value.ratio;
  throw new Error('Expected number');
}

function getComparableNumber(value, context) {
  if (value.kind === 'number') return value.value;
  if (value.kind === 'currency') {
    return convertCurrency(value.amount, value.code, context.fx.base, context.fx);
  }
  if (value.kind === 'quantity') return value.baseValue;
  throw new Error('Unsupported comparison');
}

function evaluateBinary(node, context) {
  const left = evaluateAst(node.left, context);
  const right = evaluateAst(node.right, context);
  switch (node.op) {
    case '+':
      return addValues(left, right, context);
    case '-':
      return subtractValues(left, right, context);
    case '*':
      return multiplyValues(left, right, context);
    case '/':
      return divideValues(left, right, context);
    case '^':
      return powerValues(left, right);
    case 'mod':
      return makeNumber(getComparableNumber(left, context) % getComparableNumber(right, context));
    case 'of':
      return percentOf(left, right, context);
    case 'on':
      return percentOn(left, right, context);
    case 'off':
      return percentOff(left, right, context);
    case '&':
      return makeNumber(getNumber(asPlainNumber(left)) & getNumber(asPlainNumber(right)));
    case '|':
      return makeNumber(getNumber(asPlainNumber(left)) | getNumber(asPlainNumber(right)));
    case 'xor':
      return makeNumber(getNumber(asPlainNumber(left)) ^ getNumber(asPlainNumber(right)));
    case '<<':
      return makeNumber(getNumber(asPlainNumber(left)) << getNumber(asPlainNumber(right)));
    case '>>':
      return makeNumber(getNumber(asPlainNumber(left)) >> getNumber(asPlainNumber(right)));
    default:
      throw new Error(`Unsupported operator: ${node.op}`);
  }
}

function asPlainNumber(value) {
  if (value.kind === 'number') return value;
  throw new Error('Bitwise operators require plain numbers');
}

function addValues(left, right, context) {
  if (left.kind === 'percentage' && right.kind !== 'percentage') return percentOn(left, right, context);
  if (right.kind === 'percentage' && left.kind !== 'percentage') return percentOn(right, left, context);
  if (left.kind === 'number' && right.kind === 'number') return makeNumber(left.value + right.value);
  if (left.kind === 'currency' && right.kind === 'currency') {
    const amount = left.amount + convertCurrency(right.amount, right.code, left.code, context.fx);
    return makeCurrency(amount, left.code);
  }
  if (left.kind === 'currency' && right.kind === 'number') return makeCurrency(left.amount + right.value, left.code);
  if (left.kind === 'number' && right.kind === 'currency') return makeCurrency(left.value + right.amount, right.code);
  if (left.kind === 'quantity' && right.kind === 'quantity') {
    if (!dimsEqual(left.dims, right.dims)) throw new Error('Unit mismatch');
    return makeQuantity(left.baseValue + right.baseValue, left.dims, left.preferredUnit || right.preferredUnit);
  }
  if (left.kind === 'duration' && right.kind === 'duration') return makeDuration(left.ms + right.ms, left.preferredUnit || right.preferredUnit);
  if (left.kind === 'datetime' && right.kind === 'duration') return makeDateTime(left.ts + right.ms, left.timezone);
  if (left.kind === 'duration' && right.kind === 'datetime') return makeDateTime(right.ts + left.ms, right.timezone);
  throw new Error('Unsupported addition');
}

function subtractValues(left, right, context) {
  if (right.kind === 'percentage' && left.kind !== 'percentage') return percentOff(right, left, context);
  if (left.kind === 'number' && right.kind === 'number') return makeNumber(left.value - right.value);
  if (left.kind === 'currency' && right.kind === 'currency') {
    const amount = left.amount - convertCurrency(right.amount, right.code, left.code, context.fx);
    return makeCurrency(amount, left.code);
  }
  if (left.kind === 'quantity' && right.kind === 'quantity') {
    if (!dimsEqual(left.dims, right.dims)) throw new Error('Unit mismatch');
    return makeQuantity(left.baseValue - right.baseValue, left.dims, left.preferredUnit || right.preferredUnit);
  }
  if (left.kind === 'duration' && right.kind === 'duration') return makeDuration(left.ms - right.ms, left.preferredUnit || right.preferredUnit);
  if (left.kind === 'datetime' && right.kind === 'duration') return makeDateTime(left.ts - right.ms, left.timezone);
  if (left.kind === 'datetime' && right.kind === 'datetime') return makeDuration(left.ts - right.ts);
  throw new Error('Unsupported subtraction');
}

function multiplyValues(left, right) {
  if (left.kind === 'number' && right.kind === 'number') return makeNumber(left.value * right.value);
  if (left.kind === 'currency' && right.kind === 'number') return makeCurrency(left.amount * right.value, left.code);
  if (left.kind === 'number' && right.kind === 'currency') return makeCurrency(left.value * right.amount, right.code);
  if (left.kind === 'quantity' && right.kind === 'number') return makeQuantity(left.baseValue * right.value, left.dims, left.preferredUnit, left.format);
  if (left.kind === 'number' && right.kind === 'quantity') return makeQuantity(left.value * right.baseValue, right.dims, right.preferredUnit, right.format);
  if (left.kind === 'duration' && right.kind === 'number') return makeDuration(left.ms * right.value, left.preferredUnit);
  if (left.kind === 'number' && right.kind === 'duration') return makeDuration(left.value * right.ms, right.preferredUnit);
  if (left.kind === 'quantity' && right.kind === 'quantity') {
    return makeQuantity(left.baseValue * right.baseValue, addDims(left.dims, right.dims));
  }
  throw new Error('Unsupported multiplication');
}

function divideValues(left, right) {
  if (right.kind !== 'number') throw new Error('Division expects a plain number on the right side');
  if (left.kind === 'number') return makeNumber(left.value / right.value);
  if (left.kind === 'currency') return makeCurrency(left.amount / right.value, left.code);
  if (left.kind === 'quantity') return makeQuantity(left.baseValue / right.value, left.dims, left.preferredUnit, left.format);
  if (left.kind === 'duration') return makeDuration(left.ms / right.value, left.preferredUnit);
  throw new Error('Unsupported division');
}

function powerValues(left, right) {
  if (left.kind === 'number' && right.kind === 'number') return makeNumber(left.value ** right.value);
  throw new Error('Exponentiation currently supports plain numbers only');
}

function percentOf(left, right, context) {
  if (left.kind !== 'percentage') return multiplyValues(left, right, context);
  return multiplyValues(makeNumber(left.ratio), right, context);
}

function percentOn(left, right, context) {
  if (left.kind !== 'percentage') throw new Error('Expected percentage before "on"');
  return addValues(right, multiplyValues(makeNumber(left.ratio), right, context), context);
}

function percentOff(left, right, context) {
  if (left.kind !== 'percentage') throw new Error('Expected percentage before "off"');
  return subtractValues(right, multiplyValues(makeNumber(left.ratio), right, context), context);
}

function convertCurrency(amount, from, to, fx) {
  const base = fx.base || 'USD';
  const rates = fx.rates || {};
  if (from === to) return amount;
  const inBase = from === base ? amount : amount / (rates[from] || 1);
  return to === base ? inBase : inBase * (rates[to] || 1);
}

function suggestConversionTarget(target, value, context) {
  const trimmed = target.trim();
  if (!trimmed) return null;

  const timezoneCandidate = trimmed.replace(/\s+(is|please|pls|now)$/i, '').trim();
  if (timezoneCandidate && timezoneCandidate !== trimmed && context.dateTimeProvider.resolveTimezone(timezoneCandidate)) {
    if (value.kind === 'datetime') return `time in ${timezoneCandidate}`;
    if (value.kind === 'text' && value.text.toLowerCase() === 'time') return `time in ${timezoneCandidate}`;
  }

  return null;
}

function convertValue(value, target, context) {
  const normalized = target.trim();
  const lower = normalized.toLowerCase();

  if (lower === 'sci' || lower === 'scientific') {
    if (value.kind === 'number') return { ...value, format: 'scientific' };
    if (value.kind === 'currency') return { ...value, format: 'scientific' };
    if (value.kind === 'quantity') return { ...value, format: 'scientific' };
    return value;
  }

  const currency = resolveCurrency(normalized);
  if (currency) {
    if (value.kind === 'currency') return makeCurrency(convertCurrency(value.amount, value.code, currency, context.fx), currency);
    if (value.kind === 'number') return makeCurrency(convertCurrency(value.value, context.fx.base, currency, context.fx), currency);
    throw new Error('Currency conversion requires a number or currency value');
  }

  const unit = resolveUnit(normalized, context.settings);
  if (unit?.setting) {
    context.settings[unit.setting] = getNumber(value);
    return makeNumber(context.settings[unit.setting]);
  }

  if (unit?.toKelvin) {
    if (value.kind !== 'temperature') throw new Error('Temperature conversion requires a temperature');
    return makeTemperature(value.kelvin, unit.id);
  }

  if (unit?.durationMs) {
    if (value.kind !== 'duration') throw new Error('Duration conversion requires a duration');
    return makeDuration(value.ms, unit.id);
  }

  if (unit) {
    if (value.kind !== 'quantity') throw new Error('Unit conversion requires a quantity');
    if (!dimsEqual(value.dims, unit.dims)) {
      if (canConvertCssLength(value, unit, context)) return convertCssLength(value, unit, context);
      throw new Error('Incompatible conversion');
    }
    return makeQuantity(value.baseValue, value.dims, unit, value.format);
  }

  const timezone = context.dateTimeProvider.resolveTimezone(normalized);
  if (timezone) {
    if (value.kind === 'datetime') return makeDateTime(value.ts, timezone);
    if (value.kind === 'number' && (lower === 'now' || lower === 'time')) return makeDateTime(context.dateTimeProvider.now(), timezone);
    const parsed = context.dateTimeProvider.parseExpression('now', context.locale, timezone);
    if (value.kind === 'text' && value.text.toLowerCase() === 'time') return parsed;
    throw new Error('Timezone conversion requires a datetime');
  }

  const suggestion = suggestConversionTarget(target, value, context);
  throw new Error(suggestion ? `Unknown conversion target: ${target}. Try: ${suggestion}` : `Unknown conversion target: ${target}`);
}

function evaluateCall(node, context) {
  const name = node.name.toLowerCase();
  const args = node.args.map((arg) => evaluateAst(arg, context));
  const first = args[0];

  switch (name) {
    case 'sqrt':
      return makeNumber(Math.sqrt(getNumber(first)));
    case 'cbrt':
      return makeNumber(Math.cbrt(getNumber(first)));
    case 'abs':
      return makeNumber(Math.abs(getNumber(first)));
    case 'round':
      return makeNumber(Math.round(getNumber(first)));
    case 'ceil':
      return makeNumber(Math.ceil(getNumber(first)));
    case 'floor':
      return makeNumber(Math.floor(getNumber(first)));
    case 'ln':
      return makeNumber(Math.log(getNumber(first)));
    case 'log':
      return args.length > 1 ? makeNumber(Math.log(getNumber(args[1])) / Math.log(getNumber(first))) : makeNumber(Math.log10(getNumber(first)));
    case 'sin':
      return makeNumber(Math.sin(asRadians(first)));
    case 'cos':
      return makeNumber(Math.cos(asRadians(first)));
    case 'tan':
      return makeNumber(Math.tan(asRadians(first)));
    case 'arcsin':
      return makeNumber(Math.asin(getNumber(first)));
    case 'arccos':
      return makeNumber(Math.acos(getNumber(first)));
    case 'arctan':
      return makeNumber(Math.atan(getNumber(first)));
    case 'fact':
      return makeNumber(factorial(getNumber(first)));
    case 'root':
      return makeNumber(Math.pow(getNumber(args[1]), 1 / getNumber(first)));
    case 'fromunix':
      return context.dateTimeProvider.fromUnix(getNumber(first), context.timezone);
    default:
      throw new Error(`Unknown function: ${node.name}`);
  }
}

function canConvertCssLength(value, unit) {
  const cssDims = normalizeDims({ css: 1 });
  const lengthDims = normalizeDims({ length: 1 });
  return (
    value.kind === 'quantity' &&
    ((dimsEqual(value.dims, lengthDims) && dimsEqual(unit.dims, cssDims)) ||
      (dimsEqual(value.dims, cssDims) && dimsEqual(unit.dims, lengthDims)))
  );
}

function convertCssLength(value, unit, context) {
  const cssDims = normalizeDims({ css: 1 });
  const lengthDims = normalizeDims({ length: 1 });
  const pxPerMeter = (context.settings.ppi || 96) / 0.0254;

  if (dimsEqual(value.dims, lengthDims) && dimsEqual(unit.dims, cssDims)) {
    const px = value.baseValue * pxPerMeter;
    return makeQuantity(px, cssDims, unit);
  }

  if (dimsEqual(value.dims, cssDims) && dimsEqual(unit.dims, lengthDims)) {
    const meters = value.baseValue / pxPerMeter;
    return makeQuantity(meters, lengthDims, unit);
  }

  throw new Error('Incompatible conversion');
}

function asRadians(value) {
  if (value.kind === 'quantity' && value.dims.angle === 1) return value.baseValue;
  if (value.kind === 'number') return value.value;
  throw new Error('Trigonometric functions require an angle');
}

function factorial(value) {
  if (value < 0 || !Number.isInteger(value)) throw new Error('Factorial requires a non-negative integer');
  let out = 1;
  for (let i = 2; i <= value; i += 1) out *= i;
  return out;
}

export function evaluateExpression(expression, context) {
  const cleaned = expression.trim().replace(/\binto\b/gi, 'to');

  const timePrefix = cleaned.match(/^(time|now|today)\s+in\s+(.+)$/i);
  if (timePrefix) {
    return convertValue(context.dateTimeProvider.parseExpression(timePrefix[1].toLowerCase(), context.locale, context.timezone), timePrefix[2], context);
  }

  const locationTime = cleaned.match(/^(.+?)\s+time(?:\s+is)?$/i);
  if (locationTime) {
    return convertValue(context.dateTimeProvider.parseExpression('time', context.locale, context.timezone), locationTime[1], context);
  }

  const tokens = tokenize(cleaned);
  const ast = new Parser(tokens).parse();
  return evaluateAst(ast, context);
}

function evaluateAst(ast, context) {
  return coerceValue(ast, context);
}

export function renderGhostHtml(documentModel, evaluated) {
  const knownVariables = new Set([...evaluated.variables.keys()].map((key) => key.toLowerCase()));
  return documentModel.lines.map((line) => renderLineGhost(line, knownVariables)).join('\n');
}

function renderLineGhost(line, variables) {
  if (line.kind === 'header') return `<span class="ghost-header">${escapeHtml(line.raw)}</span>`;
  if (line.kind === 'comment') return `<span class="ghost-comment">${escapeHtml(line.raw)}</span>`;
  if (line.kind === 'labelled_expression') {
    const raw = escapeHtml(line.raw);
    const colon = raw.indexOf(':');
    return `<span class="ghost-label">${raw.slice(0, colon + 1)}</span>${highlightExpression(raw.slice(colon + 1), variables)}`;
  }
  if (line.kind === 'assignment') {
    const raw = escapeHtml(line.raw);
    const eq = raw.indexOf('=');
    const left = raw.slice(0, eq);
    const [, leading = '', core = '', trailing = ''] = left.match(/^(\s*)(.*?)(\s*)$/) || [];
    return `${leading}<span class="ghost-var">${core}</span>${trailing}${raw.slice(eq, eq + 1)}${highlightExpression(raw.slice(eq + 1), variables)}`;
  }
  if (line.kind === 'expression') return highlightExpression(escapeHtml(line.raw), variables);
  return escapeHtml(line.raw || '');
}

function highlightExpression(raw, variables) {
  return raw
    .replace(/(&quot;[^&]*&quot;|"[^"]*")/g, '<span class="ghost-string">$1</span>')
    .replace(/\b([A-Za-z_][A-Za-z0-9_]*)\b/g, (match, name) => {
      const lower = name.toLowerCase();
      if (variables.has(lower)) return `<span class="ghost-id">${match}</span>`;
      if (isCurrencyCode(lower)) return `<span class="ghost-fx">${match}</span>`;
      if (UNIT_REGISTRY.get(lower) || UNIT_REGISTRY.isTemperature(lower)) return `<span class="ghost-unit">${match}</span>`;
      return match;
    });
}

export function evaluateDocument(documentModel, baseContext) {
  const context = {
    ...baseContext,
    variables: new Map(baseContext.variables || []),
    lineValues: [],
  };

  const results = [];

  for (const line of documentModel.lines) {
    const lineContext = { ...context, lineValues: context.lineValues };
    const blankBreak = line.kind === 'blank';

    if (['blank', 'header', 'comment', 'text'].includes(line.kind)) {
      const record = { line, value: null, displayText: '', error: null, kind: line.kind };
      results.push(record);
      context.lineValues.push({ value: null, breaksRollup: blankBreak });
      continue;
    }

    try {
      const expr = line.expression;
      const value = evaluateExpression(expr, lineContext);

      if (line.kind === 'assignment') {
        context.variables.set(line.name, cloneValue(value));
        context.variables.set(line.name.toLowerCase(), cloneValue(value));
      }

      if (line.kind === 'assignment' && (line.name === 'ppi' || line.name === 'em')) {
        context.settings[line.name === 'ppi' ? 'ppi' : 'emPx'] = getNumber(value);
      }

      const displayText = formatResult(value, { locale: context.locale, precision: context.settings.precision });
      const record = { line, value, displayText, error: null, kind: line.kind };
      results.push(record);
      context.lineValues.push({ value, breaksRollup: blankBreak });
    } catch (error) {
      const record = { line, value: null, displayText: '', error: error.message || String(error), kind: line.kind };
      results.push(record);
      context.lineValues.push({ value: null, breaksRollup: blankBreak });
    }
  }

  return {
    lines: results,
    variables: context.variables,
    fx: context.fx,
  };
}
