const STORAGE_KEY = 'secom_preferences_v2';

const DEFAULTS = {
  theme: 'dark',
  company: {
    advisorName: 'Jorge Alejandro Díaz Gaxiola',
    companyName: 'SECOM Energía Solar',
    companyPhone: '',
    companyEmail: '',
    companyWebsite: 'https://cashvolt.mx/public/login',
  },
  quoteDefaults: {
    yieldKwhPerKwpMonth: 135,
    panelWatts: 550,
    costPerKwp: 22000,
    contingencyPct: 0.06,
    taxPct: 0.16,
  },
  ocr: {
    aggressiveMode: true,
    preferHighContrast: true,
  }
};

function mergeDeep(base, patch) {
  if (!patch || typeof patch !== 'object') return structuredClone(base);
  const out = Array.isArray(base) ? [...base] : { ...base };
  Object.keys(patch).forEach((key) => {
    const current = patch[key];
    if (current && typeof current === 'object' && !Array.isArray(current) && typeof out[key] === 'object' && out[key] !== null) {
      out[key] = mergeDeep(out[key], current);
    } else {
      out[key] = current;
    }
  });
  return out;
}

export function getDefaultPreferences() {
  return structuredClone(DEFAULTS);
}

export function loadUserPreferences() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return getDefaultPreferences();
    const parsed = JSON.parse(raw);
    return mergeDeep(DEFAULTS, parsed);
  } catch {
    return getDefaultPreferences();
  }
}

export function saveUserPreferences(prefs) {
  const merged = mergeDeep(DEFAULTS, prefs || {});
  localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
  return merged;
}

export function getDefaultQuoteParams() {
  return structuredClone(loadUserPreferences().quoteDefaults);
}

export function applyTheme(theme) {
  document.documentElement.dataset.theme = theme || 'dark';
}
