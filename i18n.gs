/**
 * Internationalization (i18n) helpers.
 *
 * Each language lives in its own file (Lang_en.gs, Lang_ko.gs, ...) and registers itself:
 *     var LANGS = LANGS || {};
 *     LANGS.xx = { _name: 'Language name', key: 'text', ... };
 *
 * To add a language, copy Lang_en.gs to Lang_<code>.gs, change the code and translate the values.
 * Missing keys automatically fall back to English.
 *
 * Language resolution order:
 *   1. The language the user picked in the sidebar (saved per user)
 *   2. The user's Google account language (Session.getActiveUserLocale)
 *   3. English
 */

var LANGS = LANGS || {};

const I18N = {
  DEFAULT_LANG: 'en',
  PROP_LANG: 'mm_lang',
};

let currentLang_ = null; // cached for the current execution

/** Returns the language code to use for the current user. */
function getLang_() {
  if (currentLang_) return currentLang_;
  let lang = null;
  try {
    lang = PropertiesService.getUserProperties().getProperty(I18N.PROP_LANG);
  } catch (e) {
    // Properties may be unavailable in some simple-trigger contexts.
  }
  if (!lang || !LANGS[lang]) lang = localeToLang_(safeLocale_());
  currentLang_ = lang;
  return lang;
}

/** Maps a locale such as "ko", "ko_KR" or "en-US" to a supported language code. */
function localeToLang_(locale) {
  const code = String(locale || '').toLowerCase().split(/[-_]/)[0];
  return LANGS[code] ? code : I18N.DEFAULT_LANG;
}

function safeLocale_() {
  try {
    return Session.getActiveUserLocale();
  } catch (e) {
    return '';
  }
}

/** Called from the sidebar's language selector. */
function setLanguage(code) {
  if (!LANGS[code]) throw new Error('Unsupported language: ' + code);
  PropertiesService.getUserProperties().setProperty(I18N.PROP_LANG, code);
  currentLang_ = code;
  return code;
}

/** List of available languages for the selector, e.g. [{code:'en', name:'English'}, ...] */
function getLanguages_() {
  return Object.keys(LANGS)
    .sort()
    .map(code => ({ code: code, name: LANGS[code]._name || code }));
}

/**
 * Translate a key. Placeholders like {name} are replaced from params.
 * Unknown placeholders (and merge fields such as {{Name}}) are left untouched.
 */
function t(key, params, lang) {
  const dict = LANGS[lang || getLang_()] || {};
  let text = dict[key];
  if (text === undefined) text = (LANGS[I18N.DEFAULT_LANG] || {})[key];
  if (text === undefined) return key;
  return formatText_(text, params);
}

function formatText_(text, params) {
  if (!params) return text;
  return text.replace(/\{(\w+)\}/g, (m, k) => (Object.prototype.hasOwnProperty.call(params, k) ? String(params[k]) : m));
}

/** All strings for a language, with English filling any gaps (sent to the sidebar). */
function getStrings_(lang) {
  const base = LANGS[I18N.DEFAULT_LANG] || {};
  const dict = LANGS[lang] || {};
  const out = {};
  Object.keys(base).forEach(k => { out[k] = base[k]; });
  Object.keys(dict).forEach(k => { out[k] = dict[k]; });
  return out;
}
