"use client";

import React, { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react';

/**
 * Light-weight i18n: React Context + dictionaries inline. We keep it simple
 * (no external lib) because the app's translatable surface is small and we
 * don't need plural rules / date formatting / lazy-loaded namespaces.
 *
 * Adding a new locale: append a column to MESSAGES below and (if needed)
 * extend the Locale union. Strings missing in non-default locales fall
 * back to the default (Spanish).
 */

export type Locale = 'es' | 'en';
export const DEFAULT_LOCALE: Locale = 'es';
export const LOCALES: { code: Locale; label: string; flag: string }[] = [
  { code: 'es', label: 'Español', flag: '🇨🇴' },
  { code: 'en', label: 'English', flag: '🇺🇸' },
];

const STORAGE_KEY = 'socya:locale';

// Each entry: { es: '…', en: '…' }. Use {var} for interpolation.
// Keys are the literal default strings to keep migration easy.
type MessageMap = Record<string, Record<Locale, string>>;
const MESSAGES: MessageMap = {
  // ── Brand / chrome ──
  'Sube tu Excel y obtén un PPT en segundos': {
    es: 'Sube tu Excel y obtén un PPT en segundos',
    en: 'Upload your Excel and get a PPT in seconds',
  },
  'La IA lee, decide los slides y descargas listo.': {
    es: 'La IA lee, decide los slides y descargas listo.',
    en: 'AI reads it, picks the slides and you download — done.',
  },
  'Detección inteligente activa': {
    es: 'Detección inteligente activa',
    en: 'Smart detection on',
  },
  'Avanzado': { es: 'Avanzado', en: 'Advanced' },
  'Cambiar': { es: 'Cambiar', en: 'Change' },
  'Estado del sistema': { es: 'Estado del sistema', en: 'System status' },
  'Diagnóstico del backend': { es: 'Diagnóstico del backend', en: 'Backend diagnostics' },
  'Estado del sistema ↗': { es: 'Estado del sistema ↗', en: 'System status ↗' },

  // ── Dropzone ──
  'Sube el Excel con el que quieres trabajar': {
    es: 'Sube el Excel con el que quieres trabajar',
    en: 'Upload the Excel you want to work with',
  },
  'Suelta tu archivo aquí': {
    es: 'Suelta tu archivo aquí',
    en: 'Drop your file here',
  },
  'Arrastra o': { es: 'Arrastra o', en: 'Drag or' },
  'selecciónalo': { es: 'selecciónalo', en: 'select it' },
  '¿Excel desordenado? Organízalo primero (opcional)': {
    es: '¿Excel desordenado? Organízalo primero (opcional)',
    en: 'Messy Excel? Organize it first (optional)',
  },
  'Organizar este Excel y descargarlo (opcional)': {
    es: 'Organizar este Excel y descargarlo (opcional)',
    en: 'Organize this Excel and download it (optional)',
  },

  // ── Buttons / generic ──
  'Generar otra': { es: 'Generar otra', en: 'Generate another' },
  'Ver detalles': { es: 'Ver detalles', en: 'View details' },
  'Reintentar': { es: 'Reintentar', en: 'Retry' },
  'Cancelar': { es: 'Cancelar', en: 'Cancel' },
};

interface I18nContextValue {
  locale: Locale;
  setLocale: (l: Locale) => void;
  t: (key: string, vars?: Record<string, string | number>) => string;
}

const I18nContext = createContext<I18nContextValue>({
  locale: DEFAULT_LOCALE,
  setLocale: () => undefined,
  t: (k) => k,
});

export function I18nProvider({ children }: { children: React.ReactNode }) {
  const [locale, setLocaleState] = useState<Locale>(DEFAULT_LOCALE);

  // Hydrate from localStorage after mount (SSR-safe)
  useEffect(() => {
    try {
      const saved = window.localStorage.getItem(STORAGE_KEY) as Locale | null;
      if (saved && LOCALES.some((l) => l.code === saved)) {
        setLocaleState(saved);
      }
    } catch { /* ignore — private mode etc. */ }
  }, []);

  const setLocale = useCallback((l: Locale) => {
    setLocaleState(l);
    try {
      window.localStorage.setItem(STORAGE_KEY, l);
    } catch { /* ignore */ }
  }, []);

  const t = useCallback(
    (key: string, vars?: Record<string, string | number>) => {
      const entry = MESSAGES[key];
      const raw = (entry && entry[locale]) || (entry && entry[DEFAULT_LOCALE]) || key;
      if (!vars) return raw;
      return raw.replace(/\{(\w+)\}/g, (_m, name) =>
        vars[name] !== undefined ? String(vars[name]) : `{${name}}`,
      );
    },
    [locale],
  );

  const value = useMemo(() => ({ locale, setLocale, t }), [locale, setLocale, t]);
  return <I18nContext.Provider value={value}>{children}</I18nContext.Provider>;
}

export function useI18n(): I18nContextValue {
  return useContext(I18nContext);
}

/** Convenience hook to get just the translation function. */
export function useT() {
  return useContext(I18nContext).t;
}
