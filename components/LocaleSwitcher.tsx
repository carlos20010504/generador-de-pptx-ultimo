"use client";

import React from 'react';
import { Globe } from 'lucide-react';
import { LOCALES, useI18n } from '@/utils/i18n';

/**
 * Switch entre locales disponibles. Diseño minimal — un par de píldoras
 * pequeñas en el header. Persiste en localStorage vía I18nProvider.
 */
export default function LocaleSwitcher() {
  const { locale, setLocale } = useI18n();
  return (
    <div className="loc-switcher" role="group" aria-label="Idioma">
      <Globe size={11} className="loc-switcher-icon" aria-hidden />
      {LOCALES.map((l) => (
        <button
          key={l.code}
          type="button"
          onClick={() => setLocale(l.code)}
          className={`loc-switcher-btn ${locale === l.code ? 'is-active' : ''}`}
          aria-pressed={locale === l.code}
          title={l.label}
        >
          {l.code.toUpperCase()}
        </button>
      ))}
      <style>{`
        .loc-switcher {
          display: inline-flex; align-items: center;
          gap: 0.2rem;
          padding: 0.18rem 0.35rem;
          border: 1px solid var(--c-border-strong);
          background: var(--c-bg-elevated);
          border-radius: var(--r-pill);
          font-family: var(--font-heading);
        }
        .loc-switcher-icon {
          color: var(--c-text-muted);
          margin-right: 0.15rem;
        }
        .loc-switcher-btn {
          background: transparent;
          border: none;
          padding: 0.18rem 0.45rem;
          border-radius: var(--r-pill);
          color: var(--c-text-tertiary);
          font-size: 0.62rem;
          font-weight: 800;
          letter-spacing: 0.06em;
          cursor: pointer;
          transition: all 0.18s;
        }
        .loc-switcher-btn:hover { color: var(--c-primary); }
        .loc-switcher-btn.is-active {
          background: var(--c-primary);
          color: white;
          box-shadow: 0 1px 3px rgba(8, 112, 98, 0.30);
        }
      `}</style>
    </div>
  );
}
