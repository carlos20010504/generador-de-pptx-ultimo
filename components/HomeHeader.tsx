"use client";

import React from 'react';
import { Sparkles } from 'lucide-react';
import LocaleSwitcher from './LocaleSwitcher';
import { useT } from '@/utils/i18n';

/**
 * Header del home. Componente cliente para poder usar `useT()` y mostrar
 * el LocaleSwitcher. La página sigue siendo Server Component (más rápido
 * en SSR) y delega solo este bloque al cliente.
 */
export default function HomeHeader() {
  const t = useT();
  return (
    <header className="page-header animate-fade-in-up">
      <div className="brand">
        <div className="brand-mark" aria-hidden>
          {/* Stylized "S" leaf — minimal Socya glyph */}
          <svg viewBox="0 0 32 32" width="22" height="22" fill="none">
            <path
              d="M22.5 9.4c-1.6-1.7-4-2.6-6.6-2.6-3.5 0-6.4 2.1-6.4 4.7 0 2.4 2 3.7 5.7 4.5l1.6.4c4.5 1 7.1 2.7 7.1 6.1 0 3.6-3.6 6.2-8.4 6.2-3.4 0-6.3-1.4-7.9-3.5"
              stroke="#69BE28" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round"
            />
            <circle cx="24" cy="6.5" r="1.8" fill="#69BE28" />
          </svg>
        </div>
        <div className="brand-text">
          <p className="brand-title">Socya <span className="brand-title-accent">PPTX</span></p>
          <p className="brand-sub">Generador inteligente · v4.0</p>
        </div>
      </div>

      <div className="page-header-right">
        <LocaleSwitcher />
        <div className="status-pill" role="status">
          <Sparkles size={11} />
          <span>{t('Detección inteligente activa')}</span>
        </div>
      </div>

      <style>{`
        .page-header-right {
          display: inline-flex;
          align-items: center;
          gap: 0.6rem;
          flex-wrap: wrap;
        }
      `}</style>
    </header>
  );
}
