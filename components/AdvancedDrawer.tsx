"use client";

import React, { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { X } from 'lucide-react';

interface Props {
  open: boolean;
  onClose: () => void;
  children: React.ReactNode;
  title?: string;
}

/**
 * Right-side drawer that hosts the AIControlPanel without modifying its
 * internals. Renders through a portal directly on document.body so its
 * `position: fixed` is not captured by transformed ancestors (animate-*
 * classes create a containing block which would mis-position the drawer).
 */
export default function AdvancedDrawer({ open, onClose, children, title = 'Panel avanzado' }: Props) {
  const [mounted, setMounted] = useState(false);
  useEffect(() => { setMounted(true); }, []);

  // Close on ESC.
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  // Lock body scroll while open.
  useEffect(() => {
    if (!open) return;
    const prev = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => { document.body.style.overflow = prev; };
  }, [open]);

  if (!mounted) return null;

  return createPortal(
    <>
      <div
        className={`adv-backdrop ${open ? 'is-open' : ''}`}
        onClick={onClose}
        aria-hidden={!open}
      />
      <aside
        className={`adv-drawer ${open ? 'is-open' : ''}`}
        role="dialog"
        aria-modal="true"
        aria-label={title}
      >
        <header className="adv-header">
          <h3 className="adv-title">{title}</h3>
          <button
            type="button"
            onClick={onClose}
            className="adv-close press-on-active"
            aria-label="Cerrar panel avanzado"
          >
            <X size={16} />
          </button>
        </header>
        <div className="adv-body">
          {children}
        </div>
      </aside>

      <style>{`
        .adv-backdrop {
          position: fixed; inset: 0;
          background: rgba(18, 60, 73, 0.42);
          backdrop-filter: blur(4px);
          opacity: 0; pointer-events: none;
          transition: opacity 220ms var(--ease-out);
          z-index: 90;
        }
        .adv-backdrop.is-open { opacity: 1; pointer-events: auto; }

        .adv-drawer {
          position: fixed;
          top: 0; right: 0; bottom: 0;
          width: 100%;
          max-width: 380px;
          background: var(--c-bg-elevated);
          border-left: 1px solid var(--c-border);
          box-shadow: -16px 0 48px rgba(18, 60, 73, 0.18);
          transform: translateX(100%);
          transition: transform 280ms var(--ease-out);
          z-index: 100;
          display: flex;
          flex-direction: column;
        }
        .adv-drawer.is-open { transform: translateX(0); }

        @media (max-width: 640px) {
          .adv-drawer {
            top: auto;
            left: 0; right: 0; bottom: 0;
            max-width: 100%;
            max-height: 88vh;
            border-left: none;
            border-top: 1px solid var(--c-border);
            border-radius: var(--r-lg) var(--r-lg) 0 0;
            transform: translateY(100%);
          }
          .adv-drawer.is-open { transform: translateY(0); }
        }

        .adv-header {
          display: flex; align-items: center; justify-content: space-between;
          padding: 0.95rem 1.1rem;
          border-bottom: 1px solid var(--c-divider);
          background: linear-gradient(180deg, var(--c-accent-green), transparent);
          flex-shrink: 0;
        }
        .adv-title {
          color: var(--c-primary-dark);
          font-family: var(--font-heading);
          font-weight: 800; font-size: 0.92rem;
          letter-spacing: -0.01em; margin: 0;
        }
        .adv-close {
          width: 1.95rem; height: 1.95rem;
          border-radius: var(--r-md);
          background: var(--c-bg-elevated);
          border: 1px solid var(--c-border-strong);
          color: var(--c-text-secondary);
          display: flex; align-items: center; justify-content: center;
          cursor: pointer;
          transition: all var(--t-base) var(--ease-out);
        }
        .adv-close:hover {
          background: var(--c-primary);
          color: white;
          border-color: var(--c-primary);
        }

        .adv-body {
          flex: 1;
          overflow-y: auto;
          padding: 1rem 1.1rem 1.1rem;
          background: var(--c-bg-elevated);
        }
      `}</style>
    </>,
    document.body
  );
}
