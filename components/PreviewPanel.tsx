"use client";

import React, { useEffect, useState } from 'react';
import { Download, ArrowLeft, Loader2, AlertCircle, X } from 'lucide-react';

interface Props {
  token: string;
  count: number;
  filename: string;
  onConfirm: () => void;
  onBack: () => void;
  isDownloading?: boolean;
}

/**
 * Grid de vista previa de las slides generadas. El usuario revisa antes
 * de descargar — si algo se ve raro puede volver al plan y reajustar
 * sin descargar un PPT roto. Cada PNG se sirve por
 * /api/pptx-preview?token=X&slide=N (PNGs ya generados al cerrar el SSE).
 */
export default function PreviewPanel({
  token, count, filename, onConfirm, onBack, isDownloading,
}: Props) {
  const [zoomIdx, setZoomIdx] = useState<number | null>(null);

  // Cerrar zoom con ESC + bloquear scroll mientras está abierto
  useEffect(() => {
    if (zoomIdx === null) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setZoomIdx(null);
    };
    window.addEventListener('keydown', onKey);
    const prev = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      window.removeEventListener('keydown', onKey);
      document.body.style.overflow = prev;
    };
  }, [zoomIdx]);

  return (
    <div className="prv-card animate-fade-in-up">
      {/* Header */}
      <div className="prv-header">
        <div className="prv-head-icon" aria-hidden>
          <span className="prv-head-dot" />
        </div>
        <div className="prv-head-text">
          <p className="prv-eyebrow">Vista previa generada</p>
          <h3 className="prv-title">Tu PowerPoint está listo. Revisa antes de descargar.</h3>
          <p className="prv-sub">
            <strong>{count}</strong> slide{count !== 1 ? 's' : ''} renderizada{count !== 1 ? 's' : ''}.
            Click en una miniatura para ampliarla. Si algo se ve raro, vuelve al plan
            y desactiva o ajusta la slide problemática antes de descargar.
          </p>
        </div>
      </div>

      {/* Grid */}
      <div className="prv-grid">
        {Array.from({ length: count }, (_, i) => i).map((i) => (
          <PreviewTile key={i} token={token} idx={i} onZoom={() => setZoomIdx(i)} />
        ))}
      </div>

      {/* Actions */}
      <div className="prv-actions">
        <button
          type="button"
          onClick={onBack}
          disabled={isDownloading}
          className="prv-btn-back press-on-active"
        >
          <ArrowLeft size={14} />
          Volver al plan
        </button>
        <button
          type="button"
          onClick={onConfirm}
          disabled={isDownloading}
          className="prv-btn-primary press-on-active"
        >
          {isDownloading ? (
            <>
              <Loader2 size={15} style={{ animation: 'spin 1s linear infinite' }} />
              Descargando…
            </>
          ) : (
            <>
              <Download size={15} />
              Descargar PowerPoint
            </>
          )}
        </button>
      </div>

      <p className="prv-foot">
        Archivo: <span className="prv-foot-name">{filename}</span>
      </p>

      {/* Lightbox */}
      {zoomIdx !== null && (
        <div className="prv-lightbox" onClick={() => setZoomIdx(null)} role="dialog" aria-modal="true">
          <button
            type="button"
            className="prv-lightbox-close"
            onClick={(e) => { e.stopPropagation(); setZoomIdx(null); }}
            aria-label="Cerrar vista ampliada"
          >
            <X size={20} />
          </button>
          <div className="prv-lightbox-pos">
            Slide {zoomIdx + 1} / {count}
          </div>
          <img
            src={`/api/pptx-preview?token=${encodeURIComponent(token)}&slide=${zoomIdx}`}
            alt={`Slide ${zoomIdx + 1}`}
            className="prv-lightbox-img"
            onClick={(e) => e.stopPropagation()}
          />
        </div>
      )}

      <style>{PRV_STYLES}</style>
    </div>
  );
}

function PreviewTile({ token, idx, onZoom }: {
  token: string; idx: number; onZoom: () => void;
}) {
  const [state, setState] = useState<'loading' | 'ready' | 'error'>('loading');
  const src = `/api/pptx-preview?token=${encodeURIComponent(token)}&slide=${idx}`;
  return (
    <button type="button" onClick={onZoom} className={`prv-tile is-${state}`}>
      {state === 'loading' && (
        <span className="prv-tile-skel">
          <Loader2 size={18} style={{ animation: 'spin 1s linear infinite' }} />
        </span>
      )}
      {state === 'error' && (
        <span className="prv-tile-err">
          <AlertCircle size={16} />
          <span>Sin vista previa</span>
        </span>
      )}
      <img
        src={src}
        alt={`Slide ${idx + 1}`}
        loading="lazy"
        decoding="async"
        onLoad={() => setState('ready')}
        onError={() => setState('error')}
        style={{ display: state === 'ready' ? 'block' : 'none' }}
      />
      <span className="prv-tile-num">{String(idx + 1).padStart(2, '0')}</span>
    </button>
  );
}

const PRV_STYLES = `
.prv-card {
  background: var(--c-bg-elevated);
  border: 1px solid var(--c-border);
  border-radius: var(--r-lg);
  padding: clamp(0.95rem, 1.8vw, 1.25rem);
  display: flex; flex-direction: column;
  gap: clamp(0.85rem, 1.4vw, 1.1rem);
  box-shadow: var(--shadow-md);
}

.prv-header {
  display: flex; gap: 0.85rem; align-items: flex-start;
  padding-bottom: 0.85rem;
  border-bottom: 1px solid var(--c-divider);
}
.prv-head-icon {
  background: var(--c-accent-green);
  border: 1px solid rgba(105, 190, 40, 0.40);
  border-radius: var(--r-md);
  padding: 0.5rem 0.55rem;
  flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
}
.prv-head-dot {
  width: 10px; height: 10px;
  border-radius: 50%;
  background: var(--c-logo-green);
  box-shadow: 0 0 0 4px rgba(105, 190, 40, 0.20);
}
.prv-head-text { flex: 1; min-width: 0; }
.prv-eyebrow {
  color: var(--c-primary);
  font-family: var(--font-heading);
  font-size: 0.66rem; font-weight: 800;
  letter-spacing: 0.08em; text-transform: uppercase;
  margin-bottom: 0.25rem;
}
.prv-title {
  font-family: var(--font-heading);
  font-size: clamp(0.95rem, 1.5vw, 1.05rem);
  font-weight: 800; color: var(--c-text-primary);
  line-height: 1.25; margin: 0;
}
.prv-sub {
  color: var(--c-text-secondary);
  font-size: 0.78rem; line-height: 1.5;
  margin-top: 0.4rem;
}
.prv-sub strong { color: var(--c-primary); font-weight: 700; }

.prv-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 0.75rem;
}
@media (min-width: 520px) {
  .prv-grid { grid-template-columns: repeat(2, 1fr); }
}
@media (min-width: 820px) {
  .prv-grid { grid-template-columns: repeat(3, 1fr); }
}

.prv-tile {
  position: relative;
  aspect-ratio: 16 / 9;
  background: var(--c-bg-tinted);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  overflow: hidden;
  padding: 0;
  cursor: pointer;
  transition: all 0.18s;
}
.prv-tile:hover:not(.is-error) {
  border-color: var(--c-primary);
  box-shadow: 0 4px 12px rgba(8, 112, 98, 0.18);
  transform: translateY(-2px);
}
.prv-tile img {
  width: 100%; height: 100%;
  object-fit: cover;
  object-position: top;
}
.prv-tile-skel,
.prv-tile-err {
  position: absolute; inset: 0;
  display: flex; align-items: center; justify-content: center;
  gap: 0.4rem;
  color: var(--c-text-tertiary);
  font-size: 0.74rem;
}
.prv-tile-err { color: var(--c-error-300); }
.prv-tile-num {
  position: absolute;
  top: 0.4rem; left: 0.4rem;
  background: rgba(18, 60, 73, 0.85);
  color: white;
  font-family: var(--font-heading);
  font-size: 0.66rem; font-weight: 800;
  padding: 0.2rem 0.45rem;
  border-radius: var(--r-pill);
  letter-spacing: 0.04em;
}

/* Actions */
.prv-actions {
  display: flex; gap: 0.6rem;
  flex-wrap: wrap;
  padding-top: 0.4rem;
  border-top: 1px solid var(--c-divider);
}
.prv-btn-back {
  padding: 0.85rem 1.1rem;
  background: var(--c-bg-elevated);
  border: 1.5px solid var(--c-border-strong);
  border-radius: var(--r-md);
  color: var(--c-text-secondary);
  font-family: var(--font-heading);
  font-size: 0.78rem; font-weight: 700;
  letter-spacing: 0.04em; text-transform: uppercase;
  display: flex; align-items: center; gap: 0.4rem;
  cursor: pointer;
  transition: all 0.18s;
}
.prv-btn-back:hover:not(:disabled) {
  background: var(--c-bg-tinted);
  border-color: var(--c-primary);
  color: var(--c-primary);
}
.prv-btn-back:disabled { opacity: 0.5; cursor: not-allowed; }
.prv-btn-primary {
  flex: 1; min-width: 240px;
  padding: 0.95rem 1.1rem;
  background: var(--c-primary);
  border: none;
  border-radius: var(--r-md);
  color: white;
  font-family: var(--font-heading);
  font-size: 0.9rem; font-weight: 800;
  letter-spacing: 0.04em; text-transform: uppercase;
  display: flex; align-items: center; justify-content: center; gap: 0.5rem;
  box-shadow: 0 6px 18px rgba(8, 112, 98, 0.28);
  cursor: pointer;
  transition: all 0.18s;
}
.prv-btn-primary:hover:not(:disabled) {
  background: var(--c-primary-dark);
  transform: translateY(-1px);
  box-shadow: 0 10px 24px rgba(8, 112, 98, 0.36);
}
.prv-btn-primary:disabled { opacity: 0.55; cursor: not-allowed; box-shadow: none; }

.prv-foot {
  color: var(--c-text-tertiary);
  font-size: 0.7rem;
}
.prv-foot-name {
  color: var(--c-text-primary);
  font-family: var(--font-heading);
  font-weight: 700;
}

/* Lightbox */
.prv-lightbox {
  position: fixed; inset: 0;
  background: rgba(18, 60, 73, 0.85);
  backdrop-filter: blur(6px);
  z-index: 200;
  display: flex; align-items: center; justify-content: center;
  padding: 1.5rem;
  cursor: zoom-out;
  animation: fadeIn 0.18s ease-out;
}
.prv-lightbox-close {
  position: absolute;
  top: 1rem; right: 1rem;
  width: 2.4rem; height: 2.4rem;
  border-radius: var(--r-md);
  background: rgba(255, 255, 255, 0.10);
  border: 1px solid rgba(255, 255, 255, 0.25);
  color: white;
  display: flex; align-items: center; justify-content: center;
  cursor: pointer;
  transition: background 0.18s;
}
.prv-lightbox-close:hover { background: rgba(255, 255, 255, 0.20); }
.prv-lightbox-pos {
  position: absolute;
  top: 1.2rem; left: 1.2rem;
  color: white;
  font-family: var(--font-heading);
  font-size: 0.78rem; font-weight: 700;
  letter-spacing: 0.06em; text-transform: uppercase;
  background: rgba(0, 0, 0, 0.40);
  padding: 0.4rem 0.75rem;
  border-radius: var(--r-pill);
}
.prv-lightbox-img {
  max-width: 100%;
  max-height: 100%;
  object-fit: contain;
  border-radius: var(--r-md);
  box-shadow: 0 20px 60px rgba(0, 0, 0, 0.45);
  cursor: default;
}
`;
