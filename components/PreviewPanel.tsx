"use client";

import React, { useEffect, useState } from 'react';
import { Download, ArrowLeft, Loader2, AlertCircle, X, FileText, ChevronLeft, ChevronRight, ZoomIn, ZoomOut } from 'lucide-react';

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
  // imgZoomed: false = fit-to-view (ve la slide entera, sin scroll).
  // true = tamaño grande (desborda viewport → aparece scroll real). Click en
  // la imagen alterna entre los dos modos. Reset a false al cambiar de slide.
  const [imgZoomed, setImgZoomed] = useState(false);
  // Estado de carga de la imagen en el lightbox. Sin esto, si el server
  // devuelve 404 (token expirado, preview borrado), el <img> queda invisible
  // y el usuario solo ve "Slide N / total" sin entender qué pasó.
  const [imgState, setImgState] = useState<'loading' | 'ready' | 'error'>('loading');
  // Cache-buster por sesión de lightbox — si el usuario reabre la misma
  // slide después de un fallo, fuerza nuevo request en vez de servir el 404
  // cacheado por el browser.
  const [imgNonce, setImgNonce] = useState(0);
  const [pdfState, setPdfState] = useState<'idle' | 'fetching' | 'error'>('idle');
  const [pdfError, setPdfError] = useState<string | null>(null);

  const handleDownloadPdf = async () => {
    if (!token) return;
    setPdfState('fetching');
    setPdfError(null);
    try {
      const res = await fetch(`/api/pptx-pdf?token=${encodeURIComponent(token)}`);
      if (!res.ok) {
        const j = await res.json().catch(() => null);
        throw new Error(j?.error || `HTTP ${res.status}`);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename.replace(/\.pptx$/i, '.pdf');
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setPdfState('idle');
    } catch (err: unknown) {
      setPdfState('error');
      setPdfError(err instanceof Error ? err.message : 'Error al exportar PDF.');
    }
  };

  // ESC cierra, flechas ←/→ navegan entre slides. Bloqueamos el scroll del
  // body mientras el lightbox está abierto para que solo scrollee el modal.
  useEffect(() => {
    if (zoomIdx === null) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setZoomIdx(null);
      } else if (e.key === 'ArrowRight') {
        e.preventDefault();
        setZoomIdx((cur) => (cur === null ? cur : Math.min(count - 1, cur + 1)));
      } else if (e.key === 'ArrowLeft') {
        e.preventDefault();
        setZoomIdx((cur) => (cur === null ? cur : Math.max(0, cur - 1)));
      }
    };
    window.addEventListener('keydown', onKey);
    const prev = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      window.removeEventListener('keydown', onKey);
      document.body.style.overflow = prev;
    };
  }, [zoomIdx, count]);

  const goPrev = () => setZoomIdx((cur) => (cur === null ? cur : Math.max(0, cur - 1)));
  const goNext = () => setZoomIdx((cur) => (cur === null ? cur : Math.min(count - 1, cur + 1)));

  // Reset el modo zoom + estado de carga cada vez que el usuario salta a
  // otra slide. Si no reseteamos imgState, el spinner queda "viejo" de la
  // slide anterior y el usuario ve estado inconsistente.
  useEffect(() => {
    if (zoomIdx !== null) {
      setImgZoomed(false);
      setImgState('loading');
    }
  }, [zoomIdx]);

  // Handler para reintentar carga manualmente cuando hay error. Bumpea el
  // nonce → cambia el src → React monta nueva <img> → nuevo fetch al server.
  const retryImage = () => {
    setImgState('loading');
    setImgNonce((n) => n + 1);
  };

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
          onClick={handleDownloadPdf}
          disabled={isDownloading || pdfState === 'fetching'}
          className="prv-btn-pdf press-on-active"
          title="Convierte el PPTX a PDF — útil para compartir sin Office"
        >
          {pdfState === 'fetching' ? (
            <>
              <Loader2 size={14} style={{ animation: 'spin 1s linear infinite' }} />
              Generando…
            </>
          ) : (
            <>
              <FileText size={14} />
              Descargar PDF
            </>
          )}
        </button>
        <button
          type="button"
          onClick={onConfirm}
          disabled={isDownloading || pdfState === 'fetching'}
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

      {pdfError && (
        <div className="prv-pdf-err" role="alert">
          <AlertCircle size={13} />
          <span>{pdfError}</span>
        </div>
      )}

      <p className="prv-foot">
        Archivo: <span className="prv-foot-name">{filename}</span>
      </p>

      {/* Lightbox v3 — patrón simple bulletproof:
          - Container fijo a viewport con overflow:auto → SIEMPRE puede scrollear
          - Toolbar sticky top:0 → siempre visible mientras scrolleás
          - Imagen como BLOCK con margin:0 auto → centrada horizontal sin flex
          - Sin "modos" complejos: dos clases simples (medium/big) que cambian
            ancho. Click en imagen alterna entre las dos.
          - Click en backdrop fuera de la imagen cierra. */}
      {zoomIdx !== null && (
        <div
          className="prv-lightbox"
          role="dialog"
          aria-modal="true"
          onClick={(e) => {
            // Solo cerrar si el click fue en el container backdrop, no en
            // contenido interno (toolbar, imagen, flechas). Sin esto, el
            // click en cualquier parte cierra demasiado fácil.
            if (e.target === e.currentTarget) setZoomIdx(null);
          }}
        >
          {/* Toolbar SIEMPRE visible (sticky top) */}
          <div className="prv-lightbox-bar">
            <span className="prv-lightbox-pos">
              Slide <strong>{zoomIdx + 1}</strong> / {count}
            </span>
            <div className="prv-lightbox-actions">
              {zoomIdx > 0 && (
                <button
                  type="button"
                  className="prv-lightbox-btn"
                  onClick={goPrev}
                  aria-label="Slide anterior"
                  title="Anterior (←)"
                >
                  <ChevronLeft size={18} />
                </button>
              )}
              {zoomIdx < count - 1 && (
                <button
                  type="button"
                  className="prv-lightbox-btn"
                  onClick={goNext}
                  aria-label="Slide siguiente"
                  title="Siguiente (→)"
                >
                  <ChevronRight size={18} />
                </button>
              )}
              <button
                type="button"
                className="prv-lightbox-btn"
                onClick={() => setImgZoomed((z) => !z)}
                aria-label={imgZoomed ? 'Reducir' : 'Ampliar'}
                title={imgZoomed ? 'Reducir' : 'Ampliar (clic en imagen)'}
              >
                {imgZoomed ? <ZoomOut size={16} /> : <ZoomIn size={16} />}
                <span className="prv-lightbox-btn-label">
                  {imgZoomed ? 'Reducir' : 'Ampliar'}
                </span>
              </button>
              <button
                type="button"
                className="prv-lightbox-btn prv-lightbox-btn-close"
                onClick={() => setZoomIdx(null)}
                aria-label="Cerrar vista ampliada"
                title="Cerrar (Esc)"
              >
                <X size={18} />
                <span className="prv-lightbox-btn-label">Cerrar</span>
              </button>
            </div>
          </div>

          {/* Contenido: la imagen vive directamente acá. El container padre
              (prv-lightbox) tiene overflow:auto → scrollea cuando la imagen
              + padding excede el viewport. SIN flex centering. */}
          <div className="prv-lightbox-content">
            {imgState === 'loading' && (
              <div className="prv-lightbox-msg">
                <Loader2 size={28} style={{ animation: 'spin 1s linear infinite' }} />
                <span>Cargando vista previa…</span>
              </div>
            )}
            {imgState === 'error' && (
              <div className="prv-lightbox-msg prv-lightbox-msg-err">
                <AlertCircle size={28} />
                <span>No se pudo cargar la vista previa.</span>
                <span className="prv-lightbox-msg-hint">
                  La vista previa puede haber expirado. Probá reintentar o
                  cerrá y volvé a abrir.
                </span>
                <button
                  type="button"
                  className="prv-lightbox-btn"
                  onClick={retryImage}
                >
                  Reintentar
                </button>
              </div>
            )}
            <img
              key={`${zoomIdx}-${imgNonce}`}
              src={`/api/pptx-preview?token=${encodeURIComponent(token)}&slide=${zoomIdx}&t=${imgNonce}`}
              alt={`Slide ${zoomIdx + 1}`}
              className={`prv-lightbox-img ${imgZoomed ? 'is-big' : 'is-medium'}`}
              onLoad={() => setImgState('ready')}
              onError={() => setImgState('error')}
              onClick={(e) => {
                e.stopPropagation();
                if (imgState === 'ready') setImgZoomed((z) => !z);
              }}
              style={{ display: imgState === 'ready' ? 'block' : 'none' }}
            />
          </div>
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
  // Cache-buster por mount para evitar que el browser sirva una respuesta
  // de prueba cacheada de un token previo expirado.
  const [src] = useState(
    () => `/api/pptx-preview?token=${encodeURIComponent(token)}&slide=${idx}&t=${Date.now()}`
  );

  // Timeout duro: si el server no responde en 20s o el browser se cuelga
  // en lazy-load, marcamos error para que el usuario vea "Sin vista previa"
  // en vez de un spinner infinito. Antes loading="lazy" + tiles fuera del
  // viewport en mobile podían quedarse cargando para siempre.
  useEffect(() => {
    if (state !== 'loading') return;
    const timer = setTimeout(() => {
      setState((current) => (current === 'loading' ? 'error' : current));
    }, 20_000);
    return () => clearTimeout(timer);
  }, [state]);

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
        // loading="eager" — antes era "lazy" pero en mobile o viewports
        // chicos los tiles fuera de pantalla nunca disparaban onLoad y se
        // quedaban con el spinner indefinidamente.
        loading="eager"
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

.prv-btn-pdf {
  padding: 0.85rem 1rem;
  background: var(--c-bg-elevated);
  border: 1.5px solid var(--c-primary);
  border-radius: var(--r-md);
  color: var(--c-primary);
  font-family: var(--font-heading);
  font-size: 0.78rem; font-weight: 700;
  letter-spacing: 0.04em; text-transform: uppercase;
  display: flex; align-items: center; justify-content: center; gap: 0.4rem;
  cursor: pointer;
  transition: all 0.18s;
  white-space: nowrap;
}
.prv-btn-pdf:hover:not(:disabled) {
  background: var(--c-accent-green);
  color: var(--c-primary-dark);
}
.prv-btn-pdf:disabled { opacity: 0.5; cursor: not-allowed; }

.prv-pdf-err {
  display: inline-flex; align-items: center; gap: 0.4rem;
  padding: 0.5rem 0.7rem;
  background: #FEF2F2;
  border: 1px solid rgba(212, 56, 56, 0.30);
  border-radius: var(--r-md);
  color: var(--c-error-300);
  font-size: 0.74rem;
}

.prv-foot {
  color: var(--c-text-tertiary);
  font-size: 0.7rem;
}
.prv-foot-name {
  color: var(--c-text-primary);
  font-family: var(--font-heading);
  font-weight: 700;
}

/* ──────────────────────────────────────────────────────────────────
   Lightbox v3 — patrón super simple. Container fijo con overflow:auto
   (SIEMPRE scrolleable), toolbar sticky top, imagen block con margin
   auto. Sin flex centering, sin wrapper-div, sin "safe" — solo CSS
   básico que funciona en todos los browsers.
   ────────────────────────────────────────────────────────────────── */
.prv-lightbox {
  position: fixed; inset: 0;
  z-index: 200;
  background: rgba(18, 60, 73, 0.92);
  backdrop-filter: blur(6px);
  /* Aquí está la clave: el container ENTERO scrollea. Cuando la imagen
     es alta, el scroll natural del navegador funciona sin trucos. */
  overflow: auto;
  -webkit-overflow-scrolling: touch;
  animation: fadeIn 0.18s ease-out;
}

/* Toolbar sticky → siempre visible aunque scrollees hacia abajo. */
.prv-lightbox-bar {
  position: sticky;
  top: 0;
  z-index: 2;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  padding: 0.7rem clamp(0.75rem, 2vw, 1.25rem);
  background: rgba(0, 0, 0, 0.55);
  backdrop-filter: blur(6px);
  border-bottom: 1px solid rgba(255, 255, 255, 0.10);
}
.prv-lightbox-pos {
  color: rgba(255, 255, 255, 0.85);
  font-family: var(--font-heading);
  font-size: 0.78rem; font-weight: 600;
  letter-spacing: 0.04em; text-transform: uppercase;
  white-space: nowrap;
}
.prv-lightbox-pos strong {
  color: white; font-weight: 800;
}
.prv-lightbox-actions {
  display: flex; gap: 0.4rem;
  flex-wrap: nowrap;
}
.prv-lightbox-btn {
  display: inline-flex; align-items: center; gap: 0.4rem;
  padding: 0.55rem 0.8rem;
  background: rgba(255, 255, 255, 0.12);
  border: 1px solid rgba(255, 255, 255, 0.28);
  border-radius: var(--r-md);
  color: white;
  font-family: var(--font-heading);
  font-size: 0.74rem; font-weight: 700;
  letter-spacing: 0.06em; text-transform: uppercase;
  cursor: pointer;
  transition: background 0.18s, transform 0.18s;
  white-space: nowrap;
}
.prv-lightbox-btn:hover { background: rgba(255, 255, 255, 0.22); }
.prv-lightbox-btn:active { transform: translateY(1px); }
.prv-lightbox-btn-close {
  background: rgba(212, 56, 56, 0.35);
  border-color: rgba(255, 130, 130, 0.55);
}
.prv-lightbox-btn-close:hover { background: rgba(212, 56, 56, 0.65); }

/* Content wrapper — solo padding. La imagen vive directamente acá. NO
   tiene overflow propio: el scroll lo provee el container padre. */
.prv-lightbox-content {
  padding: clamp(0.75rem, 2vw, 1.5rem);
}

.prv-lightbox-img {
  display: block;
  margin: 0 auto;        /* centrado horizontal sin flex */
  border-radius: var(--r-md);
  box-shadow: 0 20px 60px rgba(0, 0, 0, 0.45);
  height: auto;
  transition: width 0.2s ease;
}
/* Tamaño "mediano" — entra cómodo en cualquier viewport. */
.prv-lightbox-img.is-medium {
  width: min(1400px, 95vw);
  max-width: 100%;
  cursor: zoom-in;
}
/* Tamaño "grande" — fuerza scroll horizontal+vertical. max() garantiza
   que en pantallas anchas (>1920px) la imagen igual desborde. */
.prv-lightbox-img.is-big {
  width: max(1800px, 150vw);
  max-width: none;
  cursor: zoom-out;
}

/* Mensajes loading/error — se posicionan donde iría la imagen. */
.prv-lightbox-msg {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 0.7rem;
  margin: 4rem auto;
  color: rgba(255, 255, 255, 0.85);
  font-family: var(--font-heading);
  font-size: 0.95rem; font-weight: 600;
  letter-spacing: 0.02em;
  text-align: center;
  max-width: 28rem;
  padding: 1.5rem;
}
.prv-lightbox-msg-err { color: #FFC4C4; }
.prv-lightbox-msg-hint {
  font-size: 0.78rem; font-weight: 400;
  color: rgba(255, 255, 255, 0.65);
  line-height: 1.5;
  letter-spacing: 0;
}

/* Mobile: oculto labels, achico padding. */
@media (max-width: 560px) {
  .prv-lightbox-btn { padding: 0.5rem 0.55rem; gap: 0.3rem; }
  .prv-lightbox-btn-label { display: none; }
  .prv-lightbox-bar { padding: 0.6rem 0.75rem; }
}
`;
