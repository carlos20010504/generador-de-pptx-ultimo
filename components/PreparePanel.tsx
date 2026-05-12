"use client";

import React, { useEffect, useMemo, useState } from 'react';
import {
  Sparkles, FileSpreadsheet, AlertTriangle, AlertCircle,
  TrendingUp, Layers, BarChart2, Loader2, ChevronDown, ChevronRight,
  Lightbulb, CheckSquare, Square, Table2, FileText, LayoutDashboard,
} from 'lucide-react';

/* ──────────────────────────────────────────────────────────────
   Types — mirror what /api/quick-summary and /api/preview-plan return
   ────────────────────────────────────────────────────────────── */

interface QuickSummary {
  filename: string;
  totals: {
    sheets: number;
    rows: number;
    columns: number;
    kpi_candidates: number;
    categorical_distributions: number;
    time_series: number;
  };
  kpis_preview: Array<{ label: string; value: string; kind: string; sheet: string }>;
  warnings: Array<{ severity: 'info' | 'warn' | 'error'; icon: string; title: string; detail: string }>;
  suggestions: Array<{ id: string; label: string; prompt: string; why: string }>;
  deck_estimate: { min_slides: number; max_slides: number; expected_sections: string[] };
}

interface SlidePreview {
  index: number;
  type: string;
  title: string;
  subtitle?: string;
  narrative?: string;
  provenance?: { sheet?: string; columns?: string[] } | null;
  mandatory?: boolean;
}

interface PreviewResponse {
  meta: { title?: string; subtitle?: string };
  slides: SlidePreview[];
  ai_status?: { model?: string; cache_hit?: boolean };
}

type Mode = 'mixed' | 'charts' | 'tables' | 'boardroom' | 'auto';

interface ThemeShape {
  key: string;
  name: string;
  primary_hex: string;
  accent_hex: string;
  text_hex: string;
  bg_hex: string;
}

interface Props {
  file: File;
  userPrompt: string;
  onPromptChange: (p: string) => void;
  audience: string;
  language: string;
  theme: ThemeShape;
  mode: Mode;
  onModeChange: (m: Mode) => void;
  onConfirm: (excludedIndices: number[]) => void;
  onOpenAdvanced?: () => void;
}

const SLIDE_TYPE_META: Record<string, { Ic: React.ComponentType<{ size?: number; color?: string }>; label: string; color: string }> = {
  title:        { Ic: FileSpreadsheet, label: 'Portada',   color: '#FBBF24' },
  kpi_row:      { Ic: LayoutDashboard, label: 'KPIs',      color: '#60A5FA' },
  chart:        { Ic: BarChart2,       label: 'Gráfico',   color: '#A78BFA' },
  table:        { Ic: Table2,          label: 'Tabla',     color: '#86EFAC' },
  text_bullets: { Ic: FileText,        label: 'Hallazgos', color: '#FDA4AF' },
  unknown:      { Ic: FileText,        label: 'Slide',     color: '#9CA3AF' },
};

const MODE_LABELS: Record<Mode, string> = {
  mixed: 'Mixto',
  charts: 'Gráficos',
  tables: 'Tablas',
  boardroom: 'Ejecutivo',
  auto: 'Auto',
};

export default function PreparePanel({
  file, userPrompt, onPromptChange,
  audience, language, theme, mode, onModeChange,
  onConfirm, onOpenAdvanced,
}: Props) {
  const [summary, setSummary] = useState<QuickSummary | null>(null);
  const [summaryError, setSummaryError] = useState<string | null>(null);
  const [summaryLoading, setSummaryLoading] = useState(true);

  const [plan, setPlan] = useState<PreviewResponse | null>(null);
  const [planError, setPlanError] = useState<string | null>(null);
  const [planLoading, setPlanLoading] = useState(true);

  const [excluded, setExcluded] = useState<Set<number>>(new Set());
  const [expandedSlide, setExpandedSlide] = useState<number | null>(null);
  const [refineOpen, setRefineOpen] = useState(false);

  // Cheap stable identity — depending on the File object itself causes
  // re-fetches on every parent re-render which abort in-flight requests.
  const fileKey = `${file.name}|${file.size}|${file.lastModified}`;
  const promptKey = `${fileKey}|${userPrompt}`;

  // Quick-summary: refetches only when file changes.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      setSummaryLoading(true);
      setSummaryError(null);
      try {
        const fd = new FormData();
        fd.append('file', file);
        const res = await fetch('/api/quick-summary', { method: 'POST', body: fd });
        if (cancelled) return;
        if (!res.ok) {
          const j = await res.json().catch(() => null);
          const msg = typeof j?.error === 'string' ? j.error : (j?.error?.message || `Error ${res.status}`);
          throw new Error(msg);
        }
        const data: QuickSummary = await res.json();
        if (!cancelled) setSummary(data);
      } catch (err: unknown) {
        if (!cancelled) setSummaryError((err as Error).message || 'No se pudo analizar el archivo.');
      } finally {
        if (!cancelled) setSummaryLoading(false);
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileKey]);

  // Preview-plan: refetches when file or prompt change. Cache hit on backend
  // for same (file, prompt) — safe to re-run without quota cost.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      setPlanLoading(true);
      setPlanError(null);
      try {
        const fd = new FormData();
        fd.append('file', file);
        if (userPrompt.trim()) fd.append('userPrompt', userPrompt.trim());
        fd.append('audience', audience);
        fd.append('language', language);
        const res = await fetch('/api/preview-plan', { method: 'POST', body: fd });
        if (cancelled) return;
        if (!res.ok) {
          const j = await res.json().catch(() => null);
          const msg = typeof j?.error === 'string' ? j.error : (j?.error?.message || `Error ${res.status}`);
          throw new Error(msg);
        }
        const data: PreviewResponse = await res.json();
        if (!cancelled) setPlan(data);
      } catch (err: unknown) {
        if (!cancelled) setPlanError((err as Error).message || 'No se pudo generar el plan.');
      } finally {
        if (!cancelled) setPlanLoading(false);
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [promptKey, audience, language]);

  const stats = useMemo(() => {
    if (!plan) return { total: 0, kept: 0, byType: {} as Record<string, number> };
    const total = plan.slides.length;
    const kept = plan.slides.filter((_, i) => !excluded.has(i)).length;
    const byType: Record<string, number> = {};
    plan.slides.forEach((s, i) => {
      if (!excluded.has(i)) byType[s.type] = (byType[s.type] || 0) + 1;
    });
    return { total, kept, byType };
  }, [plan, excluded]);

  // Filter warnings: hide info-only, group as "+N avisos menores" if any.
  const visibleWarnings = useMemo(() => {
    if (!summary) return { major: [], minorCount: 0 };
    const major = summary.warnings.filter(w => w.severity === 'warn' || w.severity === 'error');
    const minorCount = summary.warnings.length - major.length;
    return { major, minorCount };
  }, [summary]);

  const topKpis = summary?.kpis_preview.slice(0, 3) ?? [];

  const toggle = (idx: number, mandatory: boolean) => {
    if (mandatory) return;
    setExcluded(prev => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx);
      return next;
    });
  };

  const handleSuggestionClick = (s: QuickSummary['suggestions'][0]) => {
    onPromptChange(s.prompt);
    setRefineOpen(true);
  };

  const headerStats = summary ? (
    <p className="prep-head-stats">
      <strong className="prep-head-strong">Análisis listo:</strong>{' '}
      {summary.totals.sheets} hoja{summary.totals.sheets !== 1 ? 's' : ''}
      {' · '}{summary.totals.rows.toLocaleString('es-CO')} filas
      {' · '}{summary.totals.kpi_candidates} KPIs
      {' · ~'}{summary.deck_estimate.min_slides}–{summary.deck_estimate.max_slides} slides
    </p>
  ) : (
    <p className="prep-head-stats prep-head-stats-loading">Analizando archivo…</p>
  );

  return (
    <div className="prep-card animate-fade-in-up">
      {/* ────────────────────────────────────────────────
           Header (compact stats — file name lives in the dropzone)
         ──────────────────────────────────────────────── */}
      <div className="prep-header">
        <div className="prep-head-icon">
          <Sparkles size={16} color="#A78BFA" />
        </div>
        <div className="prep-head-text">
          {headerStats}
        </div>
      </div>

      {/* ────────────────────────────────────────────────
           Section 1 — Detección (chips + KPIs + warnings)
         ──────────────────────────────────────────────── */}
      <section className="prep-section">
        {summaryLoading && !summary ? (
          <SectionSkeleton label="Analizando estructura del Excel…" />
        ) : summaryError ? (
          <div className="prep-banner is-warn">
            <AlertTriangle size={14} />
            <span>{summaryError}</span>
          </div>
        ) : summary ? (
          <>
            <div className="prep-stats">
              <Stat icon={TrendingUp} label="KPIs" value={summary.totals.kpi_candidates} color="#60A5FA" />
              <Stat icon={BarChart2} label="Distribuciones" value={summary.totals.categorical_distributions} color="#A78BFA" />
              <Stat icon={Layers} label="Series" value={summary.totals.time_series} color="#86EFAC" />
              <Stat
                icon={Sparkles}
                label="Slides est."
                value={`${summary.deck_estimate.min_slides}–${summary.deck_estimate.max_slides}`}
                color="#FBBF24"
              />
            </div>

            {topKpis.length > 0 && (
              <div className="prep-kpi-row">
                {topKpis.map((k, i) => (
                  <span key={i} className="prep-kpi-chip" title={k.label}>
                    <span className="prep-kpi-val">{k.value}</span>
                    <span className="prep-kpi-lbl">{k.label}</span>
                  </span>
                ))}
                {summary.kpis_preview.length > 3 && (
                  <span className="prep-kpi-more">+{summary.kpis_preview.length - 3}</span>
                )}
              </div>
            )}

            {(visibleWarnings.major.length > 0 || visibleWarnings.minorCount > 0) && (
              <div className="prep-warnings">
                {visibleWarnings.major.map((w, i) => {
                  const Ic = w.severity === 'error' ? AlertCircle : AlertTriangle;
                  return (
                    <div key={i} className={`prep-warn is-${w.severity}`}>
                      <Ic size={12} />
                      <span><strong>{w.title}.</strong> {w.detail}</span>
                    </div>
                  );
                })}
                {visibleWarnings.minorCount > 0 && (
                  <div className="prep-warn-minor">
                    +{visibleWarnings.minorCount} aviso{visibleWarnings.minorCount !== 1 ? 's' : ''} menor{visibleWarnings.minorCount !== 1 ? 'es' : ''}
                  </div>
                )}
              </div>
            )}
          </>
        ) : null}
      </section>

      {/* ────────────────────────────────────────────────
           Section 2 — Plan de slides
         ──────────────────────────────────────────────── */}
      <section className="prep-section">
        <div className="prep-section-head">
          <h4 className="prep-section-title">Plan de slides</h4>
          <span className="prep-counter">
            <span className="prep-counter-num">{stats.kept}</span>
            <span className="prep-counter-sep">/</span>
            <span className="prep-counter-tot">{stats.total}</span>
            <span className="prep-counter-lbl">activas</span>
          </span>
        </div>

        {planLoading && !plan ? (
          <SectionSkeleton label="Construyendo el plan con IA (1 sola vez)…" />
        ) : planError ? (
          <div className="prep-banner is-error">
            <AlertCircle size={14} />
            <span>{planError}</span>
          </div>
        ) : plan ? (
          <>
            <ul className="prep-slides">
              {plan.slides.map((slide) => {
                const meta = SLIDE_TYPE_META[slide.type] || SLIDE_TYPE_META.unknown;
                const isExcluded = excluded.has(slide.index);
                const isMandatory = !!slide.mandatory;
                const isExpanded = expandedSlide === slide.index;
                return (
                  <li key={slide.index}>
                    <button
                      type="button"
                      onClick={() => toggle(slide.index, isMandatory)}
                      onDoubleClick={() => setExpandedSlide(isExpanded ? null : slide.index)}
                      disabled={isMandatory}
                      className={[
                        'prep-slide',
                        isExcluded ? 'is-excluded' : 'is-active',
                        isMandatory ? 'is-mandatory' : '',
                      ].filter(Boolean).join(' ')}
                    >
                      <span className="prep-slide-check">
                        {isExcluded
                          ? <Square size={15} color="rgba(255,255,255,0.30)" />
                          : <CheckSquare size={15} color={isMandatory ? '#FBBF24' : '#86EFAC'} />}
                      </span>
                      <span className="prep-slide-num">{String(slide.index + 1).padStart(2, '0')}</span>
                      <meta.Ic size={12} color={meta.color} />
                      <span className="prep-slide-type" style={{ color: meta.color }}>{meta.label}</span>
                      <span className="prep-slide-title">{slide.title || `${meta.label} sin título`}</span>
                      {isMandatory && <span className="prep-slide-pin">obligatoria</span>}
                      <span
                        className="prep-slide-toggle"
                        onClick={(e) => {
                          e.stopPropagation();
                          setExpandedSlide(isExpanded ? null : slide.index);
                        }}
                        role="button"
                        aria-label={isExpanded ? 'Ocultar detalles' : 'Ver detalles'}
                      >
                        {isExpanded ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                      </span>
                    </button>
                    {isExpanded && (
                      <div className="prep-slide-detail">
                        {slide.subtitle && <p className="prep-slide-sub">{slide.subtitle}</p>}
                        {slide.narrative && <p className="prep-slide-narr">{slide.narrative}</p>}
                        {slide.provenance?.sheet && (
                          <p className="prep-slide-source">
                            Fuente: <strong>{slide.provenance.sheet}</strong>
                            {slide.provenance.columns?.length
                              ? ` · ${slide.provenance.columns.slice(0, 4).join(', ')}${slide.provenance.columns.length > 4 ? '…' : ''}`
                              : ''}
                          </p>
                        )}
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>

            <div className="prep-pills">
              {Object.entries(stats.byType).map(([type, n]) => {
                const meta = SLIDE_TYPE_META[type] || SLIDE_TYPE_META.unknown;
                return (
                  <span key={type} className="prep-pill" style={{
                    background: `${meta.color}14`,
                    border: `1px solid ${meta.color}30`,
                    color: meta.color,
                  }}>
                    <meta.Ic size={10} color={meta.color} />
                    {n} {meta.label.toLowerCase()}
                  </span>
                );
              })}
              {plan.ai_status?.cache_hit && (
                <span className="prep-pill prep-pill-cache">cache · 0 tokens AI</span>
              )}
            </div>
          </>
        ) : null}
      </section>

      {/* ────────────────────────────────────────────────
           Section 3 — Refinar (collapsed by default)
         ──────────────────────────────────────────────── */}
      <section className="prep-section prep-refine">
        <button
          type="button"
          onClick={() => setRefineOpen(o => !o)}
          className="prep-refine-toggle"
        >
          <span className="prep-refine-left">
            {refineOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            <span className="prep-refine-label">Refinar prompt y contexto</span>
          </span>
          <span className="prep-refine-right">
            <span className="prep-mode-pill">{MODE_LABELS[mode] || 'Mixto'}</span>
          </span>
        </button>

        {refineOpen && (
          <div className="prep-refine-body animate-fade-in">
            <textarea
              value={userPrompt}
              onChange={(e) => onPromptChange(e.target.value)}
              placeholder="Ej: enfócate en métricas trimestrales, comparativas regionales, riesgos críticos…"
              className="prep-prompt"
              rows={2}
            />

            {summary && summary.suggestions.length > 0 && (
              <div className="prep-suggestions">
                <p className="prep-sugg-label">
                  <Lightbulb size={11} />
                  Ideas rápidas
                </p>
                <div className="prep-sugg-row">
                  {summary.suggestions.slice(0, 4).map((s) => (
                    <button
                      key={s.id}
                      type="button"
                      onClick={() => handleSuggestionClick(s)}
                      className={`prep-sugg ${userPrompt === s.prompt ? 'is-active' : ''}`}
                      title={s.why}
                    >
                      {s.label}
                    </button>
                  ))}
                </div>
              </div>
            )}

            <div className="prep-mode-row">
              <label className="prep-mode-label">Estilo visual</label>
              <div className="prep-mode-options">
                {(['mixed', 'charts', 'tables', 'boardroom'] as Mode[]).map((m) => (
                  <button
                    key={m}
                    type="button"
                    onClick={() => onModeChange(m)}
                    className={`prep-mode-opt ${mode === m ? 'is-active' : ''}`}
                  >
                    {MODE_LABELS[m]}
                  </button>
                ))}
              </div>
            </div>

            {onOpenAdvanced && (
              <p className="prep-refine-hint">
                ¿Necesitas más control (tema, audiencia, sugerencias completas)? Abre el{' '}
                <button type="button" className="prep-link" onClick={onOpenAdvanced}>panel avanzado</button>.
              </p>
            )}
          </div>
        )}
      </section>

      {/* ────────────────────────────────────────────────
           CTA
         ──────────────────────────────────────────────── */}
      <button
        type="button"
        onClick={() => onConfirm(Array.from(excluded))}
        disabled={planLoading || !plan || stats.kept === 0}
        className="prep-cta press-on-active"
      >
        <Sparkles size={15} />
        {planLoading
          ? 'Preparando plan…'
          : stats.kept === 0
            ? 'Activa al menos una slide'
            : `Generar PowerPoint con ${stats.kept} slide${stats.kept !== 1 ? 's' : ''}`}
      </button>

      <style>{PREP_STYLES}</style>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────────
   Small subcomponents
   ────────────────────────────────────────────────────────────── */

function Stat({ icon: Ic, label, value, color }: {
  icon: React.ComponentType<{ size?: number; color?: string }>;
  label: string; value: string | number; color: string;
}) {
  return (
    <div className="prep-stat">
      <span className="prep-stat-icon" style={{ background: `${color}1f`, border: `1px solid ${color}40` }}>
        <Ic size={12} color={color} />
      </span>
      <span className="prep-stat-val" style={{ color }}>{value}</span>
      <span className="prep-stat-lbl">{label}</span>
    </div>
  );
}

function SectionSkeleton({ label }: { label: string }) {
  return (
    <div className="prep-skel">
      <Loader2 size={16} className="prep-spin" color="var(--c-brand-blue-400)" />
      <span>{label}</span>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────────
   Styles
   ────────────────────────────────────────────────────────────── */

const PREP_STYLES = `
.prep-card {
  background: rgba(255, 255, 255, 0.04);
  backdrop-filter: blur(20px) saturate(160%);
  -webkit-backdrop-filter: blur(20px) saturate(160%);
  border: 1px solid var(--c-border);
  border-radius: var(--r-2xl);
  padding: clamp(0.85rem, 1.8vw, 1.15rem);
  display: flex;
  flex-direction: column;
  gap: clamp(0.7rem, 1.2vw, 0.9rem);
  box-shadow: var(--shadow-md), inset 0 0 1px rgba(255, 255, 255, 0.10);
}

/* Header */
.prep-header {
  display: flex; align-items: center; gap: 0.65rem;
  padding-bottom: 0.55rem;
  border-bottom: 1px solid var(--c-divider);
}
.prep-head-icon {
  background: rgba(124, 58, 237, 0.10);
  border: 1px solid rgba(124, 58, 237, 0.25);
  border-radius: var(--r-md);
  padding: 0.4rem;
  flex-shrink: 0;
}
.prep-head-text { flex: 1; min-width: 0; }
.prep-head-stats {
  color: var(--c-text-tertiary); font-size: 0.74rem; line-height: 1.35;
}
.prep-head-strong {
  color: white; font-weight: 700; font-size: 0.78rem;
}
.prep-head-stats-loading { opacity: 0.65; }

/* Sections */
.prep-section { display: flex; flex-direction: column; gap: 0.55rem; }
.prep-section-head {
  display: flex; align-items: baseline; justify-content: space-between; gap: 0.5rem;
}
.prep-section-title {
  color: var(--c-text-tertiary);
  font-size: 0.68rem; font-weight: 800;
  letter-spacing: 0.06em; text-transform: uppercase;
  margin: 0;
}

/* Counter */
.prep-counter { display: inline-flex; align-items: baseline; gap: 2px; font-family: 'Georgia', serif; }
.prep-counter-num { color: #86EFAC; font-size: 1.05rem; font-weight: 800; line-height: 1; }
.prep-counter-sep { color: var(--c-text-muted); margin: 0 2px; }
.prep-counter-tot { color: var(--c-text-tertiary); font-size: 0.85rem; }
.prep-counter-lbl {
  color: var(--c-text-muted); font-size: 0.62rem;
  text-transform: uppercase; letter-spacing: 0.06em;
  margin-left: 0.3rem; font-family: var(--font-sans); font-weight: 700;
}

/* Skeletons */
.prep-skel {
  display: flex; align-items: center; gap: 0.5rem;
  padding: 0.7rem 0.8rem;
  background: rgba(255, 255, 255, 0.025);
  border: 1px dashed var(--c-border);
  border-radius: var(--r-md);
  color: var(--c-text-tertiary);
  font-size: 0.74rem;
}
.prep-spin { animation: spin 1s linear infinite; }

/* Banners */
.prep-banner {
  display: flex; align-items: flex-start; gap: 0.45rem;
  padding: 0.6rem 0.7rem;
  border-radius: var(--r-md);
  font-size: 0.72rem; line-height: 1.4;
}
.prep-banner.is-warn { background: rgba(251, 191, 36, 0.06); border: 1px solid rgba(251, 191, 36, 0.18); color: #FCD34D; }
.prep-banner.is-error { background: rgba(239, 68, 68, 0.06); border: 1px solid rgba(239, 68, 68, 0.20); color: var(--c-error-300); }

/* Stats grid */
.prep-stats {
  display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.4rem;
}
@media (min-width: 520px) {
  .prep-stats { grid-template-columns: repeat(4, 1fr); }
}
.prep-stat {
  display: flex; align-items: center; gap: 0.4rem;
  padding: 0.45rem 0.55rem;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  min-width: 0;
}
.prep-stat-icon {
  width: 1.4rem; height: 1.4rem;
  border-radius: var(--r-sm, 6px);
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}
.prep-stat-val { font-size: 0.82rem; font-weight: 800; line-height: 1; }
.prep-stat-lbl { color: var(--c-text-muted); font-size: 0.62rem; line-height: 1.1; }

/* KPI inline row */
.prep-kpi-row {
  display: flex; flex-wrap: wrap; gap: 0.35rem; align-items: center;
}
.prep-kpi-chip {
  display: inline-flex; align-items: baseline; gap: 0.35rem;
  padding: 0.3rem 0.55rem;
  background: rgba(96, 165, 250, 0.07);
  border: 1px solid rgba(96, 165, 250, 0.18);
  border-radius: var(--r-pill);
  max-width: 16rem;
}
.prep-kpi-val {
  color: #93C5FD; font-weight: 800; font-size: 0.78rem;
  font-family: 'Georgia', serif; letter-spacing: -0.01em;
  flex-shrink: 0;
}
.prep-kpi-lbl {
  color: var(--c-text-secondary); font-size: 0.66rem;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  max-width: 11rem;
}
.prep-kpi-more {
  color: var(--c-text-muted); font-size: 0.68rem; font-weight: 700;
  padding: 0.3rem 0.5rem;
}

/* Warnings */
.prep-warnings { display: flex; flex-direction: column; gap: 0.3rem; }
.prep-warn {
  display: flex; align-items: flex-start; gap: 0.4rem;
  padding: 0.45rem 0.6rem;
  border-radius: var(--r-md);
  font-size: 0.7rem; line-height: 1.4;
}
.prep-warn.is-warn  { background: rgba(251, 191, 36, 0.06); border: 1px solid rgba(251, 191, 36, 0.18); color: #FCD34D; }
.prep-warn.is-error { background: rgba(239, 68, 68, 0.06);  border: 1px solid rgba(239, 68, 68, 0.20); color: var(--c-error-300); }
.prep-warn strong { font-weight: 700; }
.prep-warn-minor {
  color: var(--c-text-muted); font-size: 0.66rem;
  padding: 0.25rem 0.6rem;
  font-style: italic;
}

/* Slides list */
.prep-slides {
  display: flex; flex-direction: column; gap: 0.25rem;
  list-style: none; margin: 0; padding: 0;
  max-height: 50vh; overflow-y: auto;
}
.prep-slide {
  display: flex; align-items: center; gap: 0.5rem;
  width: 100%;
  padding: 0.45rem 0.6rem;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  text-align: left; color: white; cursor: pointer;
  transition: background var(--t-base) var(--ease-out), border-color var(--t-base) var(--ease-out);
  font-size: 0.78rem;
}
.prep-slide:hover:not(:disabled):not(.is-mandatory) {
  background: rgba(255, 255, 255, 0.05);
  border-color: var(--c-border-strong);
}
.prep-slide.is-active { border-color: rgba(74, 222, 128, 0.25); background: rgba(74, 222, 128, 0.04); }
.prep-slide.is-excluded { opacity: 0.5; background: rgba(255, 255, 255, 0.015); }
.prep-slide.is-mandatory { cursor: default; border-color: rgba(251, 191, 36, 0.30); background: rgba(251, 191, 36, 0.04); }

.prep-slide-check { flex-shrink: 0; display: flex; }
.prep-slide-num {
  color: var(--c-text-muted); font-size: 0.7rem; font-weight: 700;
  font-family: 'Georgia', serif; flex-shrink: 0; min-width: 1.5rem;
}
.prep-slide-type {
  font-size: 0.62rem; font-weight: 800;
  letter-spacing: 0.04em; text-transform: uppercase;
  flex-shrink: 0;
}
.prep-slide-title {
  flex: 1; min-width: 0;
  color: white; font-weight: 600; font-size: 0.78rem;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.prep-slide-pin {
  font-size: 0.56rem; font-weight: 700;
  padding: 1px 5px; border-radius: var(--r-pill);
  background: rgba(251, 191, 36, 0.18); color: #FBBF24;
  letter-spacing: 0.04em; text-transform: uppercase;
  flex-shrink: 0;
}
.prep-slide-toggle {
  display: inline-flex; align-items: center; justify-content: center;
  width: 1.1rem; height: 1.1rem;
  color: var(--c-text-muted);
  flex-shrink: 0;
  border-radius: var(--r-sm, 4px);
}
.prep-slide-toggle:hover { color: white; background: rgba(255, 255, 255, 0.06); }

.prep-slide-detail {
  margin: 0.2rem 0 0.4rem 0;
  padding: 0.55rem 0.7rem;
  background: rgba(255, 255, 255, 0.02);
  border-left: 2px solid rgba(96, 165, 250, 0.30);
  border-radius: 0 var(--r-md) var(--r-md) 0;
  display: flex; flex-direction: column; gap: 0.3rem;
}
.prep-slide-sub { color: var(--c-text-secondary); font-size: 0.74rem; line-height: 1.4; }
.prep-slide-narr { color: var(--c-text-tertiary); font-size: 0.7rem; line-height: 1.45; }
.prep-slide-source { color: var(--c-text-muted); font-size: 0.66rem; }
.prep-slide-source strong { color: var(--c-text-secondary); }

/* Pills resumen */
.prep-pills { display: flex; flex-wrap: wrap; gap: 0.3rem; }
.prep-pill {
  display: inline-flex; align-items: center; gap: 0.3rem;
  padding: 0.22rem 0.5rem;
  border-radius: var(--r-pill);
  font-size: 0.62rem; font-weight: 700;
}
.prep-pill-cache {
  background: rgba(134, 239, 172, 0.10);
  border: 1px solid rgba(134, 239, 172, 0.25);
  color: var(--c-success-400);
}

/* Refinar */
.prep-refine { gap: 0; }
.prep-refine-toggle {
  display: flex; align-items: center; justify-content: space-between;
  width: 100%;
  padding: 0.55rem 0.7rem;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  color: var(--c-text-secondary);
  text-align: left;
  transition: all var(--t-base) var(--ease-out);
  cursor: pointer;
}
.prep-refine-toggle:hover { background: rgba(255, 255, 255, 0.05); border-color: var(--c-border-strong); color: white; }
.prep-refine-left { display: inline-flex; align-items: center; gap: 0.4rem; }
.prep-refine-label { font-size: 0.74rem; font-weight: 700; }
.prep-mode-pill {
  display: inline-flex; align-items: center;
  padding: 0.18rem 0.55rem;
  border-radius: var(--r-pill);
  background: rgba(124, 58, 237, 0.15);
  border: 1px solid rgba(124, 58, 237, 0.30);
  color: #C4B5FD;
  font-size: 0.62rem; font-weight: 800;
  text-transform: uppercase; letter-spacing: 0.04em;
}

.prep-refine-body {
  display: flex; flex-direction: column; gap: 0.55rem;
  padding: 0.7rem 0.5rem 0.2rem 0.5rem;
}
.prep-prompt {
  width: 100%;
  background: rgba(0, 0, 0, 0.25);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  padding: 0.55rem 0.65rem;
  color: white;
  font-family: var(--font-sans);
  font-size: 0.78rem;
  line-height: 1.45;
  resize: vertical;
  min-height: 3rem;
}
.prep-prompt:focus { outline: none; border-color: var(--c-brand-blue-400); }
.prep-prompt::placeholder { color: var(--c-text-muted); }

.prep-suggestions { display: flex; flex-direction: column; gap: 0.35rem; }
.prep-sugg-label {
  display: inline-flex; align-items: center; gap: 0.3rem;
  color: var(--c-text-tertiary);
  font-size: 0.66rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.05em;
}
.prep-sugg-row { display: flex; flex-wrap: wrap; gap: 0.3rem; }
.prep-sugg {
  padding: 0.35rem 0.6rem;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid var(--c-border);
  border-radius: var(--r-pill);
  color: var(--c-text-secondary);
  font-size: 0.7rem; font-weight: 600;
  text-align: left;
  transition: all var(--t-base) var(--ease-out);
  cursor: pointer;
}
.prep-sugg:hover { background: rgba(255, 255, 255, 0.06); color: white; border-color: var(--c-border-strong); }
.prep-sugg.is-active {
  background: rgba(74, 222, 128, 0.10);
  border-color: rgba(74, 222, 128, 0.40);
  color: #86EFAC;
}

.prep-mode-row {
  display: flex; align-items: center; flex-wrap: wrap; gap: 0.5rem;
}
.prep-mode-label {
  color: var(--c-text-tertiary);
  font-size: 0.66rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.05em;
}
.prep-mode-options {
  display: inline-flex;
  background: rgba(0, 0, 0, 0.25);
  border: 1px solid var(--c-border);
  border-radius: var(--r-pill);
  padding: 2px;
  gap: 2px;
}
.prep-mode-opt {
  padding: 0.3rem 0.65rem;
  border-radius: var(--r-pill);
  background: transparent;
  color: var(--c-text-tertiary);
  font-size: 0.7rem; font-weight: 700;
  border: none;
  cursor: pointer;
  transition: all var(--t-base) var(--ease-out);
}
.prep-mode-opt:hover { color: white; }
.prep-mode-opt.is-active {
  background: linear-gradient(135deg, rgba(96, 165, 250, 0.25), rgba(124, 58, 237, 0.25));
  color: white;
  box-shadow: 0 1px 4px rgba(96, 165, 250, 0.20);
}

.prep-refine-hint {
  color: var(--c-text-muted); font-size: 0.68rem; line-height: 1.45;
  margin-top: 0.2rem;
}
.prep-link {
  background: none; border: none; padding: 0; cursor: pointer;
  color: var(--c-brand-blue-300);
  font-size: inherit; font-weight: 700;
  text-decoration: underline;
}
.prep-link:hover { color: var(--c-brand-blue-200); }

/* CTA */
.prep-cta {
  margin-top: 0.2rem;
  padding: 0.85rem 1rem;
  border: none; border-radius: var(--r-lg);
  display: flex; align-items: center; justify-content: center; gap: 0.45rem;
  font-weight: 800; font-size: 0.86rem; color: white;
  background: linear-gradient(135deg, #1D4ED8 0%, #2563EB 50%, #4F46E5 100%);
  box-shadow: 0 8px 24px rgba(37, 99, 235, 0.32),
              inset 0 0 0 1px rgba(147, 197, 253, 0.16);
  cursor: pointer;
  transition: transform var(--t-base) var(--ease-out);
}
.prep-cta:hover:not(:disabled) { transform: translateY(-1px); }
.prep-cta:disabled { opacity: 0.5; cursor: not-allowed; box-shadow: none; }
`;
