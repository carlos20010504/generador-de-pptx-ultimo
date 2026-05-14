"use client";

import React, { useEffect, useMemo, useState } from 'react';
import {
  Sparkles, FileSpreadsheet, AlertTriangle, AlertCircle,
  TrendingUp, Layers, BarChart2, Loader2, ChevronDown, ChevronRight,
  Lightbulb, CheckSquare, Square, Table2, FileText, LayoutDashboard,
  GripVertical, Pencil, Check as CheckIcon,
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
  // PII detected at parse time (emails, doc IDs, phones, etc.). Already
  // exposed by the warnings list; surfaced separately here so we can give
  // it dedicated UX (icon + count + per-column list).
  pii_findings?: Array<{ sheet: string; column: string; kind: string; label: string }>;
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

interface IntentReportSheet {
  requested: string;
  matched: string | null;
  closest: string | null;
  slide_indices: number[];
}

interface IntentReport {
  requested_slide_count: number | null;
  actual_slide_count: number;
  count_honored: boolean;
  required_sheets: IntentReportSheet[];
  skipped_sheets: string[];
}

interface PreviewResponse {
  meta: { title?: string; subtitle?: string };
  slides: SlidePreview[];
  ai_status?: { model?: string; cache_hit?: boolean };
  intent_report?: IntentReport | null;
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
  onThemeChange?: (t: ThemeShape) => void;
  mode: Mode;
  onModeChange: (m: Mode) => void;
  // Returns excluded indices PLUS user overrides applied in the panel:
  //   - titles: { originalIndex: newTitle } — only changed titles
  //   - order:  number[] of originalIndex in the new order; null when
  //             the user didn't reorder (preserves current behavior).
  onConfirm: (
    excludedIndices: number[],
    overrides: { titles?: Record<number, string>; order?: number[] | null }
  ) => void;
  onOpenAdvanced?: () => void;
}

// Themes para el PPT generado. Sincronizado con AIControlPanel.THEME_OPTIONS.
const PREP_THEMES: ThemeShape[] = [
  { key: 'socya-institucional', name: 'Institucional', primary_hex: '#087062', accent_hex: '#69BE28', text_hex: '#1A1A1A', bg_hex: '#FFFFFF' },
  { key: 'socya-comite',        name: 'Comité',        primary_hex: '#123C49', accent_hex: '#F3C400', text_hex: '#1A1A1A', bg_hex: '#FFFFFF' },
  { key: 'socya-impacto',       name: 'Impacto',       primary_hex: '#087062', accent_hex: '#FF8300', text_hex: '#1A1A1A', bg_hex: '#FFFFFF' },
  { key: 'analitica-moderna',   name: 'Analítica',     primary_hex: '#0F172A', accent_hex: '#2563EB', text_hex: '#1A1A1A', bg_hex: '#F8FAFC' },
];

const SLIDE_TYPE_META: Record<string, { Ic: React.ComponentType<{ size?: number; color?: string }>; label: string; color: string }> = {
  title:        { Ic: FileSpreadsheet, label: 'Portada',   color: '#F3C400' }, // yellow
  kpi_row:      { Ic: LayoutDashboard, label: 'KPIs',      color: '#087062' }, // primary teal
  chart:        { Ic: BarChart2,       label: 'Gráfico',   color: '#00A0DF' }, // sky
  table:        { Ic: Table2,          label: 'Tabla',     color: '#69BE28' }, // logo-green
  text_bullets: { Ic: FileText,        label: 'Hallazgos', color: '#FF8300' }, // orange
  unknown:      { Ic: FileText,        label: 'Slide',     color: '#4D4F53' }, // logo-gray
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
  audience, language, theme, onThemeChange, mode, onModeChange,
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
  const [refineOpen, setRefineOpen] = useState(true);  // abierto por default — único lugar para el prompt

  // Cheap stable identity — depending on the File object itself causes
  // re-fetches on every parent re-render which abort in-flight requests.
  const fileKey = `${file.name}|${file.size}|${file.lastModified}`;
  const promptKey = `${fileKey}|${userPrompt}`;

  // Edición inline de títulos. Map<originalIndex, nuevoTitulo>. Sólo
  // entradas presentes son overrides; las demás conservan su valor del plan.
  const [editedTitles, setEditedTitles] = useState<Record<number, string>>({});
  const [editingIdx, setEditingIdx] = useState<number | null>(null);
  const [editDraft, setEditDraft] = useState<string>('');

  // Reordenamiento por drag&drop. `customOrder` es array de originalIndex
  // en el orden que quiere el usuario. null = orden original del plan.
  const [customOrder, setCustomOrder] = useState<number[] | null>(null);
  const [dragIdx, setDragIdx] = useState<number | null>(null);   // pos en customOrder mientras drag
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null);

  // Cuando el plan llega/cambia, resetea overrides para evitar referencias
  // a slides que ya no existen.
  useEffect(() => {
    setEditedTitles({});
    setCustomOrder(null);
    setExcluded(new Set());
    setEditingIdx(null);
    setExpandedSlide(null);
  }, [plan?.slides?.length, promptKey]);

  // Lista efectiva en pantalla: aplica customOrder si existe.
  const orderedIndices = useMemo(() => {
    const total = plan?.slides?.length ?? 0;
    if (customOrder && customOrder.length === total) return customOrder;
    return Array.from({ length: total }, (_, i) => i);
  }, [customOrder, plan?.slides?.length]);

  // Timeouts client-side: deben ser >= al budget del backend, sino abortamos
  // requests que iban a responder bien. Backend actual:
  //   - quick-summary maxDuration = 60s (determinístico pero Python tiene
  //     cold-start lento en Windows con archivos grandes).
  //   - preview-plan maxDuration = 300s, AIChain.PATIENT.total_budget = 200s.
  // Damos buffer encima para que el backend gane la carrera de abortar
  // primero (con mensaje útil) en vez del cliente — antes ambos timeouts
  // estaban POR DEBAJO del budget, abortando requests que iban a responder.
  const QUICK_SUMMARY_TIMEOUT_MS = 70_000;       // 60s backend + 10s buffer
  const PREVIEW_PLAN_TIMEOUT_MS = 240_000;       // 200s budget + 40s buffer

  // Quick-summary: refetches only when file changes.
  useEffect(() => {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), QUICK_SUMMARY_TIMEOUT_MS);
    let cancelled = false;
    (async () => {
      setSummaryLoading(true);
      setSummaryError(null);
      try {
        const fd = new FormData();
        fd.append('file', file);
        const res = await fetch('/api/quick-summary', {
          method: 'POST', body: fd, signal: ctrl.signal,
        });
        if (cancelled) return;
        if (!res.ok) {
          const j = await res.json().catch(() => null);
          const msg = typeof j?.error === 'string' ? j.error : (j?.error?.message || `Error ${res.status}`);
          throw new Error(msg);
        }
        const data: QuickSummary = await res.json();
        if (!cancelled) setSummary(data);
      } catch (err: unknown) {
        if (cancelled) return;
        const isAbort = err instanceof DOMException && err.name === 'AbortError';
        setSummaryError(isAbort
          ? 'El análisis del Excel tardó más de 70 segundos. Esto suele pasar con archivos muy grandes o con muchas hojas. Probá con uno más pequeño.'
          : (err as Error).message || 'No se pudo analizar el archivo.');
      } finally {
        clearTimeout(timer);
        if (!cancelled) setSummaryLoading(false);
      }
    })();
    return () => { cancelled = true; clearTimeout(timer); ctrl.abort(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileKey]);

  // Preview-plan: refetches when file or prompt change. Cache hit on backend
  // for same (file, prompt) — safe to re-run without quota cost.
  useEffect(() => {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), PREVIEW_PLAN_TIMEOUT_MS);
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
        const res = await fetch('/api/preview-plan', {
          method: 'POST', body: fd, signal: ctrl.signal,
        });
        if (cancelled) return;
        if (!res.ok) {
          const j = await res.json().catch(() => null);
          const msg = typeof j?.error === 'string' ? j.error : (j?.error?.message || `Error ${res.status}`);
          throw new Error(msg);
        }
        const data: PreviewResponse = await res.json();
        if (!cancelled) setPlan(data);
      } catch (err: unknown) {
        if (cancelled) return;
        const isAbort = err instanceof DOMException && err.name === 'AbortError';
        setPlanError(isAbort
          ? 'La IA está saturada (todos los modelos free tier tardaron más de 4 minutos en responder). Esperá un par de minutos y volvé a intentar.'
          : (err as Error).message || 'No se pudo generar el plan.');
      } finally {
        clearTimeout(timer);
        if (!cancelled) setPlanLoading(false);
      }
    })();
    return () => { cancelled = true; clearTimeout(timer); ctrl.abort(); };
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

  // ── Inline edit ──────────────────────────────────────────────────
  const startEdit = (originalIdx: number, currentTitle: string) => {
    setEditingIdx(originalIdx);
    setEditDraft(editedTitles[originalIdx] ?? currentTitle ?? '');
  };
  const commitEdit = () => {
    if (editingIdx === null) return;
    const trimmed = editDraft.trim();
    setEditedTitles(prev => {
      const next = { ...prev };
      if (!trimmed) {
        delete next[editingIdx];   // empty → revert al original
      } else {
        next[editingIdx] = trimmed;
      }
      return next;
    });
    setEditingIdx(null);
    setEditDraft('');
  };
  const cancelEdit = () => {
    setEditingIdx(null);
    setEditDraft('');
  };

  // ── Drag & drop reorder ──────────────────────────────────────────
  // `dragIdx` y `dragOverIdx` son posiciones DENTRO de orderedIndices.
  const onDragStart = (posInOrdered: number) => (e: React.DragEvent) => {
    setDragIdx(posInOrdered);
    e.dataTransfer.effectAllowed = 'move';
    // Firefox necesita data válida para iniciar el drag
    try { e.dataTransfer.setData('text/plain', String(posInOrdered)); } catch { /* ignore */ }
  };
  const onDragOver = (posInOrdered: number) => (e: React.DragEvent) => {
    e.preventDefault();
    if (dragIdx === null || dragIdx === posInOrdered) return;
    setDragOverIdx(posInOrdered);
  };
  const onDrop = (posInOrdered: number) => (e: React.DragEvent) => {
    e.preventDefault();
    if (dragIdx === null || dragIdx === posInOrdered) {
      setDragIdx(null); setDragOverIdx(null);
      return;
    }
    const source = dragIdx;
    const target = posInOrdered;
    setCustomOrder(prev => {
      const base = prev && prev.length === orderedIndices.length
        ? [...prev]
        : [...orderedIndices];
      const [moved] = base.splice(source, 1);
      base.splice(target, 0, moved);
      return base;
    });
    setDragIdx(null);
    setDragOverIdx(null);
  };
  const onDragEnd = () => {
    setDragIdx(null);
    setDragOverIdx(null);
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
          <Sparkles size={16} color="#087062" />
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
              <Stat icon={TrendingUp} label="KPIs" value={summary.totals.kpi_candidates} color="#087062" />
              <Stat icon={BarChart2} label="Distribuciones" value={summary.totals.categorical_distributions} color="#00A0DF" />
              <Stat icon={Layers} label="Series" value={summary.totals.time_series} color="#69BE28" />
              <Stat
                icon={Sparkles}
                label="Slides est."
                value={`${summary.deck_estimate.min_slides}–${summary.deck_estimate.max_slides}`}
                color="#F3C400"
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
           Section 2 — Tu instrucción para la IA (único lugar)
         ──────────────────────────────────────────────── */}
      <section className="prep-section prep-refine">
        <button
          type="button"
          onClick={() => setRefineOpen(o => !o)}
          className="prep-refine-toggle"
        >
          <span className="prep-refine-left">
            {refineOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            <span className="prep-refine-label">Tu instrucción para la IA</span>
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
              placeholder="Ej: hazme 9 slides incluyendo Riesgos Core y Riesgos Acciones"
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

            {onThemeChange && (
              <div className="prep-mode-row">
                <label className="prep-mode-label">Tema visual</label>
                <div className="prep-theme-options">
                  {PREP_THEMES.map((t) => (
                    <button
                      key={t.key}
                      type="button"
                      onClick={() => onThemeChange(t)}
                      className={`prep-theme-opt ${theme.key === t.key ? 'is-active' : ''}`}
                      title={`${t.name}: primario ${t.primary_hex}, acento ${t.accent_hex}`}
                    >
                      <span className="prep-theme-swatch" aria-hidden>
                        <span style={{ background: t.primary_hex }} />
                        <span style={{ background: t.accent_hex }} />
                      </span>
                      {t.name}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {onOpenAdvanced && (
              <p className="prep-refine-hint">
                ¿Necesitas cambiar audiencia, idioma o tema visual? Abre{' '}
                <button type="button" className="prep-link" onClick={onOpenAdvanced}>Audiencia y tema</button>.
              </p>
            )}
          </div>
        )}
      </section>

      {/* ────────────────────────────────────────────────
           Section 3 — Plan de slides
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
          <SectionSkeleton label="Construyendo el plan con IA — puede tardar hasta 4 min en hora pico (free tier). El resultado se cachea, así que reintentos son instantáneos." />
        ) : planError ? (
          <div className="prep-banner is-error">
            <AlertCircle size={14} />
            <span>{planError}</span>
          </div>
        ) : plan ? (
          <>
            {plan.intent_report && (
              <IntentBanner report={plan.intent_report} />
            )}
            <p className="prep-slide-hint">
              Click para activar/desactivar · arrastra <GripVertical size={11} style={{ verticalAlign: 'middle' }} /> para reordenar · usa <Pencil size={11} style={{ verticalAlign: 'middle' }} /> para editar el título.
            </p>
            <ul className="prep-slides">
              {orderedIndices.map((origIdx, posInOrdered) => {
                const slide = plan.slides[origIdx];
                if (!slide) return null;
                const meta = SLIDE_TYPE_META[slide.type] || SLIDE_TYPE_META.unknown;
                const isExcluded = excluded.has(slide.index);
                const isMandatory = !!slide.mandatory;
                const isExpanded = expandedSlide === slide.index;
                const isEditing = editingIdx === slide.index;
                const displayTitle = editedTitles[slide.index] ?? slide.title;
                const isDragOver = dragOverIdx === posInOrdered;
                const isDragging = dragIdx === posInOrdered;
                return (
                  <li key={slide.index}>
                    <div
                      className={[
                        'prep-slide',
                        isExcluded ? 'is-excluded' : 'is-active',
                        isMandatory ? 'is-mandatory' : '',
                        isDragging ? 'is-dragging' : '',
                        isDragOver ? 'is-drag-over' : '',
                      ].filter(Boolean).join(' ')}
                      draggable={!isMandatory && !isEditing}
                      onDragStart={onDragStart(posInOrdered)}
                      onDragOver={onDragOver(posInOrdered)}
                      onDrop={onDrop(posInOrdered)}
                      onDragEnd={onDragEnd}
                      onDragLeave={() => setDragOverIdx(null)}
                    >
                      {/* Drag handle (no actúa como toggle) */}
                      <span
                        className={`prep-slide-drag ${isMandatory ? 'is-disabled' : ''}`}
                        title={isMandatory ? 'Slide obligatoria — no se puede mover' : 'Arrastra para reordenar'}
                        aria-label="Mover slide"
                      >
                        <GripVertical size={13} />
                      </span>

                      {/* Toggle activar/desactivar */}
                      <button
                        type="button"
                        onClick={() => toggle(slide.index, isMandatory)}
                        disabled={isMandatory}
                        className="prep-slide-check-btn"
                        aria-label={isExcluded ? 'Activar slide' : 'Desactivar slide'}
                      >
                        {isExcluded
                          ? <Square size={15} color="rgba(26,26,26,0.30)" />
                          : <CheckSquare size={15} color={isMandatory ? '#F3C400' : '#69BE28'} />}
                      </button>

                      <span className="prep-slide-num">{String(posInOrdered + 1).padStart(2, '0')}</span>
                      <meta.Ic size={12} color={meta.color} />
                      <span className="prep-slide-type" style={{ color: meta.color }}>{meta.label}</span>

                      {/* Title — inline editable */}
                      {isEditing ? (
                        <span className="prep-slide-edit">
                          <input
                            autoFocus
                            type="text"
                            value={editDraft}
                            onChange={(e) => setEditDraft(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') { e.preventDefault(); commitEdit(); }
                              if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
                            }}
                            onBlur={commitEdit}
                            className="prep-slide-input"
                            placeholder="Nuevo título"
                            maxLength={140}
                          />
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={commitEdit}
                            className="prep-slide-edit-ok"
                            aria-label="Guardar"
                          >
                            <CheckIcon size={11} />
                          </button>
                        </span>
                      ) : (
                        <button
                          type="button"
                          className="prep-slide-title-btn"
                          onClick={() => startEdit(slide.index, displayTitle || meta.label)}
                          title="Click para editar"
                        >
                          <span className="prep-slide-title">
                            {displayTitle || `${meta.label} sin título`}
                          </span>
                          {editedTitles[slide.index] !== undefined && (
                            <span className="prep-slide-edited" title="Título editado">·</span>
                          )}
                          <Pencil size={10} className="prep-slide-edit-icon" />
                        </button>
                      )}

                      {isMandatory && <span className="prep-slide-pin">obligatoria</span>}

                      {/* Expand chevron */}
                      <button
                        type="button"
                        className="prep-slide-toggle"
                        onClick={() => setExpandedSlide(isExpanded ? null : slide.index)}
                        aria-label={isExpanded ? 'Ocultar detalles' : 'Ver detalles'}
                      >
                        {isExpanded ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                      </button>
                    </div>
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
           CTA
         ──────────────────────────────────────────────── */}
      <button
        type="button"
        onClick={() => onConfirm(Array.from(excluded), {
          // Sólo enviamos overrides cuando hay algo realmente cambiado.
          titles: Object.keys(editedTitles).length > 0 ? editedTitles : undefined,
          order: customOrder && customOrder.length === (plan?.slides?.length ?? 0)
            ? customOrder
            : null,
        })}
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

function IntentBanner({ report }: { report: IntentReport }) {
  const matchedSheets = report.required_sheets.filter(s => s.matched);
  const unmatchedSheets = report.required_sheets.filter(s => !s.matched);
  const countLine = report.requested_slide_count !== null
    ? (report.count_honored
        ? `${report.actual_slide_count} slides (como pediste)`
        : `${report.actual_slide_count} slides — pediste ${report.requested_slide_count} pero el Excel sólo tiene material para ${report.actual_slide_count}`)
    : null;
  const allHonored = (report.count_honored
                       && unmatchedSheets.length === 0
                       && report.skipped_sheets.length === 0);

  if (!countLine && matchedSheets.length === 0 && unmatchedSheets.length === 0) {
    return null;
  }

  return (
    <div className={`prep-intent ${allHonored ? 'is-ok' : 'is-warn'}`}>
      <div className="prep-intent-icon" aria-hidden>
        {allHonored ? <CheckIcon size={13} /> : <AlertTriangle size={13} />}
      </div>
      <div className="prep-intent-body">
        <p className="prep-intent-title">
          {allHonored ? 'Detecté en tu instrucción:' : 'Detecté parcialmente en tu instrucción:'}
        </p>
        <ul className="prep-intent-list">
          {countLine && <li>{countLine}</li>}
          {matchedSheets.map((s, i) => (
            <li key={`m-${i}`}>
              <strong>{s.matched}</strong>
              {s.slide_indices.length
                ? ` — slide${s.slide_indices.length > 1 ? 's' : ''} ${s.slide_indices.map(n => n + 1).join(', ')}`
                : ' — incluida en el plan'}
            </li>
          ))}
          {unmatchedSheets.map((s, i) => (
            <li key={`u-${i}`} className="prep-intent-warn">
              <em>"{s.requested}"</em> no se encontró
              {s.closest ? <> — la hoja más parecida es <strong>{s.closest}</strong></> : null}
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}

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
  background: var(--c-bg-elevated);
  border: 1px solid var(--c-border);
  border-radius: var(--r-lg);
  padding: clamp(0.95rem, 1.8vw, 1.25rem);
  display: flex;
  flex-direction: column;
  gap: clamp(0.75rem, 1.3vw, 0.95rem);
  box-shadow: var(--shadow-md);
}

/* ── Header (compact stats) ── */
.prep-header {
  display: flex; align-items: center; gap: 0.7rem;
  padding-bottom: 0.65rem;
  border-bottom: 1px solid var(--c-divider);
}
.prep-head-icon {
  background: var(--c-accent-green);
  border: 1px solid rgba(8, 112, 98, 0.20);
  border-radius: var(--r-md);
  padding: 0.4rem;
  flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
}
.prep-head-text { flex: 1; min-width: 0; }
.prep-head-stats {
  color: var(--c-text-secondary);
  font-size: 0.76rem; line-height: 1.4;
}
.prep-head-strong {
  color: var(--c-text-primary);
  font-family: var(--font-heading);
  font-weight: 700;
  font-size: 0.8rem;
}
.prep-head-stats-loading { opacity: 0.65; }

/* ── Sections ── */
.prep-section { display: flex; flex-direction: column; gap: 0.55rem; }
.prep-section-head {
  display: flex; align-items: baseline; justify-content: space-between; gap: 0.5rem;
}
.prep-section-title {
  color: var(--c-text-secondary);
  font-family: var(--font-heading);
  font-size: 0.7rem; font-weight: 800;
  letter-spacing: 0.07em; text-transform: uppercase;
  margin: 0;
}

/* ── Counter ── */
.prep-counter {
  display: inline-flex; align-items: baseline; gap: 2px;
  font-family: var(--font-heading);
}
.prep-counter-num {
  color: var(--c-primary);
  font-size: 1.15rem; font-weight: 800; line-height: 1;
}
.prep-counter-sep { color: var(--c-text-muted); margin: 0 2px; }
.prep-counter-tot { color: var(--c-text-tertiary); font-size: 0.88rem; }
.prep-counter-lbl {
  color: var(--c-text-muted); font-size: 0.62rem;
  text-transform: uppercase; letter-spacing: 0.06em;
  margin-left: 0.4rem; font-weight: 700;
}

/* ── Skeletons ── */
.prep-skel {
  display: flex; align-items: center; gap: 0.55rem;
  padding: 0.75rem 0.85rem;
  background: var(--c-bg-tinted);
  border: 1px dashed var(--c-border);
  border-radius: var(--r-md);
  color: var(--c-text-secondary);
  font-size: 0.76rem;
}
.prep-spin { animation: spin 1s linear infinite; }

/* ── Banners ── */
.prep-banner {
  display: flex; align-items: flex-start; gap: 0.5rem;
  padding: 0.65rem 0.8rem;
  border-radius: var(--r-md);
  font-size: 0.74rem; line-height: 1.45;
}
.prep-banner.is-warn {
  background: #FEF7E0;
  border: 1px solid rgba(243, 196, 0, 0.40);
  color: #8A6800;
}
.prep-banner.is-error {
  background: #FEF2F2;
  border: 1px solid rgba(212, 56, 56, 0.30);
  color: var(--c-error-300);
}

/* ── Stats grid ── */
.prep-stats {
  display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.45rem;
}
@media (min-width: 520px) {
  .prep-stats { grid-template-columns: repeat(4, 1fr); }
}
.prep-stat {
  display: flex; align-items: center; gap: 0.45rem;
  padding: 0.55rem 0.65rem;
  background: var(--c-bg-elevated);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  min-width: 0;
  transition: border-color var(--t-base) var(--ease-out);
}
.prep-stat:hover { border-color: var(--c-border-strong); }
.prep-stat-icon {
  width: 1.55rem; height: 1.55rem;
  border-radius: var(--r-sm);
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}
.prep-stat-val {
  font-family: var(--font-heading);
  font-size: 0.92rem; font-weight: 800; line-height: 1;
}
.prep-stat-lbl {
  color: var(--c-text-tertiary);
  font-size: 0.64rem; line-height: 1.15;
}

/* ── KPI inline row ── */
.prep-kpi-row {
  display: flex; flex-wrap: wrap; gap: 0.4rem; align-items: center;
}
.prep-kpi-chip {
  display: inline-flex; align-items: baseline; gap: 0.4rem;
  padding: 0.35rem 0.65rem;
  background: var(--c-bg-tinted);
  border: 1px solid rgba(8, 112, 98, 0.22);
  border-radius: var(--r-pill);
  max-width: 16rem;
}
.prep-kpi-val {
  color: var(--c-primary);
  font-family: var(--font-heading);
  font-weight: 800; font-size: 0.82rem;
  letter-spacing: -0.01em;
  flex-shrink: 0;
}
.prep-kpi-lbl {
  color: var(--c-text-secondary); font-size: 0.68rem;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  max-width: 11rem;
}
.prep-kpi-more {
  color: var(--c-text-tertiary); font-size: 0.7rem; font-weight: 700;
  padding: 0.3rem 0.5rem;
}

/* ── Warnings ── */
.prep-warnings { display: flex; flex-direction: column; gap: 0.35rem; }
.prep-warn {
  display: flex; align-items: flex-start; gap: 0.45rem;
  padding: 0.5rem 0.7rem;
  border-radius: var(--r-md);
  font-size: 0.72rem; line-height: 1.45;
}
.prep-warn.is-warn  {
  background: #FEF7E0;
  border: 1px solid rgba(243, 196, 0, 0.40);
  color: #8A6800;
}
.prep-warn.is-error {
  background: #FEF2F2;
  border: 1px solid rgba(212, 56, 56, 0.30);
  color: var(--c-error-300);
}
.prep-warn strong { font-weight: 700; }
.prep-warn-minor {
  color: var(--c-text-muted); font-size: 0.68rem;
  padding: 0.25rem 0.6rem;
  font-style: italic;
}

/* ── Slides list ── */
.prep-slides {
  display: flex; flex-direction: column; gap: 0.3rem;
  list-style: none; margin: 0; padding: 0;
  max-height: 50vh; overflow-y: auto;
}
.prep-slide-hint {
  color: var(--c-text-tertiary);
  font-size: 0.7rem;
  line-height: 1.5;
  padding: 0.1rem 0.2rem;
}

.prep-slide {
  display: flex; align-items: center; gap: 0.55rem;
  width: 100%;
  padding: 0.55rem 0.7rem;
  background: var(--c-bg-elevated);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  text-align: left;
  color: var(--c-text-primary);
  transition: background var(--t-base) var(--ease-out), border-color var(--t-base) var(--ease-out), opacity var(--t-base) var(--ease-out);
  font-size: 0.78rem;
}
.prep-slide.is-dragging { opacity: 0.5; }
.prep-slide.is-drag-over {
  border-color: var(--c-primary);
  border-style: dashed;
  background: var(--c-bg-tinted);
}

/* Drag handle */
.prep-slide-drag {
  display: inline-flex; align-items: center; justify-content: center;
  width: 1.2rem; height: 1.2rem;
  color: var(--c-text-muted);
  cursor: grab;
  flex-shrink: 0;
  border-radius: var(--r-sm);
}
.prep-slide-drag:hover { color: var(--c-primary); background: var(--c-bg-tinted); }
.prep-slide-drag:active { cursor: grabbing; }
.prep-slide-drag.is-disabled { cursor: not-allowed; opacity: 0.35; }

/* Toggle activate/deactivate */
.prep-slide-check-btn {
  background: none; border: none; padding: 0;
  display: inline-flex; align-items: center;
  cursor: pointer; flex-shrink: 0;
}
.prep-slide-check-btn:disabled { cursor: default; }

/* Title button (click to edit) */
.prep-slide-title-btn {
  flex: 1; min-width: 0;
  background: none; border: none; padding: 0.2rem 0.35rem;
  border-radius: var(--r-sm);
  display: inline-flex; align-items: center; gap: 0.35rem;
  color: var(--c-text-primary);
  text-align: left;
  cursor: text;
  transition: background var(--t-base) var(--ease-out);
}
.prep-slide-title-btn:hover {
  background: var(--c-bg-tinted);
}
.prep-slide-title-btn:hover .prep-slide-edit-icon {
  opacity: 1;
}
.prep-slide-edit-icon {
  color: var(--c-text-muted);
  opacity: 0;
  flex-shrink: 0;
  transition: opacity var(--t-base) var(--ease-out);
}
.prep-slide-edited {
  color: var(--c-primary);
  font-weight: 800;
  font-size: 1rem;
  line-height: 0;
  margin: 0 -0.1rem;
}

/* Inline edit input */
.prep-slide-edit {
  flex: 1; min-width: 0;
  display: inline-flex; align-items: center; gap: 0.4rem;
}
.prep-slide-input {
  flex: 1; min-width: 0;
  font-family: var(--font-sans);
  font-size: 0.8rem;
  font-weight: 600;
  color: var(--c-text-primary);
  background: var(--c-bg-elevated);
  border: 1.5px solid var(--c-primary);
  border-radius: var(--r-sm);
  padding: 0.25rem 0.45rem;
  outline: none;
  box-shadow: 0 0 0 3px rgba(8, 112, 98, 0.15);
}
.prep-slide-edit-ok {
  display: inline-flex; align-items: center; justify-content: center;
  width: 1.55rem; height: 1.55rem;
  background: var(--c-primary);
  border: none;
  border-radius: var(--r-sm);
  color: white;
  cursor: pointer;
  flex-shrink: 0;
}
.prep-slide-edit-ok:hover { background: var(--c-primary-dark); }
.prep-slide:not(.is-mandatory):hover {
  background: var(--c-bg-tinted);
  border-color: var(--c-primary);
}
.prep-slide.is-active {
  border-color: rgba(105, 190, 40, 0.45);
  background: rgba(105, 190, 40, 0.05);
}
.prep-slide.is-excluded {
  opacity: 0.55;
  background: var(--c-bg-neutral);
}
.prep-slide.is-mandatory {
  cursor: default;
  border-color: rgba(243, 196, 0, 0.45);
  background: rgba(243, 196, 0, 0.06);
}

.prep-slide-check { flex-shrink: 0; display: flex; }
.prep-slide-num {
  color: var(--c-text-tertiary); font-size: 0.7rem; font-weight: 700;
  font-family: var(--font-heading); flex-shrink: 0; min-width: 1.5rem;
}
.prep-slide-type {
  font-family: var(--font-heading);
  font-size: 0.62rem; font-weight: 800;
  letter-spacing: 0.05em; text-transform: uppercase;
  flex-shrink: 0;
}
.prep-slide-title {
  flex: 1; min-width: 0;
  color: var(--c-text-primary);
  font-weight: 600; font-size: 0.8rem;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.prep-slide-pin {
  font-size: 0.56rem; font-weight: 700;
  padding: 1px 6px; border-radius: var(--r-pill);
  background: var(--c-yellow);
  color: #5C4A00;
  letter-spacing: 0.05em; text-transform: uppercase;
  flex-shrink: 0;
}
.prep-slide-toggle {
  display: inline-flex; align-items: center; justify-content: center;
  width: 1.15rem; height: 1.15rem;
  color: var(--c-text-muted);
  flex-shrink: 0;
  border-radius: var(--r-sm);
}
.prep-slide-toggle:hover {
  color: var(--c-primary);
  background: var(--c-bg-tinted);
}

.prep-slide-detail {
  margin: 0.25rem 0 0.45rem 0;
  padding: 0.65rem 0.8rem;
  background: var(--c-bg-tinted);
  border-left: 3px solid var(--c-primary);
  border-radius: 0 var(--r-md) var(--r-md) 0;
  display: flex; flex-direction: column; gap: 0.35rem;
}
.prep-slide-sub { color: var(--c-text-primary); font-size: 0.76rem; line-height: 1.45; font-weight: 600; }
.prep-slide-narr { color: var(--c-text-secondary); font-size: 0.72rem; line-height: 1.5; }
.prep-slide-source { color: var(--c-text-tertiary); font-size: 0.68rem; }
.prep-slide-source strong { color: var(--c-primary); font-weight: 700; }

/* ── Pills resumen ── */
.prep-pills { display: flex; flex-wrap: wrap; gap: 0.35rem; }
.prep-pill {
  display: inline-flex; align-items: center; gap: 0.3rem;
  padding: 0.25rem 0.55rem;
  border-radius: var(--r-pill);
  font-family: var(--font-heading);
  font-size: 0.62rem; font-weight: 700;
  letter-spacing: 0.04em; text-transform: uppercase;
}
.prep-pill-cache {
  background: var(--c-accent-green);
  border: 1px solid rgba(105, 190, 40, 0.40);
  color: var(--c-primary-dark);
}

/* ── Refinar ── */
.prep-refine { gap: 0; }
.prep-refine-toggle {
  display: flex; align-items: center; justify-content: space-between;
  width: 100%;
  padding: 0.65rem 0.85rem;
  background: var(--c-bg-tinted);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  color: var(--c-text-primary);
  text-align: left;
  transition: all var(--t-base) var(--ease-out);
  cursor: pointer;
}
.prep-refine-toggle:hover {
  background: var(--c-accent-green);
  border-color: var(--c-primary);
}
.prep-refine-left { display: inline-flex; align-items: center; gap: 0.45rem; }
.prep-refine-label {
  font-family: var(--font-heading);
  font-size: 0.76rem; font-weight: 700;
}
.prep-mode-pill {
  display: inline-flex; align-items: center;
  padding: 0.22rem 0.6rem;
  border-radius: var(--r-pill);
  background: var(--c-primary);
  border: 1px solid var(--c-primary);
  color: white;
  font-family: var(--font-heading);
  font-size: 0.62rem; font-weight: 800;
  text-transform: uppercase; letter-spacing: 0.05em;
}

.prep-refine-body {
  display: flex; flex-direction: column; gap: 0.65rem;
  padding: 0.8rem 0.5rem 0.2rem 0.5rem;
}
.prep-prompt {
  width: 100%;
  background: var(--c-bg-elevated);
  border: 1px solid var(--c-border-strong);
  border-radius: var(--r-md);
  padding: 0.6rem 0.75rem;
  color: var(--c-text-primary);
  font-family: var(--font-sans);
  font-size: 0.8rem;
  line-height: 1.5;
  resize: vertical;
  min-height: 3rem;
}
.prep-prompt:focus {
  outline: none;
  border-color: var(--c-primary);
  box-shadow: 0 0 0 3px rgba(8, 112, 98, 0.15);
}
.prep-prompt::placeholder { color: var(--c-text-muted); }

.prep-suggestions { display: flex; flex-direction: column; gap: 0.4rem; }
.prep-sugg-label {
  display: inline-flex; align-items: center; gap: 0.35rem;
  color: var(--c-text-secondary);
  font-family: var(--font-heading);
  font-size: 0.66rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
}
.prep-sugg-row { display: flex; flex-wrap: wrap; gap: 0.35rem; }
.prep-sugg {
  padding: 0.4rem 0.7rem;
  background: var(--c-bg-elevated);
  border: 1px solid var(--c-border);
  border-radius: var(--r-pill);
  color: var(--c-text-secondary);
  font-size: 0.72rem; font-weight: 600;
  text-align: left;
  transition: all var(--t-base) var(--ease-out);
  cursor: pointer;
}
.prep-sugg:hover {
  background: var(--c-bg-tinted);
  color: var(--c-primary);
  border-color: var(--c-primary);
}
.prep-sugg.is-active {
  background: var(--c-primary);
  border-color: var(--c-primary);
  color: white;
}

.prep-mode-row {
  display: flex; align-items: center; flex-wrap: wrap; gap: 0.55rem;
}
.prep-mode-label {
  color: var(--c-text-secondary);
  font-family: var(--font-heading);
  font-size: 0.66rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
}
.prep-mode-options {
  display: inline-flex;
  background: var(--c-bg-tinted);
  border: 1px solid var(--c-border);
  border-radius: var(--r-pill);
  padding: 3px;
  gap: 2px;
}
.prep-mode-opt {
  padding: 0.35rem 0.75rem;
  border-radius: var(--r-pill);
  background: transparent;
  color: var(--c-text-secondary);
  font-family: var(--font-heading);
  font-size: 0.7rem; font-weight: 700;
  letter-spacing: 0.03em;
  border: none;
  cursor: pointer;
  transition: all var(--t-base) var(--ease-out);
}
.prep-mode-opt:hover { color: var(--c-primary); }
.prep-mode-opt.is-active {
  background: var(--c-primary);
  color: white;
  box-shadow: 0 1px 3px rgba(8, 112, 98, 0.30);
}

.prep-theme-options {
  display: inline-flex; flex-wrap: wrap; gap: 0.4rem;
}
.prep-theme-opt {
  display: inline-flex; align-items: center; gap: 0.45rem;
  padding: 0.4rem 0.65rem 0.4rem 0.4rem;
  background: var(--c-bg-elevated);
  border: 1.5px solid var(--c-border);
  border-radius: var(--r-pill);
  color: var(--c-text-secondary);
  font-family: var(--font-heading);
  font-size: 0.7rem; font-weight: 700;
  letter-spacing: 0.03em;
  cursor: pointer;
  transition: all var(--t-base) var(--ease-out);
}
.prep-theme-opt:hover {
  background: var(--c-bg-tinted);
  border-color: var(--c-primary);
  color: var(--c-primary);
}
.prep-theme-opt.is-active {
  border-color: var(--c-primary);
  background: var(--c-accent-green);
  color: var(--c-primary-dark);
  box-shadow: 0 1px 3px rgba(8, 112, 98, 0.20);
}
.prep-theme-swatch {
  display: inline-flex;
  border-radius: var(--r-pill);
  overflow: hidden;
  border: 1px solid rgba(26, 26, 26, 0.10);
  width: 28px; height: 14px;
}
.prep-theme-swatch > span {
  flex: 1; display: block;
}

.prep-refine-hint {
  color: var(--c-text-tertiary); font-size: 0.7rem; line-height: 1.5;
  margin-top: 0.25rem;
}
.prep-link {
  background: none; border: none; padding: 0; cursor: pointer;
  color: var(--c-primary);
  font-size: inherit; font-weight: 700;
  text-decoration: underline;
}
.prep-link:hover { color: var(--c-primary-dark); }

/* ── CTA ── */
.prep-cta {
  margin-top: 0.25rem;
  padding: 0.95rem 1.1rem;
  border: none; border-radius: var(--r-md);
  display: flex; align-items: center; justify-content: center; gap: 0.5rem;
  font-family: var(--font-heading);
  font-weight: 800; font-size: 0.9rem;
  letter-spacing: 0.04em; text-transform: uppercase;
  color: white;
  background: var(--c-primary);
  box-shadow: 0 6px 18px rgba(8, 112, 98, 0.28);
  cursor: pointer;
  transition: all var(--t-base) var(--ease-out);
}
.prep-cta:hover:not(:disabled) {
  background: var(--c-primary-dark);
  transform: translateY(-1px);
  box-shadow: 0 10px 24px rgba(8, 112, 98, 0.36);
}
.prep-cta:disabled {
  opacity: 0.45;
  cursor: not-allowed;
  background: var(--c-text-muted);
  box-shadow: none;
}

/* ── Intent banner ── */
.prep-intent {
  display: flex; align-items: flex-start; gap: 0.55rem;
  padding: 0.7rem 0.85rem;
  border-radius: var(--r-md);
  margin-bottom: 0.65rem;
  font-size: 0.74rem; line-height: 1.5;
}
.prep-intent.is-ok {
  background: rgba(105, 190, 40, 0.10);
  border: 1px solid rgba(105, 190, 40, 0.35);
  color: #2a5510;
}
.prep-intent.is-warn {
  background: #FEF7E0;
  border: 1px solid rgba(243, 196, 0, 0.45);
  color: #7a5e00;
}
.prep-intent-icon {
  flex-shrink: 0;
  margin-top: 1px;
}
.prep-intent-body { flex: 1; min-width: 0; }
.prep-intent-title {
  font-family: var(--font-heading);
  font-weight: 700;
  font-size: 0.72rem;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  margin-bottom: 0.3rem;
}
.prep-intent-list {
  list-style: none;
  padding: 0; margin: 0;
  display: flex; flex-direction: column; gap: 0.2rem;
}
.prep-intent-list li {
  font-size: 0.74rem;
}
.prep-intent-list strong { font-weight: 700; }
.prep-intent-warn { color: #7a5e00; }
`;
