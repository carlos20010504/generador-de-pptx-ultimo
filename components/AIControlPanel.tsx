"use client";

import React, { useState, useEffect, useRef } from 'react';
import {
  Sparkles, Lightbulb, LayoutDashboard, Table2, BarChart2,
  FileText, Loader2, Send, ChevronRight, RefreshCw, Wand2, Check, Brain, TrendingUp,
} from 'lucide-react';
import { aiStatusBadge, AIStatus } from '@/utils/ai-status';

interface AIControlPanelProps {
  onPromptChange: (prompt: string) => void;
  onFocusChange: (focus: 'tables' | 'charts' | 'text' | 'mixed') => void;
  onContextChange: (context: {
    audience: string;
    language: string;
    theme: {
      key: string;
      name: string;
      primary_hex: string;
      accent_hex: string;
      text_hex: string;
      bg_hex: string;
    };
  }) => void;
  file: File | null;
  isOrganizedFile?: boolean;
  originalFileName?: string | null;
}

const FOCUS_OPTIONS = [
  { id: 'mixed' as const, label: 'Auto', icon: LayoutDashboard, color: '#087062' },
  { id: 'tables' as const, label: 'Tablas', icon: Table2, color: '#69BE28' },
  { id: 'charts' as const, label: 'Gráficas', icon: BarChart2, color: '#00A0DF' },
  { id: 'text' as const, label: 'Texto', icon: FileText, color: '#F9A8D4' },
];

// Themes para el PPT generado. Default Socya Institucional usa la paleta
// oficial (teal #087062 + logo-green #69BE28); el resto son alternativas
// para audiencias específicas.
const THEME_OPTIONS = [
  { key: 'socya-institucional', name: 'Socya Institucional', primary_hex: '#087062', accent_hex: '#69BE28', text_hex: '#1A1A1A', bg_hex: '#FFFFFF' },
  { key: 'socya-comite',        name: 'Socya Comité',        primary_hex: '#123C49', accent_hex: '#F3C400', text_hex: '#1A1A1A', bg_hex: '#FFFFFF' },
  { key: 'socya-impacto',       name: 'Socya Impacto',       primary_hex: '#087062', accent_hex: '#FF8300', text_hex: '#1A1A1A', bg_hex: '#FFFFFF' },
  { key: 'analitica-moderna',   name: 'Analítica Moderna',   primary_hex: '#0F172A', accent_hex: '#2563EB', text_hex: '#E5E7EB', bg_hex: '#F8FAFC' },
];

interface RecommendedSlide {
  type?: string;
  title?: string;
  reason?: string;
}

interface IntelligenceResponse {
  executiveSummary?: string;
  suggestions?: string[];
  keyFindings?: string[];
  trends?: string[];
  promptHints?: string[];
  healthSignals?: string[];
  semanticSummary?: {
    topic?: string;
    informationType?: string;
    aboutText?: string;
    emphasis?: string[];
  };
  dataset?: {
    fileName?: string;
    workbookType?: string;
    primarySheet?: string | null;
    sheetCount?: number;
    totalRows?: number;
    totalColumns?: number;
  };
  processing?: {
    timeoutMinutes?: number;
    tier?: string;
  };
  powerPointPlan?: {
    recommendedSlides?: RecommendedSlide[];
  };
  ai_status?: { model?: string; cache_hit?: boolean; fallback_steps?: { from: string; reason: string }[] };
  audit?: { slides_planned: number; slides_validated: number; slides_dropped: unknown[]; bullets_dropped: number };
}

function compactNumber(value?: number): string {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '0';
  return new Intl.NumberFormat('es-CO').format(value);
}

function formatWorkbookType(value?: string): string {
  const normalized = String(value || '').trim();
  if (!normalized) return 'General';
  return normalized.replace(/[_-]+/g, ' ').replace(/\b\w/g, (char) => char.toUpperCase());
}

export default function AIControlPanel({
  onPromptChange,
  onFocusChange,
  onContextChange,
  file,
  isOrganizedFile = false,
  originalFileName = null,
}: AIControlPanelProps) {
  const defaultTheme = THEME_OPTIONS[0];
  const [prompt, setPrompt] = useState('');
  const [focus, setFocus] = useState<'tables' | 'charts' | 'text' | 'mixed'>('mixed');
  const [suggestions, setSuggestions] = useState<string[]>([]);
  const [findings, setFindings] = useState<string[]>([]);
  const [trends, setTrends] = useState<string[]>([]);
  const [promptHints, setPromptHints] = useState<string[]>([]);
  const [recommendedSlides, setRecommendedSlides] = useState<RecommendedSlide[]>([]);
  const [executiveSummary, setExecutiveSummary] = useState('');
  const [semanticSummary, setSemanticSummary] = useState<IntelligenceResponse['semanticSummary'] | null>(null);
  const [datasetSummary, setDatasetSummary] = useState<IntelligenceResponse['dataset'] | null>(null);
  const [healthSignals, setHealthSignals] = useState<string[]>([]);
  const [processingMessage, setProcessingMessage] = useState('');
  const [isLoadingSuggestions, setIsLoadingSuggestions] = useState(false);
  const [promptSent, setPromptSent] = useState(false);
  const [lastFile, setLastFile] = useState<File | null>(null);
  const [aiStatus, setAiStatus] = useState<AIStatus | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    onContextChange({
      audience: 'ejecutivos',
      language: 'Español',
      theme: defaultTheme,
    });
  }, [defaultTheme, onContextChange]);

  // Auto-load suggestions whenever a new file is uploaded
  useEffect(() => {
    if (!file) {
      setSuggestions([]);
      setFindings([]);
      setTrends([]);
      setPromptHints([]);
      setRecommendedSlides([]);
      setExecutiveSummary('');
      setSemanticSummary(null);
      setDatasetSummary(null);
      setHealthSignals([]);
      setProcessingMessage('');
      setLastFile(null);
      setAiStatus(null);
      return;
    }
    if (file === lastFile) return;
    setLastFile(file);
    loadSuggestions(file);
  }, [file]); // eslint-disable-line react-hooks/exhaustive-deps

  const loadSuggestions = async (f: File) => {
    setIsLoadingSuggestions(true);
    setSuggestions([]);
    try {
      const formData = new FormData();
      formData.append('file', f);
      if (prompt.trim()) {
        formData.append('userPrompt', prompt.trim());
      }
      formData.append('audience', 'ejecutivos');
      formData.append('language', 'Español');
      formData.append('theme', JSON.stringify(defaultTheme));
      const res = await fetch('/api/excel-intelligence', { method: 'POST', body: formData });
      // Antes parseábamos JSON ANTES de chequear res.ok — un 500 con HTML
      // hacía que res.json() reventara y el catch silencioso dejaba al usuario
      // sin feedback. Ahora intentamos JSON con tolerancia y damos un mensaje
      // útil cuando el body no es JSON.
      if (!res.ok) {
        let serverMsg = `No se pudo analizar el Excel (HTTP ${res.status}).`;
        try {
          const errBody = (await res.json()) as { error?: string } | null;
          if (errBody?.error) serverMsg = String(errBody.error);
        } catch { /* respuesta no-JSON, mantenemos el mensaje genérico */ }
        throw new Error(serverMsg);
      }
      const data = (await res.json()) as IntelligenceResponse;

      if (Array.isArray(data.suggestions) && data.suggestions.length > 0) {
        setSuggestions(data.suggestions);
      } else {
        setSuggestions([]);
      }

      setAiStatus(data.ai_status ?? null);
      setFindings(Array.isArray(data.keyFindings) ? data.keyFindings.slice(0, 4) : []);
      setTrends(Array.isArray(data.trends) ? data.trends.slice(0, 3) : []);
      setPromptHints(Array.isArray(data.promptHints) ? data.promptHints.slice(0, 3) : []);
      setRecommendedSlides(Array.isArray(data.powerPointPlan?.recommendedSlides) ? data.powerPointPlan!.recommendedSlides!.slice(0, 3) : []);
      setExecutiveSummary(String(data.executiveSummary || '').trim());
      setSemanticSummary(data.semanticSummary && typeof data.semanticSummary === 'object' ? data.semanticSummary : null);
      setDatasetSummary(data.dataset && typeof data.dataset === 'object' ? data.dataset : null);
      setHealthSignals(Array.isArray(data.healthSignals) ? data.healthSignals.slice(0, 3) : []);

      const timeoutMinutes = Number(data.processing?.timeoutMinutes || 0);
      const tier = String(data.processing?.tier || '').trim();
      if (timeoutMinutes > 0) {
        const tierLabel = tier ? ` · perfil ${tier}` : '';
        setProcessingMessage(`Ventana de análisis habilitada: ${timeoutMinutes} min${tierLabel}.`);
      } else {
        setProcessingMessage('');
      }
    } catch {
      setSuggestions([]);
      setAiStatus(null);
      setFindings([]);
      setTrends([]);
      setPromptHints([]);
      setRecommendedSlides([]);
      setExecutiveSummary('');
      setSemanticSummary(null);
      setDatasetSummary(null);
      setHealthSignals([]);
      setProcessingMessage('');
    } finally {
      setIsLoadingSuggestions(false);
    }
  };

  const handleFocusChange = (val: typeof focus) => {
    setFocus(val);
    onFocusChange(val);
    // Auto-suggest prompt based on focus
    const map: Record<typeof val, string> = {
      tables: 'Haz énfasis en tablas detalladas con datos reales del Excel.',
      charts: 'Enfócate en generar gráficas visuales con las métricas más importantes.',
      text: 'Prioriza el análisis textual y las conclusiones ejecutivas.',
      mixed: '',
    };
    if (!prompt && map[val]) {
      const newP = map[val];
      setPrompt(newP);
      onPromptChange(newP);
    }
  };

  const handlePromptInput = (val: string) => {
    setPrompt(val);
    setPromptSent(false);
    onPromptChange(val);
  };

  const handleSend = () => {
    onPromptChange(prompt);
    if (file) {
      void loadSuggestions(file);
    }
    setPromptSent(true);
    setTimeout(() => setPromptSent(false), 2500);
  };

  const applySuggestion = (s: string) => {
    setPrompt(s);
    setPromptSent(false);
    onPromptChange(s);
    textareaRef.current?.focus();
  };

  const applyRecommendedSlide = (slide: RecommendedSlide) => {
    const promptText = [
      slide.title ? `Quiero una diapositiva llamada "${slide.title}".` : '',
      slide.reason || '',
      'Usa datos reales del Excel y conviértelo en una parte importante de la presentación.',
    ].filter(Boolean).join(' ');
    applySuggestion(promptText);
  };

  const clearPrompt = () => {
    setPrompt('');
    onPromptChange('');
    setPromptSent(false);
  };

  const hasFile = Boolean(file);
  const currentFileName = file?.name || '';
  const sourceFileLabel = originalFileName || currentFileName;
  const summaryCards = [
    { label: 'Hoja principal', value: datasetSummary?.primarySheet || 'Sin definir' },
    { label: 'Filas', value: compactNumber(datasetSummary?.totalRows) },
    { label: 'Columnas', value: compactNumber(datasetSummary?.totalColumns) },
    { label: 'Hojas', value: compactNumber(datasetSummary?.sheetCount) },
    { label: 'Tipo', value: formatWorkbookType(datasetSummary?.workbookType) },
  ];

  return (
    <div className="ai-panel-root">
      <style>{`
        .ai-panel-root {
          width: 100%;
          display: flex;
          flex-direction: column;
          gap: 0.75rem;
        }
        /* Sidebar behavior only kicks in on desktop where the parent grid
           actually makes us a sidebar (≥1024px). On smaller screens we sit
           inline in the page flow so scroll stays on <html> (no inner
           scroll container that traps mobile gestures). */
        @media (min-width: 1024px) {
          .ai-panel-root {
            width: 340px;
            min-width: 320px;
            max-width: 360px;
            flex-shrink: 0;
            position: sticky;
            top: 1rem;
            max-height: calc(100dvh - 2rem);
            overflow-y: auto;
          }
        }
      `}</style>
      {/* Header */}
      <div style={{
        background: 'rgba(8,112,98,0.06)',
        border: '1px solid rgba(8,112,98,0.20)',
        borderRadius: '16px',
        padding: '0.85rem 0.9rem',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.55rem', marginBottom: '0.35rem' }}>
          <div style={{
            padding: '0.38rem',
            background: 'rgba(8,112,98,0.18)',
            borderRadius: '9px',
            display: 'flex',
            flexShrink: 0,
          }}>
            <Wand2 size={15} color="#A78BFA" />
          </div>
          <div>
            <p style={{ color: '#1A1A1A', fontSize: '0.82rem', fontWeight: 800, margin: 0, lineHeight: 1.2 }}>
              Asistente IA
            </p>
            <p style={{
              color: aiStatus
                ? (aiStatusBadge(aiStatus).tone === 'cache' ? '#69BE28'
                   : aiStatusBadge(aiStatus).tone === 'warn' ? '#8A6800'
                   : '#087062')
                : 'rgba(26,26,26,0.45)',
              fontSize: '0.62rem', margin: 0,
            }}>
              {aiStatus ? `✦ ${aiStatusBadge(aiStatus).label}` : 'Hermes 3 · OpenRouter'}
            </p>
          </div>
        </div>
        <p style={{ color: '#4D4F53', fontSize: '0.68rem', margin: 0, lineHeight: 1.4 }}>
          {hasFile
            ? isOrganizedFile
              ? 'La IA está preparada para leer el Excel ya organizado, resumir de qué trata y sugerir cómo construir la presentación.'
              : 'La IA leerá tu Excel, resumirá de qué trata y te sugerirá cómo convertirlo en PowerPoint.'
            : 'Sube un Excel para activar el asistente IA.'}
        </p>
        {hasFile && (
          <div style={{
            marginTop: '0.55rem',
            padding: '0.5rem 0.6rem',
            borderRadius: '10px',
            background: isOrganizedFile ? 'rgba(105,190,40,0.12)' : '#FFFFFF',
            border: isOrganizedFile ? '1px solid rgba(105,190,40,0.30)' : '1px solid rgba(26,26,26,0.10)',
          }}>
            <p style={{ color: isOrganizedFile ? '#69BE28' : '#4D4F53', fontSize: '0.64rem', fontWeight: 700, margin: '0 0 0.15rem' }}>
              {isOrganizedFile ? 'Excel organizado listo para IA' : 'Excel cargado para análisis'}
            </p>
            <p style={{ color: '#4D4F53', fontSize: '0.62rem', margin: 0, lineHeight: 1.35 }}>
              {isOrganizedFile
                ? `La lectura IA se hará sobre ${currentFileName}${sourceFileLabel && sourceFileLabel !== currentFileName ? `, generado desde ${sourceFileLabel}` : ''}.`
                : `Archivo actual: ${currentFileName}.`}
            </p>
          </div>
        )}
        {processingMessage && (
          <p style={{ color: 'rgba(167,139,250,0.8)', fontSize: '0.61rem', margin: '0.35rem 0 0', lineHeight: 1.35 }}>
            {processingMessage}
          </p>
        )}
      </div>

      {/* Focus Selector */}
      <div style={{
        background: '#FAFAFA',
        border: '1px solid rgba(26,26,26,0.08)',
        borderRadius: '14px',
        padding: '0.8rem',
      }}>
        <p style={{
          color: '#4D4F53',
          fontSize: '0.62rem',
          fontWeight: 700,
          textTransform: 'uppercase',
          letterSpacing: '0.06em',
          margin: '0 0 0.55rem',
        }}>
          Enfoque del PowerPoint
        </p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '0.4rem' }}>
          {FOCUS_OPTIONS.map(({ id, label, icon: Icon, color }) => {
            const active = focus === id;
            return (
              <button
                key={id}
                onClick={() => handleFocusChange(id)}
                style={{
                  padding: '0.5rem 0.2rem',
                  borderRadius: '10px',
                  border: `1px solid ${active ? color + '55' : 'rgba(26,26,26,0.10)'}`,
                  background: active ? `${color}18` : '#FAFAFA',
                  color: active ? color : 'rgba(26,26,26,0.45)',
                  cursor: 'pointer',
                  fontSize: '0.6rem',
                  fontWeight: 700,
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  gap: '0.28rem',
                  transition: 'all 0.18s',
                }}
              >
                <Icon size={13} />
                {label}
              </button>
            );
          })}
        </div>
      </div>

      {/* Prompt Input */}
      <div style={{
        background: '#FAFAFA',
        border: '1px solid rgba(26,26,26,0.08)',
        borderRadius: '14px',
        padding: '0.8rem',
      }}>
        <p style={{
          color: '#4D4F53',
          fontSize: '0.62rem',
          fontWeight: 700,
          textTransform: 'uppercase',
          letterSpacing: '0.06em',
          margin: '0 0 0.5rem',
        }}>
          Tus instrucciones para la IA
        </p>
        <div style={{ position: 'relative' }}>
          <textarea
            ref={textareaRef}
            value={prompt}
            onChange={(e) => handlePromptInput(e.target.value)}
            placeholder={hasFile
              ? 'Ej: Resalta los costos más altos y genera un gráfico de barras por mes...'
              : 'Sube un Excel para escribir instrucciones...'}
            disabled={!hasFile}
            style={{
              width: '100%',
              minHeight: '90px',
              background: 'rgba(0,0,0,0.25)',
              border: '1px solid rgba(26,26,26,0.15)',
              borderRadius: '10px',
              padding: '0.65rem 0.7rem 2.2rem',
              color: hasFile ? '#1A1A1A' : 'rgba(26,26,26,0.30)',
              fontSize: '0.75rem',
              fontFamily: 'inherit',
              resize: 'none',
              outline: 'none',
              transition: 'border-color 0.2s',
              boxSizing: 'border-box',
              cursor: hasFile ? 'text' : 'not-allowed',
            }}
            onFocus={(e) => { e.currentTarget.style.borderColor = 'rgba(8,112,98,0.40)'; }}
            onBlur={(e) => { e.currentTarget.style.borderColor = 'rgba(26,26,26,0.15)'; }}
          />
          {/* Send / Clear buttons row */}
          <div style={{
            position: 'absolute',
            bottom: '0.5rem',
            right: '0.5rem',
            display: 'flex',
            gap: '0.3rem',
            alignItems: 'center',
          }}>
            {prompt && (
              <button
                onClick={clearPrompt}
                title="Borrar"
                style={{
                  background: '#EEF7E6',
                  border: '1px solid rgba(26,26,26,0.10)',
                  borderRadius: '6px',
                  padding: '0.25rem 0.4rem',
                  color: 'rgba(26,26,26,0.45)',
                  fontSize: '0.58rem',
                  fontWeight: 700,
                  cursor: 'pointer',
                }}
              >
                Borrar
              </button>
            )}
            <button
              onClick={handleSend}
              disabled={!hasFile || !prompt.trim()}
              title="Confirmar instrucciones"
              style={{
                background: promptSent ? 'rgba(105,190,40,0.22)' : 'rgba(8,112,98,0.22)',
                border: `1px solid ${promptSent ? 'rgba(105,190,40,0.40)' : 'rgba(8,112,98,0.40)'}`,
                borderRadius: '6px',
                padding: '0.25rem 0.38rem',
                color: promptSent ? '#69BE28' : '#087062',
                cursor: (!hasFile || !prompt.trim()) ? 'not-allowed' : 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '0.25rem',
                fontSize: '0.6rem',
                fontWeight: 700,
                transition: 'all 0.2s',
                opacity: (!hasFile || !prompt.trim()) ? 0.4 : 1,
              }}
            >
              {promptSent ? <><Check size={11} /> Listo</> : <><Send size={11} /> Aplicar</>}
            </button>
          </div>
        </div>
      </div>

      {/* Suggestions */}
      <div style={{
        background: '#FAFAFA',
        border: '1px solid rgba(26,26,26,0.08)',
        borderRadius: '14px',
        padding: '0.8rem',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '0.55rem' }}>
          <p style={{
            color: '#4D4F53',
            fontSize: '0.62rem',
            fontWeight: 700,
            textTransform: 'uppercase',
            letterSpacing: '0.06em',
            margin: 0,
          }}>
            {isLoadingSuggestions ? 'Analizando Excel...' : 'Sugerencias para hacerlo'}
          </p>
          {hasFile && !isLoadingSuggestions && (
            <button
              onClick={() => file && loadSuggestions(file)}
              title="Regenerar sugerencias"
              style={{
                background: 'none',
                border: 'none',
                color: 'rgba(26,26,26,0.30)',
                cursor: 'pointer',
                padding: '0.2rem',
                display: 'flex',
              }}
            >
              <RefreshCw size={11} />
            </button>
          )}
        </div>

        {isLoadingSuggestions ? (
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', padding: '0.5rem 0' }}>
            <Loader2 size={13} color="#A78BFA" style={{ animation: 'spin 1s linear infinite' }} />
            <span style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.7rem' }}>
              {isOrganizedFile
                ? 'La IA está leyendo el Excel organizado y proponiendo la mejor forma de presentarlo...'
                : 'La IA está leyendo tu Excel y construyendo una estrategia de presentación...'}
            </span>
          </div>
        ) : suggestions.length > 0 ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.35rem' }}>
            {suggestions.map((s, idx) => (
              <button
                key={idx}
                onClick={() => applySuggestion(s)}
                style={{
                  padding: '0.5rem 0.6rem',
                  borderRadius: '9px',
                  background: '#F2F8F6',
                  border: '1px solid rgba(26,26,26,0.10)',
                  color: '#4D4F53',
                  fontSize: '0.68rem',
                  cursor: 'pointer',
                  textAlign: 'left',
                  display: 'flex',
                  alignItems: 'flex-start',
                  gap: '0.4rem',
                  transition: 'all 0.18s',
                  lineHeight: 1.35,
                }}
                onMouseOver={(e) => {
                  e.currentTarget.style.background = 'rgba(8,112,98,0.10)';
                  e.currentTarget.style.borderColor = 'rgba(8,112,98,0.22)';
                  e.currentTarget.style.color = 'white';
                }}
                onMouseOut={(e) => {
                  e.currentTarget.style.background = '#F2F8F6';
                  e.currentTarget.style.borderColor = 'rgba(26,26,26,0.10)';
                  e.currentTarget.style.color = '#4D4F53';
                }}
              >
                <ChevronRight size={11} style={{ marginTop: '2px', flexShrink: 0, color: '#087062' }} />
                {s}
              </button>
            ))}
          </div>
        ) : (
          <div style={{ padding: '0.5rem 0', display: 'flex', alignItems: 'center', gap: '0.4rem' }}>
            <Lightbulb size={12} color="rgba(26,26,26,0.25)" />
            <span style={{ color: 'rgba(26,26,26,0.30)', fontSize: '0.68rem' }}>
              {hasFile ? 'Sin sugerencias IA por ahora — escribe tu propio prompt.' : 'Sube un Excel para ver sugerencias.'}
            </span>
          </div>
        )}

        {/* Status indicator */}
        {hasFile && !isLoadingSuggestions && (
          <div style={{
            marginTop: '0.6rem',
            paddingTop: '0.6rem',
            borderTop: '1px solid rgba(26,26,26,0.06)',
            display: 'flex',
            alignItems: 'center',
            gap: '0.35rem',
          }}>
            <Sparkles size={10} color="#A78BFA" />
            <span style={{ color: 'rgba(26,26,26,0.30)', fontSize: '0.6rem' }}>
              Usa una sugerencia como base o escribe tu propio prompt para personalizar el PowerPoint
            </span>
          </div>
        )}
      </div>

      {/* Inteligencia contextual */}
      {hasFile && !isLoadingSuggestions && (executiveSummary || semanticSummary || datasetSummary || findings.length > 0 || trends.length > 0 || recommendedSlides.length > 0 || healthSignals.length > 0) && (
        <div style={{
          background: '#FAFAFA',
          border: '1px solid rgba(26,26,26,0.08)',
          borderRadius: '14px',
          padding: '0.8rem',
        }}>
          <p style={{
            color: '#4D4F53',
            fontSize: '0.62rem',
            fontWeight: 700,
            textTransform: 'uppercase',
            letterSpacing: '0.06em',
            margin: '0 0 0.5rem',
            display: 'flex',
            alignItems: 'center',
            gap: '0.3rem',
          }}>
            <Brain size={11} />
            Inteligencia del Excel
          </p>

          {datasetSummary && (
            <div style={{ marginBottom: '0.65rem' }}>
              <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.6rem', margin: '0 0 0.38rem', fontWeight: 700 }}>
                Resumen del Excel
              </p>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '0.35rem' }}>
                {summaryCards.map((item) => (
                  <div key={item.label} style={{ padding: '0.42rem 0.48rem', borderRadius: '9px', background: '#FAFAFA', border: '1px solid rgba(26,26,26,0.08)' }}>
                    <p style={{ color: 'rgba(26,26,26,0.30)', fontSize: '0.55rem', margin: '0 0 0.1rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                      {item.label}
                    </p>
                    <p style={{ color: 'rgba(26,26,26,0.70)', fontSize: '0.65rem', margin: 0, lineHeight: 1.3, fontWeight: 700 }}>
                      {item.value}
                    </p>
                  </div>
                ))}
              </div>
            </div>
          )}

          {semanticSummary && (
            <div style={{ marginBottom: '0.6rem' }}>
              <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.6rem', margin: '0 0 0.3rem', fontWeight: 700 }}>
                Tema del Excel
              </p>
              {semanticSummary.topic && (
                <p style={{ color: 'rgba(26,26,26,0.72)', fontSize: '0.68rem', margin: '0 0 0.18rem', lineHeight: 1.4, fontWeight: 700 }}>
                  {semanticSummary.topic}
                </p>
              )}
              {semanticSummary.informationType && (
                <p style={{ color: 'rgba(26,26,26,0.55)', fontSize: '0.64rem', margin: '0 0 0.22rem', lineHeight: 1.4 }}>
                  Tipo de información: {semanticSummary.informationType}
                </p>
              )}
              {Array.isArray(semanticSummary.emphasis) && semanticSummary.emphasis.length > 0 && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '0.24rem' }}>
                  {semanticSummary.emphasis.map((item, idx) => (
                    <div key={`semantic-${idx}`} style={{ color: 'rgba(26,26,26,0.55)', fontSize: '0.62rem', lineHeight: 1.35 }}>
                      • {item}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {executiveSummary && (
            <div style={{ marginBottom: '0.55rem' }}>
              <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.6rem', margin: '0 0 0.25rem', fontWeight: 700 }}>
                De qué trata el Excel
              </p>
              <p style={{ color: '#4D4F53', fontSize: '0.69rem', margin: 0, lineHeight: 1.45 }}>
              {executiveSummary}
              </p>
            </div>
          )}

          {findings.length > 0 && (
            <div style={{ marginBottom: '0.55rem' }}>
              <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.6rem', margin: '0 0 0.32rem', fontWeight: 700 }}>
                Hallazgos clave
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.32rem' }}>
                {findings.map((item, idx) => (
                  <div key={`finding-${idx}`} style={{ color: '#4D4F53', fontSize: '0.66rem', lineHeight: 1.35 }}>
                    • {item}
                  </div>
                ))}
              </div>
            </div>
          )}

          {trends.length > 0 && (
            <div style={{ marginBottom: '0.55rem' }}>
              <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.6rem', margin: '0 0 0.32rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                <TrendingUp size={10} />
                Tendencias
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.28rem' }}>
                {trends.map((item, idx) => (
                  <div key={`trend-${idx}`} style={{ color: 'rgba(26,26,26,0.55)', fontSize: '0.64rem', lineHeight: 1.35 }}>
                    • {item}
                  </div>
                ))}
              </div>
            </div>
          )}

          {recommendedSlides.length > 0 && (
            <div style={{ marginBottom: healthSignals.length > 0 ? '0.55rem' : 0 }}>
              <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.6rem', margin: '0 0 0.32rem', fontWeight: 700 }}>
                Cómo recomienda presentarlo
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.3rem' }}>
                {recommendedSlides.map((slide, idx) => (
                  <div key={`rec-slide-${idx}`} style={{ padding: '0.45rem 0.5rem', borderRadius: '8px', background: '#FAFAFA', border: '1px solid rgba(26,26,26,0.08)' }}>
                    <p style={{ color: '#4D4F53', fontSize: '0.65rem', margin: '0 0 0.12rem', fontWeight: 700 }}>
                      {slide.title || `Diapositiva ${idx + 1}`}
                    </p>
                    <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.61rem', margin: 0, lineHeight: 1.35 }}>
                      {slide.reason || 'Bloque recomendado por la IA para mejorar la narrativa.'}
                    </p>
                    <button
                      type="button"
                      onClick={() => applyRecommendedSlide(slide)}
                      style={{
                        marginTop: '0.4rem',
                        padding: '0.32rem 0.48rem',
                        borderRadius: '8px',
                        border: '1px solid rgba(8,112,98,0.25)',
                        background: 'rgba(8,112,98,0.10)',
                        color: '#C4B5FD',
                        cursor: 'pointer',
                        fontSize: '0.6rem',
                        fontWeight: 700,
                      }}
                    >
                      Usar como prompt
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {healthSignals.length > 0 && (
            <div>
              <p style={{ color: 'rgba(26,26,26,0.45)', fontSize: '0.6rem', margin: '0 0 0.32rem', fontWeight: 700 }}>
                Estado del análisis
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.3rem' }}>
                {healthSignals.map((signal, idx) => (
                  <div key={`health-${idx}`} style={{ color: '#4D4F53', fontSize: '0.63rem', lineHeight: 1.35 }}>
                    • {signal}
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Ayuda de prompts */}
      {hasFile && !isLoadingSuggestions && promptHints.length > 0 && (
        <div style={{
          background: '#FAFAFA',
          border: '1px solid rgba(26,26,26,0.08)',
          borderRadius: '14px',
          padding: '0.8rem',
        }}>
          <p style={{
            color: '#4D4F53',
            fontSize: '0.62rem',
            fontWeight: 700,
            textTransform: 'uppercase',
            letterSpacing: '0.06em',
            margin: '0 0 0.5rem',
          }}>
            Guía para tu prompt
          </p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.32rem' }}>
            {promptHints.map((hint, idx) => (
              <div key={`prompt-hint-${idx}`} style={{ color: 'rgba(26,26,26,0.55)', fontSize: '0.65rem', lineHeight: 1.38 }}>
                • {hint}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
