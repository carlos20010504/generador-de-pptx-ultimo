"use client";

import React, { useState, useEffect, useRef } from 'react';
import {
  Sparkles, Lightbulb, LayoutDashboard, Table2, BarChart2,
  FileText, Loader2, Send, ChevronRight, RefreshCw, Wand2, Check, Brain, TrendingUp,
} from 'lucide-react';

interface AIControlPanelProps {
  onPromptChange: (prompt: string) => void;
  onFocusChange: (focus: 'tables' | 'charts' | 'text' | 'mixed') => void;
  file: File | null;
}

const FOCUS_OPTIONS = [
  { id: 'mixed' as const, label: 'Auto', icon: LayoutDashboard, color: '#A78BFA' },
  { id: 'tables' as const, label: 'Tablas', icon: Table2, color: '#34D399' },
  { id: 'charts' as const, label: 'Gráficas', icon: BarChart2, color: '#60A5FA' },
  { id: 'text' as const, label: 'Texto', icon: FileText, color: '#F9A8D4' },
];

const PLACEHOLDER_SUGGESTIONS = [
  'Resalta los 3 KPIs más críticos en la portada',
  'Enfócate en la comparación de períodos mensuales',
  'Genera una diapositiva ejecutiva con conclusiones',
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
  processing?: {
    timeoutMinutes?: number;
    tier?: string;
  };
  powerPointPlan?: {
    recommendedSlides?: RecommendedSlide[];
  };
}

export default function AIControlPanel({ onPromptChange, onFocusChange, file }: AIControlPanelProps) {
  const [prompt, setPrompt] = useState('');
  const [focus, setFocus] = useState<'tables' | 'charts' | 'text' | 'mixed'>('mixed');
  const [suggestions, setSuggestions] = useState<string[]>([]);
  const [findings, setFindings] = useState<string[]>([]);
  const [trends, setTrends] = useState<string[]>([]);
  const [promptHints, setPromptHints] = useState<string[]>([]);
  const [recommendedSlides, setRecommendedSlides] = useState<RecommendedSlide[]>([]);
  const [executiveSummary, setExecutiveSummary] = useState('');
  const [processingMessage, setProcessingMessage] = useState('');
  const [isLoadingSuggestions, setIsLoadingSuggestions] = useState(false);
  const [promptSent, setPromptSent] = useState(false);
  const [lastFile, setLastFile] = useState<File | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Auto-load suggestions whenever a new file is uploaded
  useEffect(() => {
    if (!file) {
      setSuggestions([]);
      setFindings([]);
      setTrends([]);
      setPromptHints([]);
      setRecommendedSlides([]);
      setExecutiveSummary('');
      setProcessingMessage('');
      setLastFile(null);
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
      const res = await fetch('/api/excel-intelligence', { method: 'POST', body: formData });
      const data = (await res.json()) as IntelligenceResponse;

      if (!res.ok) {
        throw new Error(String((data as { error?: string })?.error || 'No se pudo analizar el Excel.'));
      }

      if (Array.isArray(data.suggestions) && data.suggestions.length > 0) {
        setSuggestions(data.suggestions);
      } else {
        setSuggestions(PLACEHOLDER_SUGGESTIONS);
      }

      setFindings(Array.isArray(data.keyFindings) ? data.keyFindings.slice(0, 4) : []);
      setTrends(Array.isArray(data.trends) ? data.trends.slice(0, 3) : []);
      setPromptHints(Array.isArray(data.promptHints) ? data.promptHints.slice(0, 3) : []);
      setRecommendedSlides(Array.isArray(data.powerPointPlan?.recommendedSlides) ? data.powerPointPlan!.recommendedSlides!.slice(0, 3) : []);
      setExecutiveSummary(String(data.executiveSummary || '').trim());

      const timeoutMinutes = Number(data.processing?.timeoutMinutes || 0);
      const tier = String(data.processing?.tier || '').trim();
      if (timeoutMinutes > 0) {
        const tierLabel = tier ? ` · perfil ${tier}` : '';
        setProcessingMessage(`Ventana de análisis habilitada: ${timeoutMinutes} min${tierLabel}.`);
      } else {
        setProcessingMessage('');
      }
    } catch {
      setSuggestions(PLACEHOLDER_SUGGESTIONS);
      setFindings([]);
      setTrends([]);
      setPromptHints([]);
      setRecommendedSlides([]);
      setExecutiveSummary('');
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

  const clearPrompt = () => {
    setPrompt('');
    onPromptChange('');
    setPromptSent(false);
  };

  const hasFile = Boolean(file);

  return (
    <div style={{
      width: '340px',
      minWidth: '320px',
      maxWidth: '360px',
      flexShrink: 0,
      display: 'flex',
      flexDirection: 'column',
      gap: '0.75rem',
      position: 'sticky',
      top: '1rem',
      maxHeight: 'calc(100dvh - 2rem)',
      overflowY: 'auto',
    }}>
      {/* Header */}
      <div style={{
        background: 'rgba(124,58,237,0.08)',
        border: '1px solid rgba(124,58,237,0.2)',
        borderRadius: '16px',
        padding: '0.85rem 0.9rem',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.55rem', marginBottom: '0.35rem' }}>
          <div style={{
            padding: '0.38rem',
            background: 'rgba(124,58,237,0.18)',
            borderRadius: '9px',
            display: 'flex',
            flexShrink: 0,
          }}>
            <Wand2 size={15} color="#A78BFA" />
          </div>
          <div>
            <p style={{ color: 'white', fontSize: '0.82rem', fontWeight: 800, margin: 0, lineHeight: 1.2 }}>
              Asistente IA
            </p>
            <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: '0.62rem', margin: 0 }}>
              Hermes 3 · OpenRouter
            </p>
          </div>
        </div>
        <p style={{ color: 'rgba(255,255,255,0.42)', fontSize: '0.68rem', margin: 0, lineHeight: 1.4 }}>
          {hasFile
            ? 'La IA leerá tu Excel completo, detectará insights y aplicará tus instrucciones al PowerPoint.'
            : 'Sube un Excel para activar el asistente IA.'}
        </p>
        {processingMessage && (
          <p style={{ color: 'rgba(167,139,250,0.8)', fontSize: '0.61rem', margin: '0.35rem 0 0', lineHeight: 1.35 }}>
            {processingMessage}
          </p>
        )}
      </div>

      {/* Focus Selector */}
      <div style={{
        background: 'rgba(255,255,255,0.02)',
        border: '1px solid rgba(255,255,255,0.06)',
        borderRadius: '14px',
        padding: '0.8rem',
      }}>
        <p style={{
          color: 'rgba(255,255,255,0.45)',
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
                  border: `1px solid ${active ? color + '55' : 'rgba(255,255,255,0.07)'}`,
                  background: active ? `${color}18` : 'rgba(255,255,255,0.02)',
                  color: active ? color : 'rgba(255,255,255,0.35)',
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
        background: 'rgba(255,255,255,0.02)',
        border: '1px solid rgba(255,255,255,0.06)',
        borderRadius: '14px',
        padding: '0.8rem',
      }}>
        <p style={{
          color: 'rgba(255,255,255,0.45)',
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
              border: '1px solid rgba(255,255,255,0.09)',
              borderRadius: '10px',
              padding: '0.65rem 0.7rem 2.2rem',
              color: hasFile ? 'white' : 'rgba(255,255,255,0.25)',
              fontSize: '0.75rem',
              fontFamily: 'inherit',
              resize: 'none',
              outline: 'none',
              transition: 'border-color 0.2s',
              boxSizing: 'border-box',
              cursor: hasFile ? 'text' : 'not-allowed',
            }}
            onFocus={(e) => { e.currentTarget.style.borderColor = 'rgba(124,58,237,0.4)'; }}
            onBlur={(e) => { e.currentTarget.style.borderColor = 'rgba(255,255,255,0.09)'; }}
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
                  background: 'rgba(255,255,255,0.06)',
                  border: '1px solid rgba(255,255,255,0.1)',
                  borderRadius: '6px',
                  padding: '0.25rem 0.4rem',
                  color: 'rgba(255,255,255,0.35)',
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
                background: promptSent ? 'rgba(74,222,128,0.2)' : 'rgba(124,58,237,0.25)',
                border: `1px solid ${promptSent ? 'rgba(74,222,128,0.4)' : 'rgba(124,58,237,0.4)'}`,
                borderRadius: '6px',
                padding: '0.25rem 0.38rem',
                color: promptSent ? '#4ADE80' : '#A78BFA',
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
        background: 'rgba(255,255,255,0.02)',
        border: '1px solid rgba(255,255,255,0.06)',
        borderRadius: '14px',
        padding: '0.8rem',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '0.55rem' }}>
          <p style={{
            color: 'rgba(255,255,255,0.45)',
            fontSize: '0.62rem',
            fontWeight: 700,
            textTransform: 'uppercase',
            letterSpacing: '0.06em',
            margin: 0,
          }}>
            {isLoadingSuggestions ? 'Analizando Excel...' : 'Sugerencias IA'}
          </p>
          {hasFile && !isLoadingSuggestions && (
            <button
              onClick={() => file && loadSuggestions(file)}
              title="Regenerar sugerencias"
              style={{
                background: 'none',
                border: 'none',
                color: 'rgba(255,255,255,0.3)',
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
            <span style={{ color: 'rgba(255,255,255,0.35)', fontSize: '0.7rem' }}>
              La IA está leyendo tu Excel y construyendo una estrategia de presentación...
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
                  background: 'rgba(255,255,255,0.03)',
                  border: '1px solid rgba(255,255,255,0.07)',
                  color: 'rgba(255,255,255,0.55)',
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
                  e.currentTarget.style.background = 'rgba(124,58,237,0.1)';
                  e.currentTarget.style.borderColor = 'rgba(124,58,237,0.25)';
                  e.currentTarget.style.color = 'white';
                }}
                onMouseOut={(e) => {
                  e.currentTarget.style.background = 'rgba(255,255,255,0.03)';
                  e.currentTarget.style.borderColor = 'rgba(255,255,255,0.07)';
                  e.currentTarget.style.color = 'rgba(255,255,255,0.55)';
                }}
              >
                <ChevronRight size={11} style={{ marginTop: '2px', flexShrink: 0, color: '#A78BFA' }} />
                {s}
              </button>
            ))}
          </div>
        ) : (
          <div style={{ padding: '0.5rem 0', display: 'flex', alignItems: 'center', gap: '0.4rem' }}>
            <Lightbulb size={12} color="rgba(255,255,255,0.2)" />
            <span style={{ color: 'rgba(255,255,255,0.25)', fontSize: '0.68rem' }}>
              {hasFile ? 'No se cargaron sugerencias.' : 'Sube un Excel para ver sugerencias.'}
            </span>
          </div>
        )}

        {/* Status indicator */}
        {hasFile && !isLoadingSuggestions && (
          <div style={{
            marginTop: '0.6rem',
            paddingTop: '0.6rem',
            borderTop: '1px solid rgba(255,255,255,0.05)',
            display: 'flex',
            alignItems: 'center',
            gap: '0.35rem',
          }}>
            <Sparkles size={10} color="#A78BFA" />
            <span style={{ color: 'rgba(255,255,255,0.28)', fontSize: '0.6rem' }}>
              Las instrucciones se aplican al generar el PowerPoint
            </span>
          </div>
        )}
      </div>

      {/* Inteligencia contextual */}
      {hasFile && !isLoadingSuggestions && (executiveSummary || findings.length > 0 || trends.length > 0 || recommendedSlides.length > 0) && (
        <div style={{
          background: 'rgba(255,255,255,0.02)',
          border: '1px solid rgba(255,255,255,0.06)',
          borderRadius: '14px',
          padding: '0.8rem',
        }}>
          <p style={{
            color: 'rgba(255,255,255,0.45)',
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

          {executiveSummary && (
            <p style={{ color: 'rgba(255,255,255,0.62)', fontSize: '0.69rem', margin: '0 0 0.55rem', lineHeight: 1.45 }}>
              {executiveSummary}
            </p>
          )}

          {findings.length > 0 && (
            <div style={{ marginBottom: '0.55rem' }}>
              <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: '0.6rem', margin: '0 0 0.32rem', fontWeight: 700 }}>
                Hallazgos clave
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.32rem' }}>
                {findings.map((item, idx) => (
                  <div key={`finding-${idx}`} style={{ color: 'rgba(255,255,255,0.55)', fontSize: '0.66rem', lineHeight: 1.35 }}>
                    • {item}
                  </div>
                ))}
              </div>
            </div>
          )}

          {trends.length > 0 && (
            <div style={{ marginBottom: '0.55rem' }}>
              <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: '0.6rem', margin: '0 0 0.32rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                <TrendingUp size={10} />
                Tendencias
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.28rem' }}>
                {trends.map((item, idx) => (
                  <div key={`trend-${idx}`} style={{ color: 'rgba(255,255,255,0.5)', fontSize: '0.64rem', lineHeight: 1.35 }}>
                    • {item}
                  </div>
                ))}
              </div>
            </div>
          )}

          {recommendedSlides.length > 0 && (
            <div>
              <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: '0.6rem', margin: '0 0 0.32rem', fontWeight: 700 }}>
                Estructura sugerida
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.3rem' }}>
                {recommendedSlides.map((slide, idx) => (
                  <div key={`rec-slide-${idx}`} style={{ padding: '0.45rem 0.5rem', borderRadius: '8px', background: 'rgba(255,255,255,0.025)', border: '1px solid rgba(255,255,255,0.06)' }}>
                    <p style={{ color: 'rgba(255,255,255,0.62)', fontSize: '0.65rem', margin: '0 0 0.12rem', fontWeight: 700 }}>
                      {slide.title || `Diapositiva ${idx + 1}`}
                    </p>
                    <p style={{ color: 'rgba(255,255,255,0.38)', fontSize: '0.61rem', margin: 0, lineHeight: 1.35 }}>
                      {slide.reason || 'Bloque recomendado por la IA para mejorar la narrativa.'}
                    </p>
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
          background: 'rgba(255,255,255,0.02)',
          border: '1px solid rgba(255,255,255,0.06)',
          borderRadius: '14px',
          padding: '0.8rem',
        }}>
          <p style={{
            color: 'rgba(255,255,255,0.45)',
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
              <div key={`prompt-hint-${idx}`} style={{ color: 'rgba(255,255,255,0.5)', fontSize: '0.65rem', lineHeight: 1.38 }}>
                • {hint}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
