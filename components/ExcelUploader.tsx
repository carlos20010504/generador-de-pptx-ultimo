"use client";

import React, { useState, useRef, useCallback, useEffect } from 'react';
import {
  UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle,
  Loader2, Settings2, BarChart2, Table2, LayoutDashboard,
  Sparkles, Download, RefreshCw, ChevronRight, Zap, Shield, Layers,
  TrendingUp, PieChart,
} from 'lucide-react';
import type { LucideIcon } from 'lucide-react';
import { ACCEPTED_EXCEL_EXTENSIONS, MAX_EXCEL_UPLOAD_BYTES, validateExcelUpload } from '@/utils/excel-file';
import { generatePowerPointFromExcel } from '@/utils/pptx-helper';
import { autoOrganizeExcel, OrganizerMode } from '@/utils/excel-organizer';

type Tab = 'generate' | 'organize';
type Status = 'idle' | 'processing' | 'success' | 'organized' | 'error';

interface GenerationStats {
  duration: number;
  mode: OrganizerMode;
  fileName: string;
}

const STEPS = [
  { label: 'Analizando estructura del Excel', icon: Layers, pct: 20 },
  { label: 'Detectando tipos de datos y columnas', icon: TrendingUp, pct: 40 },
  { label: 'Construyendo diapositivas inteligentes', icon: PieChart, pct: 70 },
  { label: 'Renderizando presentación final', icon: Sparkles, pct: 90 },
  { label: 'Generando archivo y descargando', icon: Download, pct: 100 },
];

const MODE_OPTIONS: { id: OrganizerMode; label: string; desc: string; Icon: LucideIcon; color: string; glow: string }[] = [
  {
    id: 'charts', label: 'Gráficos',
    desc: 'Énfasis en charts y visualizaciones',
    Icon: BarChart2, color: '#0891B2', glow: 'rgba(8,145,178,0.35)',
  },
  {
    id: 'tables', label: 'Tablas',
    desc: 'Énfasis en tablas de datos limpias',
    Icon: Table2, color: '#059669', glow: 'rgba(5,150,105,0.35)',
  },
  {
    id: 'mixed', label: 'Mixto',
    desc: 'Balance entre gráficos y tablas',
    Icon: LayoutDashboard, color: '#7C3AED', glow: 'rgba(124,58,237,0.35)',
  },
];

function hexToRgb(hex: string): string {
  const h = hex.replace('#', '');
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  return `${r},${g},${b}`;
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getErrorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  return fallback;
}

function validateSelectedFile(file: File): string | null {
  const validationError = validateExcelUpload(file);
  if (!validationError) {
    return null;
  }

  return validationError
    .replace('El archivo debe ser un Excel valido', 'Por favor sube un archivo Excel válido')
    .replace('El archivo subido no tiene un tipo MIME de Excel permitido.', 'El archivo seleccionado no tiene un formato Excel permitido.')
    .replace('El archivo Excel esta vacio.', 'El archivo está vacío.')
    .replace('El archivo excede el limite permitido', 'El archivo supera el límite');
}

export default function ExcelUploader() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragActive, setIsDragActive] = useState(false);
  const [status, setStatus] = useState<Status>('idle');
  const [errorMessage, setErrorMessage] = useState('');
  const [activeTab, setActiveTab] = useState<Tab>('generate');
  const [orgMode, setOrgMode] = useState<OrganizerMode>('mixed');
  const [currentStep, setCurrentStep] = useState(0);
  const [stats, setStats] = useState<GenerationStats | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Progress simulation during processing
  useEffect(() => {
    if (status !== 'processing') return;
    setCurrentStep(0);
    const intervals = [800, 1200, 1500, 2000, 1000];
    let step = 0;
    let timeoutId: ReturnType<typeof setTimeout>;

    const advance = () => {
      if (step < STEPS.length - 1) {
        step++;
        setCurrentStep(step);
        timeoutId = setTimeout(advance, intervals[step] || 1000);
      }
    };
    timeoutId = setTimeout(advance, intervals[0]);
    return () => clearTimeout(timeoutId);
  }, [status]);

  // File handlers
  const setValidFile = useCallback((f: File) => {
    const validationError = validateSelectedFile(f);
    if (!validationError) {
      setFile(f);
      setStatus('idle');
      setErrorMessage('');
      setStats(null);
    } else {
      setFile(null);
      setStatus('error');
      setErrorMessage(validationError);
    }
  }, []);

  const handleDragOver = (e: React.DragEvent) => { e.preventDefault(); setIsDragActive(true); };
  const handleDragLeave = (e: React.DragEvent) => { e.preventDefault(); setIsDragActive(false); };
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragActive(false);
    if (e.dataTransfer.files[0]) setValidFile(e.dataTransfer.files[0]);
  };
  const handleInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.[0]) setValidFile(e.target.files[0]);
  };

  // Generate action
  const handleGenerate = async () => {
    if (!file) return;
    setStatus('processing');
    setErrorMessage('');
    setStats(null);
    const start = performance.now();
    try {
      await generatePowerPointFromExcel(file, orgMode);
      const duration = ((performance.now() - start) / 1000);
      setStats({ duration, mode: orgMode, fileName: file.name });
      setStatus('success');
    } catch (err: unknown) {
      console.error('Error generating PPTX:', err);
      setStatus('error');
      setErrorMessage(getErrorMessage(err, 'Error al generar la presentación. Revisa la consola.'));
    }
  };

  // Organize action
  const handleOrganize = async () => {
    if (!file) return;
    setStatus('processing');
    setErrorMessage('');
    try {
      const blob = await autoOrganizeExcel(file, orgMode);
      setStatus('organized');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `ORGANIZADO_${orgMode.toUpperCase()}_${file.name}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: unknown) {
      console.error('Error organizing Excel:', err);
      setStatus('error');
      setErrorMessage(getErrorMessage(err, 'Error al organizar el archivo Excel.'));
    }
  };

  const isLoading = status === 'processing';
  const activeModeInfo = MODE_OPTIONS.find((m) => m.id === orgMode) ?? MODE_OPTIONS[2];

  return (
    <div style={{
      background: 'rgba(255,255,255,0.03)',
      backdropFilter: 'blur(24px) saturate(180%)',
      borderRadius: '20px',
      border: '1px solid rgba(255,255,255,0.08)',
      overflow: 'hidden',
      boxShadow: '0 24px 80px rgba(0,0,0,0.5), 0 0 1px rgba(255,255,255,0.1) inset',
    }}
    className="animate-scale-in"
    >
      {/* ── Header Tabs ── */}
      <div style={{
        display: 'flex', gap: '0.35rem', padding: '0.7rem 0.7rem 0',
        background: 'rgba(0,0,0,0.25)',
        borderBottom: '1px solid rgba(255,255,255,0.05)',
      }}>
        {([['generate', '✦ Generar PPTX'], ['organize', '⚙ Organizar Excel']] as [Tab, string][]).map(([tab, label]) => (
          <button
            key={tab}
            onClick={() => { setActiveTab(tab); setStatus('idle'); setErrorMessage(''); setStats(null); }}
            style={{
              flex: 1, padding: '0.65rem 0.5rem', border: 'none', cursor: 'pointer',
              borderRadius: '12px 12px 0 0', fontWeight: 700, fontSize: '0.82rem',
              transition: 'all 0.25s ease',
              background: activeTab === tab ? 'rgba(255,255,255,0.06)' : 'transparent',
              color: activeTab === tab ? 'white' : 'rgba(255,255,255,0.35)',
              borderBottom: activeTab === tab ? '2px solid #3B82F6' : '2px solid transparent',
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.4rem',
            }}
          >
            {label}
          </button>
        ))}
      </div>

      <div style={{ padding: '1.5rem' }}>

        {/* ── Helper Text ── */}
        <div style={{
          display: 'flex', alignItems: 'center', gap: '0.5rem',
          marginBottom: '1.15rem', padding: '0.6rem 0.75rem',
          background: 'rgba(59,130,246,0.06)',
          borderRadius: '10px',
          border: '1px solid rgba(59,130,246,0.12)',
        }}
        className="animate-fade-in"
        >
          <Zap size={14} color="#60A5FA" style={{ flexShrink: 0 }} />
          <p style={{
            color: 'rgba(255,255,255,0.45)', fontSize: '0.77rem', margin: 0,
            lineHeight: 1.5,
          }}>
            {activeTab === 'generate'
              ? <>Sube tu Excel — el motor analiza <strong style={{ color: 'rgba(255,255,255,0.65)' }}>todas las hojas</strong>, detecta tipos de datos, y genera gráficos, tablas y KPIs automáticamente.</>
              : 'Elige el modo de organización y descarga un Excel optimizado para generar el PPTX.'}
          </p>
        </div>

        {/* ── Mode Selector (Organize tab only) ── */}
        {activeTab === 'organize' && (
          <div style={{ marginBottom: '1.25rem' }} className="animate-fade-in-up">
            <p style={{
              fontSize: '0.72rem', fontWeight: 700, color: 'rgba(255,255,255,0.45)',
              marginBottom: '0.55rem', letterSpacing: '0.08em', textTransform: 'uppercase',
            }}>
              Modo de organización
            </p>
            <div style={{ display: 'flex', gap: '0.45rem' }}>
              {MODE_OPTIONS.map(({ id, label, desc, Icon, color, glow }) => {
                const active = orgMode === id;
                return (
                  <button
                    key={id}
                    onClick={() => setOrgMode(id)}
                    style={{
                      flex: 1, padding: '0.75rem 0.4rem',
                      border: `1.5px solid ${active ? color : 'rgba(255,255,255,0.08)'}`,
                      borderRadius: '12px', cursor: 'pointer', transition: 'all 0.25s ease',
                      background: active ? `rgba(${hexToRgb(color)},0.12)` : 'rgba(255,255,255,0.02)',
                      boxShadow: active ? `0 0 16px ${glow}` : 'none',
                      display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '0.35rem',
                    }}
                  >
                    <div style={{
                      background: active ? `rgba(${hexToRgb(color)},0.2)` : 'rgba(255,255,255,0.05)',
                      borderRadius: '8px', padding: '0.4rem',
                      transition: 'all 0.25s',
                    }}>
                      <Icon size={18} color={active ? color : 'rgba(255,255,255,0.3)'} />
                    </div>
                    <span style={{ fontSize: '0.78rem', fontWeight: 700, color: active ? color : 'rgba(255,255,255,0.4)' }}>
                      {label}
                    </span>
                    <span style={{ fontSize: '0.66rem', color: 'rgba(255,255,255,0.25)', textAlign: 'center', lineHeight: 1.3 }}>
                      {desc}
                    </span>
                  </button>
                );
              })}
            </div>
            <div style={{
              marginTop: '0.55rem', padding: '0.5rem 0.75rem',
              background: `rgba(${hexToRgb(activeModeInfo.color)},0.06)`,
              borderRadius: '8px', border: `1px solid rgba(${hexToRgb(activeModeInfo.color)},0.15)`,
            }}>
              <p style={{ fontSize: '0.73rem', color: 'rgba(255,255,255,0.45)', margin: 0 }}>
                {orgMode === 'charts' && <><>{'\u{1F3A8}'}</> <strong style={{ color: activeModeInfo.color }}>Modo Gráficos:</strong> Priorizará visualizaciones y charts siempre que sea posible.</>}
                {orgMode === 'tables' && <><>{'\u{1F4CB}'}</> <strong style={{ color: activeModeInfo.color }}>Modo Tablas:</strong> Tablas bien estructuradas con columnas limpias y proporcionales.</>}
                {orgMode === 'mixed' && <><>{'\u26A1'}</> <strong style={{ color: activeModeInfo.color }}>Modo Mixto:</strong> Balance inteligente — gráficos numéricos, tablas cualitativas.</>}
              </p>
            </div>
          </div>
        )}

        {/* ── Progress Bar (during processing) ── */}
        {isLoading && (
          <div style={{ marginBottom: '1.25rem' }} className="animate-fade-in">
            <div style={{
              background: 'rgba(0,0,0,0.3)',
              borderRadius: '12px',
              padding: '1rem 1.15rem',
              border: '1px solid rgba(59,130,246,0.15)',
            }}>
              {/* Step indicators */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.45rem', marginBottom: '0.85rem' }}>
                {STEPS.map((step, i) => {
                  const isActive = i === currentStep;
                  const isDone = i < currentStep;
                  const StepIcon = step.icon;
                  return (
                    <div
                      key={i}
                      style={{
                        display: 'flex', alignItems: 'center', gap: '0.55rem',
                        opacity: isDone ? 0.5 : isActive ? 1 : 0.25,
                        transition: 'all 0.4s ease',
                      }}
                    >
                      <div style={{
                        width: '22px', height: '22px', borderRadius: '6px',
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        background: isDone ? 'rgba(74,222,128,0.2)' : isActive ? 'rgba(59,130,246,0.2)' : 'transparent',
                        border: isDone ? '1px solid rgba(74,222,128,0.3)' : isActive ? '1px solid rgba(59,130,246,0.3)' : '1px solid rgba(255,255,255,0.08)',
                        transition: 'all 0.3s',
                      }}>
                        {isDone
                          ? <CheckCircle2 size={12} color="#4ADE80" />
                          : isActive
                            ? <StepIcon size={12} color="#60A5FA" style={{ animation: 'spin 2s linear infinite' }} />
                            : <StepIcon size={12} color="rgba(255,255,255,0.2)" />
                        }
                      </div>
                      <span style={{
                        fontSize: '0.78rem',
                        fontWeight: isActive ? 700 : 500,
                        color: isDone ? '#86EFAC' : isActive ? '#93C5FD' : 'rgba(255,255,255,0.2)',
                        transition: 'all 0.3s',
                      }}>
                        {step.label}
                        {isDone && ' ✓'}
                      </span>
                    </div>
                  );
                })}
              </div>

              {/* Progress bar */}
              <div style={{
                height: '4px', background: 'rgba(255,255,255,0.06)', borderRadius: '2px',
                overflow: 'hidden',
              }}>
                <div style={{
                  height: '100%',
                  width: `${STEPS[currentStep]?.pct || 0}%`,
                  background: 'linear-gradient(90deg, #3B82F6, #818CF8)',
                  borderRadius: '2px',
                  transition: 'width 0.8s ease-out',
                }} />
              </div>
            </div>
          </div>
        )}

        {/* ── Drop Zone ── */}
        {!isLoading && (
          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => !file && fileInputRef.current?.click()}
            style={{
              border: `2px dashed ${isDragActive ? '#3B82F6' : file ? 'rgba(74,222,128,0.4)' : 'rgba(255,255,255,0.1)'}`,
              borderRadius: '14px',
              padding: file ? '1.25rem 1rem' : '2rem 1rem',
              display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
              background: isDragActive
                ? 'rgba(59,130,246,0.08)'
                : file ? 'rgba(74,222,128,0.03)' : 'rgba(255,255,255,0.015)',
              transition: 'all 0.3s ease',
              cursor: file ? 'default' : 'pointer',
              minHeight: file ? 'auto' : '150px',
              position: 'relative',
              overflow: 'hidden',
            }}
          >
            <input type="file" ref={fileInputRef} onChange={handleInput} accept={ACCEPTED_EXCEL_EXTENSIONS.join(',')} style={{ display: 'none' }} />

            {/* Drag active shimmer effect */}
            {isDragActive && (
              <div style={{
                position: 'absolute', inset: 0,
                background: 'linear-gradient(90deg, transparent, rgba(59,130,246,0.05), transparent)',
                backgroundSize: '200% 100%',
                animation: 'shimmer 1.5s infinite',
                pointerEvents: 'none',
              }} />
            )}

            {!file ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '0.7rem', textAlign: 'center', zIndex: 1 }}>
                <div style={{
                  background: isDragActive ? 'rgba(59,130,246,0.15)' : 'rgba(255,255,255,0.06)',
                  borderRadius: '14px', padding: '1rem',
                  transition: 'all 0.3s',
                  border: isDragActive ? '1px solid rgba(59,130,246,0.2)' : '1px solid transparent',
                }}>
                  <UploadCloud
                    size={30}
                    color={isDragActive ? '#60A5FA' : 'rgba(255,255,255,0.35)'}
                    style={{ transition: 'color 0.3s' }}
                  />
                </div>
                <div>
                  <p style={{
                    color: isDragActive ? '#93C5FD' : 'rgba(255,255,255,0.55)',
                    fontWeight: 700, margin: '0 0 0.25rem', fontSize: '0.92rem',
                    transition: 'color 0.3s',
                  }}>
                    {isDragActive ? 'Suelta tu archivo aquí' : 'Arrastra y suelta tu archivo Excel'}
                  </p>
                  <p style={{ color: 'rgba(255,255,255,0.28)', fontSize: '0.78rem', margin: 0 }}>
                    o{' '}
                    <span style={{ color: '#60A5FA', fontWeight: 700, textDecoration: 'underline', cursor: 'pointer' }}>
                      explora tus archivos
                    </span>
                    {' '}· .xlsx / .xls / .xlsm
                  </p>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.85rem', width: '100%' }}>
                <div style={{
                  background: 'rgba(74,222,128,0.12)',
                  borderRadius: '12px', padding: '0.7rem',
                  border: '1px solid rgba(74,222,128,0.15)',
                  flexShrink: 0,
                }}>
                  <FileSpreadsheet size={24} color="#4ADE80" />
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <p style={{
                    color: 'white', fontWeight: 700, margin: '0 0 0.15rem', fontSize: '0.88rem',
                    overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                  }}>
                    {file.name}
                  </p>
                  <p style={{ color: 'rgba(255,255,255,0.3)', fontSize: '0.73rem', margin: 0 }}>
                    {formatFileSize(file.size)} · Límite {formatFileSize(MAX_EXCEL_UPLOAD_BYTES)}
                  </p>
                </div>
                <button
                  onClick={(e) => { e.stopPropagation(); setFile(null); setStatus('idle'); setErrorMessage(''); setStats(null); }}
                  style={{
                    color: 'rgba(255,255,255,0.35)', background: 'rgba(255,255,255,0.05)',
                    border: '1px solid rgba(255,255,255,0.08)', cursor: 'pointer',
                    fontSize: '0.72rem', fontWeight: 600, padding: '0.35rem 0.65rem',
                    borderRadius: '8px', transition: 'all 0.2s',
                    display: 'flex', alignItems: 'center', gap: '0.3rem',
                    flexShrink: 0,
                  }}
                >
                  <RefreshCw size={12} />
                  Cambiar
                </button>
              </div>
            )}
          </div>
        )}

        {/* ── Status Messages ── */}
        {status === 'error' && (
          <div
            style={{
              marginTop: '0.875rem', padding: '0.7rem 1rem',
              background: 'rgba(239,68,68,0.08)',
              border: '1px solid rgba(239,68,68,0.2)',
              borderRadius: '10px', display: 'flex', gap: '0.6rem', alignItems: 'flex-start',
            }}
            className="animate-fade-in-up"
          >
            <AlertCircle size={16} color="#F87171" style={{ flexShrink: 0, marginTop: '2px' }} />
            <div>
              <p style={{ color: '#FCA5A5', fontSize: '0.82rem', margin: '0 0 0.2rem', fontWeight: 600 }}>Error en el proceso</p>
              <p style={{ color: 'rgba(252,165,165,0.7)', fontSize: '0.75rem', margin: 0 }}>{errorMessage}</p>
            </div>
          </div>
        )}

        {status === 'success' && (
          <div
            style={{
              marginTop: '0.875rem', padding: '0.85rem 1rem',
              background: 'rgba(74,222,128,0.06)',
              border: '1px solid rgba(74,222,128,0.2)',
              borderRadius: '12px',
            }}
            className="animate-fade-in-up"
          >
            <div style={{ display: 'flex', gap: '0.6rem', alignItems: 'center', marginBottom: '0.45rem' }}>
              <div style={{
                background: 'rgba(74,222,128,0.15)',
                borderRadius: '8px', padding: '0.35rem',
              }}>
                <CheckCircle2 size={16} color="#4ADE80" />
              </div>
              <p style={{ color: '#86EFAC', fontSize: '0.88rem', margin: 0, fontWeight: 700 }}>
                ¡Presentación generada exitosamente!
              </p>
            </div>
            <div style={{
              display: 'flex', gap: '0.85rem', marginTop: '0.5rem',
              paddingTop: '0.5rem', borderTop: '1px solid rgba(74,222,128,0.1)',
            }}>
              {[
                { icon: Layers, label: 'Modo', value: stats ? MODE_OPTIONS.find((option) => option.id === stats.mode)?.label ?? activeModeInfo.label : activeModeInfo.label },
                { icon: Sparkles, label: 'Duración', value: stats ? `${stats.duration.toFixed(1)} s` : 'Completada' },
                { icon: Shield, label: 'Archivo', value: stats ? stats.fileName : (file?.name ?? 'Procesado') },
              ].map(({ icon: Ic, label, value }) => (
                <div key={label} style={{ display: 'flex', alignItems: 'center', gap: '0.35rem' }}>
                  <Ic size={12} color="rgba(134,239,172,0.5)" />
                  <span style={{ color: 'rgba(255,255,255,0.3)', fontSize: '0.68rem' }}>{label}:</span>
                  <span style={{ color: '#86EFAC', fontSize: '0.72rem', fontWeight: 700 }}>{value}</span>
                </div>
              ))}
            </div>
            <button
              onClick={() => { setStatus('idle'); setStats(null); }}
              style={{
                marginTop: '0.6rem', padding: '0.4rem 0.75rem',
                background: 'rgba(74,222,128,0.1)', border: '1px solid rgba(74,222,128,0.2)',
                borderRadius: '8px', color: '#86EFAC', cursor: 'pointer',
                fontSize: '0.75rem', fontWeight: 600,
                display: 'flex', alignItems: 'center', gap: '0.3rem',
                transition: 'all 0.2s',
              }}
            >
              <RefreshCw size={12} />
              Generar otra presentación
            </button>
          </div>
        )}

        {status === 'organized' && (
          <div
            style={{
              marginTop: '0.875rem', padding: '0.8rem 1rem',
              background: 'rgba(74,222,128,0.06)',
              border: '1px solid rgba(74,222,128,0.2)',
              borderRadius: '12px',
            }}
            className="animate-fade-in-up"
          >
            <div style={{ display: 'flex', gap: '0.6rem', alignItems: 'center', marginBottom: '0.3rem' }}>
              <CheckCircle2 size={16} color="#4ADE80" />
              <p style={{ color: '#86EFAC', fontSize: '0.82rem', margin: 0, fontWeight: 700 }}>
                Excel organizado en modo <strong>{activeModeInfo.label}</strong> y descargado
              </p>
            </div>
            <div style={{
              display: 'flex', alignItems: 'center', gap: '0.4rem',
              marginLeft: '1.4rem', marginTop: '0.3rem',
            }}>
              <ChevronRight size={12} color="rgba(255,255,255,0.3)" />
              <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: '0.73rem', margin: 0 }}>
                Usa el archivo descargado en <strong style={{ color: 'rgba(255,255,255,0.5)' }}>Generar PPTX</strong> para el mejor resultado.
              </p>
            </div>
          </div>
        )}

        {/* ── Action Button ── */}
        <div style={{ marginTop: '1.25rem' }}>
          {activeTab === 'generate' ? (
            <button
              id="btn-generate-pptx"
              onClick={handleGenerate}
              disabled={!file || isLoading}
              style={{
                width: '100%', padding: '0.95rem', border: 'none', borderRadius: '12px',
                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem',
                fontWeight: 700, fontSize: '0.92rem',
                cursor: (!file || isLoading) ? 'not-allowed' : 'pointer',
                transition: 'all 0.3s ease',
                background: (!file || isLoading)
                  ? 'rgba(255,255,255,0.05)'
                  : 'linear-gradient(135deg, #1E40AF 0%, #3B82F6 50%, #6366F1 100%)',
                color: (!file || isLoading) ? 'rgba(255,255,255,0.2)' : 'white',
                boxShadow: (!file || isLoading)
                  ? 'none'
                  : '0 8px 28px rgba(59,130,246,0.4), 0 0 0 1px rgba(59,130,246,0.2) inset',
                letterSpacing: '-0.01em',
              }}
            >
              {isLoading
                ? <><Loader2 size={18} style={{ animation: 'spin 1s linear infinite' }} /> Generando presentación...</>
                : <><Sparkles size={18} /> Generar PowerPoint Inteligente</>
              }
            </button>
          ) : (
            <button
              id="btn-organize-excel"
              onClick={handleOrganize}
              disabled={!file || isLoading}
              style={{
                width: '100%', padding: '0.95rem', border: 'none', borderRadius: '12px',
                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem',
                fontWeight: 700, fontSize: '0.92rem',
                cursor: (!file || isLoading) ? 'not-allowed' : 'pointer',
                transition: 'all 0.3s ease',
                background: (!file || isLoading)
                  ? 'rgba(255,255,255,0.05)'
                  : orgMode === 'charts'
                    ? 'linear-gradient(135deg, #155E75 0%, #0891B2 100%)'
                    : orgMode === 'tables'
                      ? 'linear-gradient(135deg, #065F46 0%, #059669 100%)'
                      : 'linear-gradient(135deg, #3B0764 0%, #7C3AED 100%)',
                color: (!file || isLoading) ? 'rgba(255,255,255,0.2)' : 'white',
                boxShadow: (!file || isLoading) ? 'none' : `0 8px 28px ${activeModeInfo.glow}`,
                letterSpacing: '-0.01em',
              }}
            >
              {isLoading
                ? <><Loader2 size={18} style={{ animation: 'spin 1s linear infinite' }} /> Organizando en modo {activeModeInfo.label}...</>
                : <><Settings2 size={18} /> Organizar Excel ({activeModeInfo.label})</>
              }
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
