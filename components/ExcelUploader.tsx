"use client";

import React, { useState, useRef, useCallback, useEffect } from 'react';
import {
  UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle,
  Loader2, Settings2, BarChart2, Table2, LayoutDashboard,
  Sparkles, Download, RefreshCw, ChevronRight, Shield, Layers,
  TrendingUp, PieChart, Wand2, Eye,
} from 'lucide-react';
import type { LucideIcon } from 'lucide-react';
import { ACCEPTED_EXCEL_EXTENSIONS, MAX_EXCEL_UPLOAD_BYTES, validateExcelUpload, validateExcelContents } from '@/utils/excel-file';
import { generatePowerPointFromExcel } from '@/utils/pptx-helper';
import { autoOrganizeExcel, OrganizerMode } from '@/utils/excel-organizer';

type Tab = 'generate' | 'organize';
type Status = 'idle' | 'processing' | 'success' | 'organized' | 'previewed' | 'error';

interface GenerationStats {
  duration: number;
  mode: OrganizerMode;
  fileName: string;
}

interface BackendHealth {
  ok: boolean;
  message: string;
}

interface PreviewSlide {
  type?: string;
  title?: string;
  subtitle?: string;
  content?: unknown;
}

interface ModeOption {
  id: OrganizerMode;
  label: string;
  desc: string;
  Icon: LucideIcon;
  color: string;
  glow: string;
  recommended?: boolean;
}

const STEPS = [
  { label: 'Analizando estructura del Excel', icon: Layers, pct: 20 },
  { label: 'Detectando tipos de datos y columnas', icon: TrendingUp, pct: 40 },
  { label: 'Construyendo diapositivas inteligentes', icon: PieChart, pct: 70 },
  { label: 'Renderizando presentación final', icon: Sparkles, pct: 90 },
  { label: 'Generando archivo y descargando', icon: Download, pct: 100 },
];

const PRIMARY_MODE_OPTIONS: ModeOption[] = [
  {
    id: 'mixed', label: 'Mixto',
    desc: 'La opción más segura: balancea gráficos y tablas automáticamente.',
    Icon: LayoutDashboard, color: '#7C3AED', glow: 'rgba(124,58,237,0.35)', recommended: true,
  },
  {
    id: 'charts', label: 'Gráficos',
    desc: 'Prioriza charts y lecturas visuales cuando el Excel tiene buena base numérica.',
    Icon: BarChart2, color: '#0891B2', glow: 'rgba(8,145,178,0.35)',
  },
  {
    id: 'tables', label: 'Tablas',
    desc: 'Enfoca la salida en tablas claras, comparables y fáciles de revisar.',
    Icon: Table2, color: '#059669', glow: 'rgba(5,150,105,0.35)',
  },
];

const ADVANCED_MODE_OPTION: ModeOption = {
  id: 'boardroom',
  label: 'Ejecutivo avanzado',
  desc: 'Pensado para riesgos, semáforos, comparativos y señales tipo comité.',
  Icon: Sparkles,
  color: '#EA580C',
  glow: 'rgba(234,88,12,0.35)',
};

const STEP_TABS: { id: Tab; title: string; short: string; hint: string }[] = [
  {
    id: 'organize',
    title: 'Paso 1. Organiza tu Excel',
    short: 'Organizar Excel',
    hint: 'Recomendado antes de generar el PowerPoint.',
  },
  {
    id: 'generate',
    title: 'Paso 2. Genera el PowerPoint',
    short: 'Generar PowerPoint',
    hint: 'Usa idealmente el archivo organizado descargado en el paso 1.',
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
  const [activeTab, setActiveTab] = useState<Tab>('organize');
  const [orgMode, setOrgMode] = useState<OrganizerMode>('mixed');
  const [currentStep, setCurrentStep] = useState(0);
  const [stats, setStats] = useState<GenerationStats | null>(null);
  const [previewSlides, setPreviewSlides] = useState<PreviewSlide[]>([]);
  const [backendHealth, setBackendHealth] = useState<BackendHealth | null>(null);
  const [isCheckingHealth, setIsCheckingHealth] = useState(false);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [justAutoAdvanced, setJustAutoAdvanced] = useState(false);
  const [viewportHeight, setViewportHeight] = useState(960);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (typeof window === 'undefined') return undefined;
    const syncViewportHeight = () => setViewportHeight(window.innerHeight || 960);
    syncViewportHeight();
    window.addEventListener('resize', syncViewportHeight);
    return () => window.removeEventListener('resize', syncViewportHeight);
  }, []);

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

  useEffect(() => {
    if (!justAutoAdvanced) return;
    const timeoutId = setTimeout(() => setJustAutoAdvanced(false), 2400);
    return () => clearTimeout(timeoutId);
  }, [justAutoAdvanced]);

  const setValidFile = useCallback(async (f: File) => {
    setStatus('processing');
    setErrorMessage('');
    
    const validationError = validateSelectedFile(f);
    if (validationError) {
      setFile(null);
      setStatus('error');
      setErrorMessage(validationError);
      setJustAutoAdvanced(false);
      return;
    }

    try {
      const contentError = await validateExcelContents(f);
      if (contentError) {
        setFile(null);
        setStatus('error');
        setErrorMessage(contentError);
        setJustAutoAdvanced(false);
        return;
      }
      
      setFile(f);
      setStatus('idle');
      setErrorMessage('');
      setStats(null);
      setPreviewSlides([]);
      setJustAutoAdvanced(false);
    } catch (err) {
      setFile(null);
      setStatus('error');
      setErrorMessage('Error validando el contenido del archivo.');
      setJustAutoAdvanced(false);
    }
  }, []);

  const resetMessages = useCallback(() => {
    setStatus('idle');
    setErrorMessage('');
    setStats(null);
    setPreviewSlides([]);
    setJustAutoAdvanced(false);
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

  const handlePreview = async () => {
    if (!file) return;
    setStatus('processing');
    setErrorMessage('');
    setStats(null);
    setPreviewSlides([]);

    try {
      const formData = new FormData();
      formData.append('file', file);

      const response = await fetch('/api/advanced-generate', {
        method: 'POST',
        body: formData,
      });

      const payload = await response.json().catch(() => null);
      if (!response.ok) {
        throw new Error(payload?.error || 'No se pudo revisar la estructura del archivo.');
      }

      const slides = Array.isArray(payload?.slides) ? (payload.slides as PreviewSlide[]) : [];
      if (!slides.length) {
        throw new Error('La revisión no devolvió resultados.');
      }

      setPreviewSlides(slides);
      setStatus('previewed');
    } catch (err: unknown) {
      console.error('Error generating preview:', err);
      setStatus('error');
      setErrorMessage(getErrorMessage(err, 'Error al revisar la estructura del análisis.'));
    }
  };

  const handleOrganize = async () => {
    if (!file) return;
    setStatus('processing');
    setErrorMessage('');
    setPreviewSlides([]);
    try {
      const blob = await autoOrganizeExcel(file, orgMode);
      const organizedFileName = `ORGANIZADO_${orgMode.toUpperCase()}_${file.name}`;
      const organizedFile = new File(
        [blob],
        organizedFileName,
        { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      );
      setFile(organizedFile);
      setActiveTab('generate');
      setJustAutoAdvanced(true);
      setStatus('organized');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = organizedFileName;
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

  const handleHealthCheck = async () => {
    setIsCheckingHealth(true);
    try {
      const response = await fetch('/api/health', {
        method: 'GET',
        cache: 'no-store',
      });
      const payload = await response.json().catch(() => null);
      setBackendHealth({
        ok: Boolean(payload?.ok && response.ok),
        message: String(payload?.message || 'No se pudo validar el backend.'),
      });
    } catch (err: unknown) {
      setBackendHealth({
        ok: false,
        message: getErrorMessage(err, 'No se pudo validar el backend.'),
      });
    } finally {
      setIsCheckingHealth(false);
    }
  };

  const isLoading = status === 'processing';
  const allModes = [...PRIMARY_MODE_OPTIONS, ADVANCED_MODE_OPTION];
  const activeModeInfo = allModes.find((m) => m.id === orgMode) ?? PRIMARY_MODE_OPTIONS[0];
  const visibleModes = showAdvanced ? allModes : PRIMARY_MODE_OPTIONS;
  const currentTabInfo = STEP_TABS.find((step) => step.id === activeTab) ?? STEP_TABS[0];
  const isReadyToGenerate = status === 'organized' && activeTab === 'generate' && Boolean(file);
  const isCompactViewport = viewportHeight <= 920;
  const isShortViewport = viewportHeight <= 840;

  return (
    <div
      style={{
        background: 'rgba(255,255,255,0.03)',
        backdropFilter: 'blur(24px) saturate(180%)',
        borderRadius: '24px',
        border: '1px solid rgba(255,255,255,0.08)',
        overflow: 'hidden',
        boxShadow: '0 24px 80px rgba(0,0,0,0.5), 0 0 1px rgba(255,255,255,0.1) inset',
      }}
      className="animate-scale-in"
    >
      <div
        style={{
          padding: isShortViewport ? '0.8rem 0.85rem 0.7rem' : (isCompactViewport ? '0.9rem 0.95rem 0.75rem' : '1rem 1rem 0.85rem'),
          borderBottom: '1px solid rgba(255,255,255,0.05)',
          background: 'linear-gradient(180deg, rgba(255,255,255,0.03), rgba(255,255,255,0.01))',
        }}
      >
        <div
          style={{
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'space-between',
            gap: '0.9rem',
            marginBottom: '1rem',
            flexWrap: 'wrap',
          }}
        >
          <div>
            <p style={{ color: '#93C5FD', fontSize: '0.74rem', fontWeight: 800, letterSpacing: '0.08em', textTransform: 'uppercase', margin: '0 0 0.35rem' }}>
              Flujo recomendado
            </p>
            <h2 style={{ color: 'white', fontSize: isShortViewport ? '0.92rem' : (isCompactViewport ? '0.98rem' : '1.02rem'), margin: '0 0 0.18rem', fontWeight: 800 }}>
              Organiza primero. Genera después.
            </h2>
            {!isShortViewport && (
              <p style={{ color: 'rgba(255,255,255,0.38)', fontSize: '0.75rem', margin: 0, lineHeight: 1.45, maxWidth: '30rem' }}>
                Menos pasos visibles, misma lógica guiada.
              </p>
            )}
          </div>
          <div
            style={{
              padding: isShortViewport ? '0.42rem 0.6rem' : '0.55rem 0.75rem',
              borderRadius: '999px',
              border: '1px solid rgba(74,222,128,0.18)',
              background: 'rgba(74,222,128,0.08)',
              color: '#86EFAC',
              fontSize: '0.7rem',
              fontWeight: 700,
            }}
          >
            Modo recomendado: Mixto
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: isShortViewport ? '0.5rem' : '0.7rem' }}>
          {STEP_TABS.map((step, index) => {
            const active = activeTab === step.id;
            return (
              <button
                key={step.id}
                type="button"
                onClick={() => {
                  setActiveTab(step.id);
                  resetMessages();
                }}
                className={step.id === 'generate' && justAutoAdvanced ? 'animate-step-focus' : undefined}
                style={{
                  textAlign: 'left',
                  padding: isShortViewport ? '0.72rem 0.8rem' : (isCompactViewport ? '0.82rem 0.9rem' : '0.95rem 1rem'),
                  borderRadius: '18px',
                  border: active ? '1px solid rgba(96,165,250,0.4)' : '1px solid rgba(255,255,255,0.07)',
                  background: active ? 'rgba(59,130,246,0.12)' : 'rgba(255,255,255,0.025)',
                  color: 'white',
                  cursor: 'pointer',
                  transition: 'all 0.25s ease',
                  boxShadow: active ? '0 0 0 1px rgba(59,130,246,0.08) inset' : 'none',
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.6rem', marginBottom: '0.35rem' }}>
                  <div
                    style={{
                      width: '1.8rem',
                      height: '1.8rem',
                      borderRadius: '999px',
                      background: active ? 'linear-gradient(135deg, #1D4ED8, #60A5FA)' : 'rgba(255,255,255,0.08)',
                      color: active ? 'white' : 'rgba(255,255,255,0.55)',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      fontWeight: 800,
                      fontSize: '0.82rem',
                    }}
                  >
                    {index + 1}
                  </div>
                  <div>
                    <p style={{ color: active ? 'white' : 'rgba(255,255,255,0.72)', fontSize: isShortViewport ? '0.78rem' : '0.84rem', fontWeight: 800, margin: 0 }}>
                      {step.short}
                    </p>
                    {!isShortViewport && (
                      <p style={{ color: active ? '#BFDBFE' : 'rgba(255,255,255,0.30)', fontSize: '0.68rem', margin: 0 }}>
                        {step.hint}
                      </p>
                    )}
                  </div>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      <div style={{ padding: isShortViewport ? '0.92rem' : (isCompactViewport ? '1.08rem' : '1.35rem') }}>
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '0.6rem',
            marginBottom: isShortViewport ? '0.72rem' : '1rem',
            padding: isShortViewport ? '0.6rem 0.72rem' : '0.75rem 0.85rem',
            borderRadius: '14px',
            background: activeTab === 'organize' ? 'rgba(74,222,128,0.07)' : 'rgba(59,130,246,0.08)',
            border: activeTab === 'organize' ? '1px solid rgba(74,222,128,0.16)' : '1px solid rgba(59,130,246,0.18)',
          }}
        >
          {activeTab === 'organize' ? <Wand2 size={16} color="#86EFAC" /> : <Sparkles size={16} color="#93C5FD" />}
          <div>
            <p style={{ color: 'white', fontSize: '0.82rem', fontWeight: 700, margin: '0 0 0.1rem' }}>
              {currentTabInfo.title}
            </p>
            {!isShortViewport && (
              <p style={{ color: 'rgba(255,255,255,0.45)', fontSize: '0.73rem', margin: 0, lineHeight: 1.45 }}>
                {activeTab === 'organize'
                  ? 'Descarga un Excel preparado para que el sistema lo interprete mejor antes de crear las diapositivas.'
                  : 'Sube el Excel final que quieres convertir. Si vienes del paso 1, usa el archivo organizado que acabas de descargar.'}
              </p>
            )}
          </div>
        </div>

        {isLoading && (
          <div style={{ marginBottom: '1rem' }} className="animate-fade-in">
            <div
              style={{
                background: 'rgba(0,0,0,0.3)',
                borderRadius: '16px',
                padding: '1rem 1.05rem',
                border: '1px solid rgba(59,130,246,0.15)',
              }}
            >
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.45rem', marginBottom: '0.85rem' }}>
                {STEPS.map((step, i) => {
                  const isActive = i === currentStep;
                  const isDone = i < currentStep;
                  const StepIcon = step.icon;
                  return (
                    <div
                      key={step.label}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '0.55rem',
                        opacity: isDone ? 0.55 : isActive ? 1 : 0.25,
                        transition: 'all 0.4s ease',
                      }}
                    >
                      <div
                        style={{
                          width: '22px',
                          height: '22px',
                          borderRadius: '6px',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          background: isDone ? 'rgba(74,222,128,0.2)' : isActive ? 'rgba(59,130,246,0.2)' : 'transparent',
                          border: isDone ? '1px solid rgba(74,222,128,0.3)' : isActive ? '1px solid rgba(59,130,246,0.3)' : '1px solid rgba(255,255,255,0.08)',
                        }}
                      >
                        {isDone
                          ? <CheckCircle2 size={12} color="#4ADE80" />
                          : isActive
                            ? <StepIcon size={12} color="#60A5FA" style={{ animation: 'spin 2s linear infinite' }} />
                            : <StepIcon size={12} color="rgba(255,255,255,0.2)" />}
                      </div>
                      <span style={{ fontSize: '0.78rem', fontWeight: isActive ? 700 : 500, color: isDone ? '#86EFAC' : isActive ? '#93C5FD' : 'rgba(255,255,255,0.2)' }}>
                        {step.label}
                      </span>
                    </div>
                  );
                })}
              </div>
              <div style={{ height: '4px', background: 'rgba(255,255,255,0.06)', borderRadius: '2px', overflow: 'hidden' }}>
                <div
                  style={{
                    height: '100%',
                    width: `${STEPS[currentStep]?.pct || 0}%`,
                    background: 'linear-gradient(90deg, #3B82F6, #818CF8)',
                    borderRadius: '2px',
                    transition: 'width 0.8s ease-out',
                  }}
                />
              </div>
            </div>
          </div>
        )}

        {!isLoading && (
          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => !file && fileInputRef.current?.click()}
            style={{
              border: `2px dashed ${isDragActive ? '#3B82F6' : file ? 'rgba(74,222,128,0.4)' : 'rgba(255,255,255,0.12)'}`,
              borderRadius: '18px',
              padding: file ? (isShortViewport ? '0.85rem 0.85rem' : '1.25rem 1rem') : (isShortViewport ? '1rem 0.9rem' : '2rem 1rem'),
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              background: isDragActive
                ? 'rgba(59,130,246,0.08)'
                : file ? 'rgba(74,222,128,0.03)' : 'rgba(255,255,255,0.015)',
              transition: 'all 0.3s ease',
              cursor: file ? 'default' : 'pointer',
              minHeight: file ? 'auto' : (isShortViewport ? '118px' : (isCompactViewport ? '140px' : '170px')),
              position: 'relative',
              overflow: 'hidden',
            }}
          >
            <input type="file" ref={fileInputRef} onChange={handleInput} accept={ACCEPTED_EXCEL_EXTENSIONS.join(',')} style={{ display: 'none' }} />

            {isDragActive && (
              <div
                style={{
                  position: 'absolute',
                  inset: 0,
                  background: 'linear-gradient(90deg, transparent, rgba(59,130,246,0.05), transparent)',
                  backgroundSize: '200% 100%',
                  animation: 'shimmer 1.5s infinite',
                  pointerEvents: 'none',
                }}
              />
            )}

            {!file ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '0.7rem', textAlign: 'center', zIndex: 1 }}>
                <div
                  style={{
                    background: isDragActive ? 'rgba(59,130,246,0.15)' : 'rgba(255,255,255,0.06)',
                    borderRadius: '16px',
                    padding: isShortViewport ? '0.75rem' : '1rem',
                    border: isDragActive ? '1px solid rgba(59,130,246,0.2)' : '1px solid transparent',
                  }}
                >
                  <UploadCloud size={isShortViewport ? 24 : 30} color={isDragActive ? '#60A5FA' : 'rgba(255,255,255,0.35)'} />
                </div>
                <div>
                  <p style={{ color: isDragActive ? '#93C5FD' : 'rgba(255,255,255,0.62)', fontWeight: 700, margin: '0 0 0.28rem', fontSize: isShortViewport ? '0.86rem' : '0.94rem' }}>
                    {isDragActive ? 'Suelta tu archivo aquí' : 'Sube el Excel con el que quieres trabajar'}
                  </p>
                  <p style={{ color: 'rgba(255,255,255,0.3)', fontSize: isShortViewport ? '0.72rem' : '0.78rem', margin: 0, lineHeight: 1.45 }}>
                    Arrastra el archivo o{' '}
                    <span style={{ color: '#60A5FA', fontWeight: 700, textDecoration: 'underline' }}>
                      selecciónalo desde tu equipo
                    </span>
                    {' '}· .xlsx / .xls / .xlsm
                  </p>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.85rem', width: '100%' }}>
                <div
                  style={{
                    background: 'rgba(74,222,128,0.12)',
                    borderRadius: '14px',
                    padding: isShortViewport ? '0.58rem' : '0.75rem',
                    border: '1px solid rgba(74,222,128,0.15)',
                    flexShrink: 0,
                  }}
                >
                  <FileSpreadsheet size={isShortViewport ? 20 : 24} color="#4ADE80" />
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <p style={{ color: 'white', fontWeight: 700, margin: '0 0 0.15rem', fontSize: isShortViewport ? '0.82rem' : '0.9rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {file.name}
                  </p>
                  <p style={{ color: 'rgba(255,255,255,0.34)', fontSize: isShortViewport ? '0.68rem' : '0.74rem', margin: '0 0 0.25rem' }}>
                    {formatFileSize(file.size)} · Límite {formatFileSize(MAX_EXCEL_UPLOAD_BYTES)}
                  </p>
                  {!isShortViewport && (
                    <p style={{ color: 'rgba(255,255,255,0.28)', fontSize: '0.7rem', margin: 0 }}>
                      {activeTab === 'organize'
                        ? 'Este archivo se usará para preparar una versión más limpia y lista para el generador.'
                        : 'Este archivo se convertirá directamente en una presentación PowerPoint.'}
                    </p>
                  )}
                </div>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setFile(null);
                    resetMessages();
                  }}
                  style={{
                    color: 'rgba(255,255,255,0.45)',
                    background: 'rgba(255,255,255,0.05)',
                    border: '1px solid rgba(255,255,255,0.08)',
                    cursor: 'pointer',
                    fontSize: '0.72rem',
                    fontWeight: 600,
                    padding: '0.45rem 0.72rem',
                    borderRadius: '10px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '0.3rem',
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

        <div style={{ marginTop: '1rem', marginBottom: '1rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '0.7rem', marginBottom: '0.55rem', flexWrap: 'wrap' }}>
            <div>
              <p style={{ fontSize: '0.72rem', fontWeight: 800, color: 'rgba(255,255,255,0.48)', margin: '0 0 0.18rem', letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                {activeTab === 'organize' ? 'Cómo quieres preparar el Excel' : 'Cómo quieres que se vea el PowerPoint'}
              </p>
              {!isShortViewport && (
                <p style={{ color: 'rgba(255,255,255,0.32)', fontSize: '0.71rem', margin: 0 }}>
                  Si dudas, usa <strong style={{ color: '#C4B5FD' }}>Mixto</strong>.
                </p>
              )}
            </div>
            <button
              type="button"
              onClick={() => setShowAdvanced((current) => !current)}
              style={{
                padding: '0.45rem 0.7rem',
                borderRadius: '999px',
                border: '1px solid rgba(255,255,255,0.08)',
                background: 'rgba(255,255,255,0.03)',
                color: 'rgba(255,255,255,0.6)',
                cursor: 'pointer',
                fontSize: '0.73rem',
                fontWeight: 700,
              }}
            >
              {showAdvanced ? 'Ocultar opciones avanzadas' : 'Mostrar opciones avanzadas'}
            </button>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: showAdvanced ? 'repeat(2, minmax(0, 1fr))' : 'repeat(3, minmax(0, 1fr))', gap: isShortViewport ? '0.42rem' : '0.55rem' }}>
            {visibleModes.map(({ id, label, desc, Icon, color, glow, recommended }) => {
              const active = orgMode === id;
              return (
                <button
                  key={id}
                  type="button"
                  onClick={() => setOrgMode(id)}
                  style={{
                    padding: isShortViewport ? '0.72rem 0.6rem' : '0.9rem 0.7rem',
                    border: `1.5px solid ${active ? color : 'rgba(255,255,255,0.08)'}`,
                    borderRadius: '16px',
                    cursor: 'pointer',
                    transition: 'all 0.25s ease',
                    background: active ? `rgba(${hexToRgb(color)},0.12)` : 'rgba(255,255,255,0.02)',
                    boxShadow: active ? `0 0 18px ${glow}` : 'none',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'flex-start',
                    gap: '0.45rem',
                    textAlign: 'left',
                    position: 'relative',
                  }}
                >
                  {recommended && (
                    <span
                      style={{
                        position: 'absolute',
                        top: '0.55rem',
                        right: '0.55rem',
                        padding: '0.2rem 0.45rem',
                        borderRadius: '999px',
                        background: 'rgba(124,58,237,0.2)',
                        color: '#C4B5FD',
                        fontSize: '0.62rem',
                        fontWeight: 800,
                        letterSpacing: '0.05em',
                        textTransform: 'uppercase',
                      }}
                    >
                      Recomendado
                    </span>
                  )}
                  <div
                    style={{
                      background: active ? `rgba(${hexToRgb(color)},0.2)` : 'rgba(255,255,255,0.05)',
                      borderRadius: '10px',
                      padding: '0.45rem',
                    }}
                  >
                    <Icon size={18} color={active ? color : 'rgba(255,255,255,0.35)'} />
                  </div>
                  <span style={{ fontSize: isShortViewport ? '0.76rem' : '0.83rem', fontWeight: 800, color: active ? color : 'rgba(255,255,255,0.74)' }}>
                    {label}
                  </span>
                  {!isShortViewport && (
                    <span style={{ fontSize: '0.7rem', color: 'rgba(255,255,255,0.32)', lineHeight: 1.45 }}>
                      {desc}
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </div>

        {showAdvanced && (
          <div
            style={{
              marginBottom: '1rem',
              padding: '0.9rem 1rem',
              borderRadius: '16px',
              background: 'rgba(255,255,255,0.025)',
              border: '1px solid rgba(255,255,255,0.06)',
            }}
            className="animate-fade-in"
          >
            <p style={{ color: 'white', fontSize: '0.77rem', fontWeight: 800, margin: '0 0 0.25rem' }}>
              Herramientas técnicas opcionales
            </p>
            <p style={{ color: 'rgba(255,255,255,0.30)', fontSize: '0.7rem', margin: '0 0 0.7rem', lineHeight: 1.4 }}>
              Solo úsalo si necesitas revisar el estado técnico.
            </p>
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '0.75rem',
                flexWrap: 'wrap',
                padding: '0.75rem 0.85rem',
                background: 'rgba(255,255,255,0.03)',
                border: '1px solid rgba(255,255,255,0.06)',
                borderRadius: '12px',
              }}
            >
              <div style={{ minWidth: 0 }}>
                <p style={{ color: 'rgba(255,255,255,0.65)', fontSize: '0.74rem', margin: '0 0 0.15rem', fontWeight: 700 }}>
                  Estado del backend
                </p>
                <p style={{ color: backendHealth ? (backendHealth.ok ? '#86EFAC' : '#FCA5A5') : 'rgba(255,255,255,0.32)', fontSize: '0.72rem', margin: 0 }}>
                  {backendHealth?.message || 'Aún no revisado'}
                </p>
              </div>
              <button
                type="button"
                onClick={handleHealthCheck}
                disabled={isCheckingHealth || isLoading}
                style={{
                  padding: '0.55rem 0.8rem',
                  background: 'rgba(59,130,246,0.12)',
                  border: '1px solid rgba(59,130,246,0.2)',
                  borderRadius: '10px',
                  color: isCheckingHealth || isLoading ? 'rgba(255,255,255,0.25)' : '#93C5FD',
                  cursor: isCheckingHealth || isLoading ? 'not-allowed' : 'pointer',
                  fontSize: '0.74rem',
                  fontWeight: 700,
                  display: 'flex',
                  alignItems: 'center',
                  gap: '0.35rem',
                }}
              >
                {isCheckingHealth ? <Loader2 size={14} style={{ animation: 'spin 1s linear infinite' }} /> : <Shield size={14} />}
                Revisar backend
              </button>
            </div>
          </div>
        )}

        {status === 'error' && (
          <div
            style={{
              marginTop: '0.875rem',
              padding: '0.8rem 1rem',
              background: 'rgba(239,68,68,0.08)',
              border: '1px solid rgba(239,68,68,0.2)',
              borderRadius: '12px',
              display: 'flex',
              gap: '0.6rem',
              alignItems: 'flex-start',
            }}
            className="animate-fade-in-up"
          >
            <AlertCircle size={16} color="#F87171" style={{ flexShrink: 0, marginTop: '2px' }} />
            <div>
              <p style={{ color: '#FCA5A5', fontSize: '0.82rem', margin: '0 0 0.2rem', fontWeight: 700 }}>
                Hubo un problema en el proceso
              </p>
              <p style={{ color: 'rgba(252,165,165,0.74)', fontSize: '0.75rem', margin: 0, lineHeight: 1.45 }}>
                {errorMessage}
              </p>
            </div>
          </div>
        )}

        {status === 'organized' && (
          <div
            style={{
              marginTop: '0.875rem',
              padding: '0.9rem 1rem',
              background: 'rgba(74,222,128,0.06)',
              border: '1px solid rgba(74,222,128,0.2)',
              borderRadius: '14px',
            }}
            className="animate-fade-in-up"
          >
            <div style={{ display: 'flex', gap: '0.6rem', alignItems: 'center', marginBottom: '0.4rem' }}>
              <CheckCircle2 size={16} color="#4ADE80" />
              <p style={{ color: '#86EFAC', fontSize: '0.86rem', margin: 0, fontWeight: 800 }}>
                Excel organizado, descargado y cargado en el paso 2
              </p>
            </div>
            <p style={{ color: 'rgba(255,255,255,0.42)', fontSize: '0.74rem', margin: '0 0 0 1.45rem', lineHeight: 1.5 }}>
              El archivo organizado ya quedó seleccionado automáticamente. Genera el PowerPoint con el botón principal de abajo.
            </p>
          </div>
        )}

        {status === 'success' && (
          <div
            style={{
              marginTop: '0.875rem',
              padding: '0.95rem 1rem',
              background: 'rgba(74,222,128,0.06)',
              border: '1px solid rgba(74,222,128,0.2)',
              borderRadius: '14px',
            }}
            className="animate-fade-in-up"
          >
            <div style={{ display: 'flex', gap: '0.6rem', alignItems: 'center', marginBottom: '0.45rem' }}>
              <div style={{ background: 'rgba(74,222,128,0.15)', borderRadius: '8px', padding: '0.35rem' }}>
                <CheckCircle2 size={16} color="#4ADE80" />
              </div>
              <p style={{ color: '#86EFAC', fontSize: '0.9rem', margin: 0, fontWeight: 800 }}>
                ¡PowerPoint generado exitosamente!
              </p>
            </div>
            <div style={{ display: 'flex', gap: '0.9rem', flexWrap: 'wrap', marginTop: '0.55rem', paddingTop: '0.55rem', borderTop: '1px solid rgba(74,222,128,0.1)' }}>
              {[
                { icon: Layers, label: 'Modo', value: stats ? allModes.find((option) => option.id === stats.mode)?.label ?? activeModeInfo.label : activeModeInfo.label },
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
              type="button"
              onClick={() => {
                setStatus('idle');
                setStats(null);
              }}
              style={{
                marginTop: '0.7rem',
                padding: '0.45rem 0.78rem',
                background: 'rgba(74,222,128,0.1)',
                border: '1px solid rgba(74,222,128,0.2)',
                borderRadius: '8px',
                color: '#86EFAC',
                cursor: 'pointer',
                fontSize: '0.75rem',
                fontWeight: 700,
                display: 'flex',
                alignItems: 'center',
                gap: '0.3rem',
              }}
            >
              <RefreshCw size={12} />
              Generar otra presentación
            </button>
          </div>
        )}

        {status === 'previewed' && previewSlides.length > 0 && (
          <div
            style={{
              marginTop: '0.875rem',
              padding: '0.95rem 1rem',
              background: 'rgba(96,165,250,0.06)',
              border: '1px solid rgba(96,165,250,0.18)',
              borderRadius: '14px',
            }}
            className="animate-fade-in-up"
          >
            <div style={{ display: 'flex', gap: '0.6rem', alignItems: 'center', marginBottom: '0.55rem' }}>
              <Eye size={16} color="#60A5FA" />
              <p style={{ color: '#93C5FD', fontSize: '0.84rem', margin: 0, fontWeight: 800 }}>
                Revisión rápida lista
              </p>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.45rem' }}>
              {previewSlides.slice(0, isShortViewport ? 3 : 5).map((slide, index) => (
                <div
                  key={`${slide.title || slide.type || 'slide'}-${index}`}
                  style={{
                    padding: isShortViewport ? '0.55rem 0.65rem' : '0.7rem 0.78rem',
                    background: 'rgba(255,255,255,0.03)',
                    border: '1px solid rgba(255,255,255,0.05)',
                    borderRadius: '10px',
                  }}
                >
                  <p style={{ color: 'white', fontSize: '0.78rem', margin: '0 0 0.2rem', fontWeight: 700 }}>
                    {String(index + 1).padStart(2, '0')}. {slide.title || 'Diapositiva sin título'}
                  </p>
                  <p style={{ color: 'rgba(255,255,255,0.34)', fontSize: '0.7rem', margin: 0 }}>
                    Tipo: {slide.type || 'sin tipo'}{slide.subtitle ? ` · ${slide.subtitle}` : ''}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}

        <div
          style={{ marginTop: isShortViewport ? '0.9rem' : '1.2rem' }}
          className={activeTab === 'generate' && justAutoAdvanced ? 'animate-step-arrival' : undefined}
        >
          {activeTab === 'organize' ? (
            <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.8fr', gap: '0.65rem' }}>
              <button
                id="btn-organize-excel"
                onClick={handleOrganize}
                disabled={!file || isLoading}
                style={{
                  width: '100%',
                  padding: '1rem',
                  border: 'none',
                  borderRadius: '14px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '0.5rem',
                  fontWeight: 800,
                  fontSize: '0.94rem',
                  cursor: (!file || isLoading) ? 'not-allowed' : 'pointer',
                  background: (!file || isLoading)
                    ? 'rgba(255,255,255,0.05)'
                    : orgMode === 'charts'
                      ? 'linear-gradient(135deg, #155E75 0%, #0891B2 100%)'
                      : orgMode === 'tables'
                        ? 'linear-gradient(135deg, #065F46 0%, #059669 100%)'
                        : orgMode === 'boardroom'
                          ? 'linear-gradient(135deg, #9A3412 0%, #EA580C 100%)'
                          : 'linear-gradient(135deg, #3B0764 0%, #7C3AED 100%)',
                  color: (!file || isLoading) ? 'rgba(255,255,255,0.2)' : 'white',
                  boxShadow: (!file || isLoading) ? 'none' : `0 8px 28px ${activeModeInfo.glow}`,
                }}
              >
                {isLoading
                  ? <><Loader2 size={18} style={{ animation: 'spin 1s linear infinite' }} /> Organizando archivo...</>
                  : <><Settings2 size={18} /> Organizar y descargar Excel</>}
              </button>
              <button
                type="button"
                onClick={() => {
                  setActiveTab('generate');
                  resetMessages();
                }}
                style={{
                  width: '100%',
                  padding: '1rem',
                  borderRadius: '14px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '0.45rem',
                  fontWeight: 700,
                  fontSize: '0.86rem',
                  cursor: 'pointer',
                  background: 'rgba(255,255,255,0.04)',
                  border: '1px solid rgba(255,255,255,0.08)',
                  color: 'rgba(255,255,255,0.72)',
                }}
              >
                <ChevronRight size={16} />
                Ya quiero generar
              </button>
            </div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.35fr', gap: '0.65rem' }}>
              <button
                id="btn-preview-analysis"
                onClick={handlePreview}
                disabled={!file || isLoading}
                style={{
                  width: '100%',
                  padding: '1rem',
                  borderRadius: '14px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '0.5rem',
                  fontWeight: 700,
                  fontSize: '0.86rem',
                  cursor: (!file || isLoading) ? 'not-allowed' : 'pointer',
                  background: isReadyToGenerate ? 'rgba(255,255,255,0.03)' : 'rgba(255,255,255,0.05)',
                  border: isReadyToGenerate ? '1px solid rgba(255,255,255,0.08)' : '1px solid rgba(96,165,250,0.18)',
                  color: (!file || isLoading) ? 'rgba(255,255,255,0.2)' : (isReadyToGenerate ? 'rgba(255,255,255,0.55)' : '#BFDBFE'),
                }}
              >
                {isLoading
                  ? <><Loader2 size={18} style={{ animation: 'spin 1s linear infinite' }} /> Revisando...</>
                  : <><Eye size={18} /> Revisar estructura</>}
              </button>
              <button
                id="btn-generate-pptx"
                onClick={handleGenerate}
                disabled={!file || isLoading}
                className={isReadyToGenerate ? 'animate-attention-glow' : undefined}
                style={{
                  width: '100%',
                  padding: '1rem',
                  border: 'none',
                  borderRadius: '14px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '0.5rem',
                  fontWeight: 800,
                  fontSize: '0.94rem',
                  cursor: (!file || isLoading) ? 'not-allowed' : 'pointer',
                  background: (!file || isLoading)
                    ? 'rgba(255,255,255,0.05)'
                    : isReadyToGenerate
                      ? 'linear-gradient(135deg, #1D4ED8 0%, #2563EB 35%, #4F46E5 100%)'
                      : 'linear-gradient(135deg, #1E40AF 0%, #3B82F6 50%, #6366F1 100%)',
                  color: (!file || isLoading) ? 'rgba(255,255,255,0.2)' : 'white',
                  boxShadow: (!file || isLoading)
                    ? 'none'
                    : isReadyToGenerate
                      ? '0 14px 34px rgba(37,99,235,0.46), 0 0 0 2px rgba(147,197,253,0.18) inset, 0 0 0 3px rgba(59,130,246,0.12)'
                      : '0 8px 28px rgba(59,130,246,0.4), 0 0 0 1px rgba(59,130,246,0.2) inset',
                  transformOrigin: 'center',
                }}
              >
                {isLoading
                  ? <><Loader2 size={18} style={{ animation: 'spin 1s linear infinite' }} /> Generando presentación...</>
                  : <><Sparkles size={18} /> {isReadyToGenerate ? 'Generar PowerPoint ahora' : 'Generar PowerPoint'}</>}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
