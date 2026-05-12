"use client";

import React, { useState, useRef, useCallback, useEffect } from 'react';
import {
  UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle,
  Loader2, Sparkles, Download, RefreshCw, Shield, Layers,
  TrendingUp, PieChart, Wand2, ChevronDown, ChevronRight,
  Settings2,
} from 'lucide-react';
import { ACCEPTED_EXCEL_EXTENSIONS, MAX_EXCEL_UPLOAD_BYTES, shouldSkipClientContentValidation, validateExcelUpload, validateExcelContents } from '@/utils/excel-file';
import { autoOrganizeExcel, OrganizerMode } from '@/utils/excel-organizer';
import AIControlPanel from './AIControlPanel';
import GenerationProgress from './GenerationProgress';
import AuditModal from './AuditModal';
import PreparePanel from './PreparePanel';
import AdvancedDrawer from './AdvancedDrawer';
import { formatErrorForUser, isPipelineError, PipelineErrorPayload } from '@/utils/error-codes';

type Status = 'idle' | 'processing' | 'success' | 'organized' | 'error';

interface GenerationStats {
  duration: number;
  mode: OrganizerMode;
  fileName: string;
}

interface BackendHealth {
  ok: boolean;
  message: string;
}

interface PresentationContext {
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
}

const STEPS = [
  { label: 'Analizando estructura del Excel', icon: Layers, pct: 20 },
  { label: 'Detectando tipos de datos y columnas', icon: TrendingUp, pct: 40 },
  { label: 'Construyendo diapositivas inteligentes', icon: PieChart, pct: 70 },
  { label: 'Renderizando presentación final', icon: Sparkles, pct: 90 },
  { label: 'Generando archivo y descargando', icon: Download, pct: 100 },
];

const ORGANIZE_MODES: { id: OrganizerMode; label: string; desc: string }[] = [
  { id: 'mixed', label: 'Mixto', desc: 'Balanceado' },
  { id: 'charts', label: 'Gráficos', desc: 'Énfasis visual' },
  { id: 'tables', label: 'Tablas', desc: 'Énfasis tabular' },
  { id: 'boardroom', label: 'Ejecutivo', desc: 'Comité / riesgos' },
];

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getErrorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) return error.message;
  return fallback;
}

function validateSelectedFile(file: File): string | null {
  const validationError = validateExcelUpload(file);
  if (!validationError) return null;
  return validationError
    .replace('El archivo debe ser un Excel valido', 'Por favor sube un archivo Excel válido')
    .replace('El archivo subido no tiene un tipo MIME de Excel permitido.', 'El archivo seleccionado no tiene un formato Excel permitido.')
    .replace('El archivo Excel esta vacio.', 'El archivo está vacío.')
    .replace('El archivo excede el limite permitido', 'El archivo supera el límite');
}

export default function ExcelUploader() {
  const [file, setFile] = useState<File | null>(null);
  const [originalFileName, setOriginalFileName] = useState<string | null>(null);
  const [isDragActive, setIsDragActive] = useState(false);
  const [status, setStatus] = useState<Status>('idle');
  const [errorMessage, setErrorMessage] = useState('');
  const [orgMode, setOrgMode] = useState<OrganizerMode>('mixed');
  const [currentStep, setCurrentStep] = useState(0);
  const [stats, setStats] = useState<GenerationStats | null>(null);
  const [backendHealth, setBackendHealth] = useState<BackendHealth | null>(null);
  const [isCheckingHealth, setIsCheckingHealth] = useState(false);
  const [userPrompt, setUserPrompt] = useState('');
  const [presentationContext, setPresentationContext] = useState<PresentationContext>({
    audience: 'ejecutivos',
    language: 'Español',
    theme: {
      key: 'analitica-moderna',
      name: 'Analitica Moderna',
      primary_hex: '#0F172A',
      accent_hex: '#2563EB',
      text_hex: '#E5E7EB',
      bg_hex: '#F8FAFC',
    },
  });
  const fileInputRef = useRef<HTMLInputElement>(null);

  type ProgressPhase = 'parsing' | 'inventory' | 'planning' | 'validating' | 'rendering' | 'done' | 'error';
  const [progressPhase, setProgressPhase] = useState<ProgressPhase | null>(null);
  const [progressMessage, setProgressMessage] = useState<string>('');
  const [audit, setAudit] = useState<unknown>(null);
  const [showAudit, setShowAudit] = useState(false);
  const [retryError, setRetryError] = useState<PipelineErrorPayload | null>(null);
  const [excludedSlideIndices, setExcludedSlideIndices] = useState<number[]>([]);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [showOrganizer, setShowOrganizer] = useState(false);
  const [showHealth, setShowHealth] = useState(false);

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

  const setValidFile = useCallback(async (f: File) => {
    setStatus('processing');
    setErrorMessage('');

    const validationError = validateSelectedFile(f);
    if (validationError) {
      setFile(null);
      setStatus('error');
      setErrorMessage(validationError);
      return;
    }

    try {
      if (!shouldSkipClientContentValidation(f)) {
        const contentError = await validateExcelContents(f);
        if (contentError) {
          setFile(null);
          setStatus('error');
          setErrorMessage(contentError);
          return;
        }
      }

      setFile(f);
      setOriginalFileName(f.name);
      setStatus('idle');
      setErrorMessage('');
      setStats(null);
      setExcludedSlideIndices([]);
      setShowOrganizer(false);
    } catch {
      setFile(null);
      setStatus('error');
      setErrorMessage('Error validando el contenido del archivo.');
    }
  }, []);

  const resetMessages = useCallback(() => {
    setStatus('idle');
    setErrorMessage('');
    setStats(null);
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
    setRetryError(null);
    setProgressPhase('parsing');
    setProgressMessage('');
    setStatus('processing');
    setErrorMessage('');
    setStats(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('visualMode', orgMode === 'auto' ? 'mixed' : orgMode);
    if (userPrompt.trim()) formData.append('userPrompt', userPrompt.trim());
    formData.append('audience', presentationContext.audience);
    formData.append('language', presentationContext.language);
    formData.append('theme', JSON.stringify(presentationContext.theme));
    if (excludedSlideIndices.length > 0) {
      formData.append('excludeSlideIndices', JSON.stringify(excludedSlideIndices));
    }

    try {
      const res = await fetch('/api/generate-pptx', { method: 'POST', body: formData });
      if (!res.ok || !res.body) {
        let errPayload: PipelineErrorPayload | null = null;
        try {
          const j = await res.json();
          if (j?.error && isPipelineError(j.error)) errPayload = j.error;
        } catch { /* ignore */ }
        setRetryError(errPayload ?? {
          code: 'PYTHON_RUNTIME_ERROR' as const,
          message: 'Error inesperado al iniciar la generación.',
          user_action: 'retry' as const,
        });
        setProgressPhase(null);
        setStatus('error');
        return;
      }

      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const events = buffer.split('\n\n');
        buffer = events.pop() || '';
        for (const ev of events) {
          const m = /^data: (.+)$/m.exec(ev.trim());
          if (!m) continue;
          let payload: { phase: ProgressPhase; message?: string; data?: unknown };
          try {
            payload = JSON.parse(m[1]);
          } catch {
            continue;
          }
          setProgressPhase(payload.phase);
          if (payload.message) setProgressMessage(payload.message);

          if (payload.phase === 'done') {
            const data = payload.data as { downloadToken?: string; filename?: string; audit?: unknown } | undefined;
            if (data?.downloadToken) {
              const dlRes = await fetch(`/api/generate-pptx?token=${encodeURIComponent(data.downloadToken)}`);
              if (dlRes.ok) {
                const blob = await dlRes.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = data.filename || 'presentacion.pptx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
              }
            }
            if (data?.audit) setAudit(data.audit);
            setStatus('success');
            setStats({ duration: 0, mode: orgMode, fileName: file.name });
            setTimeout(() => setProgressPhase(null), 800);
            return;
          }

          if (payload.phase === 'error') {
            const errData = payload.data as PipelineErrorPayload | undefined;
            setRetryError(
              errData && isPipelineError(errData)
                ? errData
                : {
                    code: 'PYTHON_RUNTIME_ERROR' as const,
                    message: payload.message || 'Algo salió mal.',
                    user_action: 'retry' as const,
                  }
            );
            setProgressPhase(null);
            setStatus('error');
            return;
          }
        }
      }
    } catch {
      setRetryError({
        code: 'PYTHON_RUNTIME_ERROR' as const,
        message: 'Algo salió mal en el navegador.',
        user_action: 'retry' as const,
      });
      setProgressPhase(null);
      setStatus('error');
    }
  };

  const handleOrganize = async () => {
    if (!file) return;
    setStatus('processing');
    setErrorMessage('');
    try {
      const blob = await autoOrganizeExcel(file, orgMode);
      const organizedFileName = `ORGANIZADO_${orgMode.toUpperCase()}_${file.name}`;
      const organizedFile = new File(
        [blob],
        organizedFileName,
        { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      );
      setFile(organizedFile);
      setOriginalFileName(file.name);
      setStatus('organized');
      setShowOrganizer(false);
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
      const response = await fetch('/api/health', { method: 'GET', cache: 'no-store' });
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

  return (
    <div className="upl-root">
      <div className="upl-card animate-scale-in">
        {/* ─── Compact header ─── */}
        <div className="upl-header">
          <div className="upl-head-left">
            <p className="upl-eyebrow">Sube tu Excel y obtén un PPT en segundos</p>
            <p className="upl-tagline">La IA lee, decide los slides y descargas listo.</p>
          </div>
          {file && (
            <button
              type="button"
              onClick={() => setShowAdvanced(true)}
              className="upl-adv-trigger press-on-active"
              title="Panel avanzado: prompt, audiencia, tema y sugerencias completas"
            >
              <Settings2 size={14} />
              Avanzado
            </button>
          )}
        </div>

        {/* ─── Body ─── */}
        <div className="upl-body">
          {/* Generation progress (active during PPTX gen) */}
          {isLoading && !progressPhase && (
            <div className="upl-progress-card animate-fade-in">
              <div className="upl-progress-steps">
                {STEPS.map((step, i) => {
                  const isActive = i === currentStep;
                  const isDone = i < currentStep;
                  const StepIcon = step.icon;
                  return (
                    <div
                      key={step.label}
                      className={`upl-progress-step ${isDone ? 'is-done' : isActive ? 'is-active' : 'is-pending'}`}
                    >
                      <div className="upl-progress-icon">
                        {isDone
                          ? <CheckCircle2 size={12} color="#4ADE80" />
                          : isActive
                            ? <StepIcon size={12} color="#60A5FA" style={{ animation: 'spin 2s linear infinite' }} />
                            : <StepIcon size={12} color="rgba(255,255,255,0.2)" />}
                      </div>
                      <span className="upl-progress-label">{step.label}</span>
                    </div>
                  );
                })}
              </div>
              <div className="upl-progress-bar">
                <div
                  className="upl-progress-bar-fill"
                  style={{ width: `${STEPS[currentStep]?.pct || 0}%` }}
                />
              </div>
            </div>
          )}

          {progressPhase && (
            <div className="upl-stream-card animate-fade-in">
              <GenerationProgress currentPhase={progressPhase} message={progressMessage} />
            </div>
          )}

          {/* Dropzone — always visible unless processing */}
          {!isLoading && (
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={() => !file && fileInputRef.current?.click()}
              className={[
                'upl-dropzone',
                isDragActive ? 'is-dragging' : '',
                file ? 'has-file' : 'is-empty',
              ].filter(Boolean).join(' ')}
            >
              <input type="file" ref={fileInputRef} onChange={handleInput} accept={ACCEPTED_EXCEL_EXTENSIONS.join(',')} style={{ display: 'none' }} />

              {isDragActive && <div className="upl-dropzone-shimmer" />}

              {!file ? (
                <div className="upl-dropzone-empty">
                  <div className="upl-dropzone-icon">
                    <UploadCloud size={28} color={isDragActive ? '#60A5FA' : 'rgba(255,255,255,0.35)'} />
                  </div>
                  <div>
                    <p className="upl-dropzone-headline">
                      {isDragActive ? 'Suelta tu archivo aquí' : 'Sube el Excel con el que quieres trabajar'}
                    </p>
                    <p className="upl-dropzone-help">
                      Arrastra o <span className="upl-dropzone-link">selecciónalo</span>{' · '}.xlsx · .xls · .xlsm
                    </p>
                  </div>
                </div>
              ) : (
                <div className="upl-file-row">
                  <div className="upl-file-icon">
                    <FileSpreadsheet size={20} color="#4ADE80" />
                  </div>
                  <div className="upl-file-meta">
                    <p className="upl-file-name">{file.name}</p>
                    <p className="upl-file-size">
                      {formatFileSize(file.size)} · Límite {formatFileSize(MAX_EXCEL_UPLOAD_BYTES)}
                    </p>
                  </div>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      setFile(null);
                      setOriginalFileName(null);
                      resetMessages();
                    }}
                    className="upl-file-change press-on-active"
                    type="button"
                  >
                    <RefreshCw size={12} />
                    Cambiar
                  </button>
                </div>
              )}
            </div>
          )}

          {/* Optional discrete organizer — collapsed by default */}
          {!isLoading && !file && (
            <button
              type="button"
              onClick={() => setShowOrganizer(o => !o)}
              className="upl-organizer-toggle"
            >
              {showOrganizer ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
              <Wand2 size={12} />
              ¿Excel desordenado? Organízalo primero (opcional)
            </button>
          )}

          {!isLoading && !file && showOrganizer && (
            <div className="upl-organizer-box animate-fade-in">
              <p className="upl-organizer-help">
                Genera una versión limpia y normalizada antes del PPT. Útil si el archivo tiene encabezados raros, celdas combinadas o filas vacías.
              </p>
              <div className="upl-organizer-modes">
                {ORGANIZE_MODES.map((m) => (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => setOrgMode(m.id)}
                    className={`upl-org-mode ${orgMode === m.id ? 'is-active' : ''}`}
                  >
                    <span className="upl-org-mode-label">{m.label}</span>
                    <span className="upl-org-mode-desc">{m.desc}</span>
                  </button>
                ))}
              </div>
              <p className="upl-organizer-foot">
                Sube primero un archivo, luego usa este botón para organizarlo.
              </p>
            </div>
          )}

          {!isLoading && file && (
            <button
              type="button"
              onClick={handleOrganize}
              className="upl-organize-cta press-on-active"
              title="Genera y descarga una versión organizada del Excel cargado"
            >
              <Wand2 size={12} />
              Organizar este Excel y descargarlo (opcional)
            </button>
          )}

          {/* PreparePanel — fused onboarding + plan preview */}
          {file && !isLoading && status !== 'success' && status !== 'organized' && (
            <PreparePanel
              file={file}
              userPrompt={userPrompt}
              onPromptChange={setUserPrompt}
              audience={presentationContext.audience}
              language={presentationContext.language}
              theme={presentationContext.theme}
              mode={orgMode}
              onModeChange={(m) => setOrgMode(m as OrganizerMode)}
              onConfirm={(excluded) => {
                setExcludedSlideIndices(excluded);
                setTimeout(() => handleGenerate(), 50);
              }}
              onOpenAdvanced={() => setShowAdvanced(true)}
            />
          )}

          {/* Status banners */}
          {status === 'error' && !retryError && (
            <div className="upl-banner upl-banner-error animate-fade-in-up">
              <AlertCircle size={16} color="#F87171" className="upl-banner-error-icon" />
              <div>
                <p className="upl-banner-error-title">Hubo un problema en el proceso</p>
                <p className="upl-banner-error-msg">{errorMessage}</p>
              </div>
            </div>
          )}

          {retryError && (
            <div className="upl-banner upl-banner-retry animate-fade-in-up">
              <p className="upl-banner-retry-title">{formatErrorForUser(retryError).title}</p>
              <p className="upl-banner-retry-action">{formatErrorForUser(retryError).action}</p>
              <div className="upl-banner-retry-actions">
                <button
                  type="button"
                  onClick={() => { setRetryError(null); handleGenerate(); }}
                  className="upl-btn-retry press-on-active"
                >Reintentar</button>
                <button
                  type="button"
                  onClick={() => setRetryError(null)}
                  className="upl-btn-cancel press-on-active"
                >Cancelar</button>
              </div>
            </div>
          )}

          {status === 'organized' && (
            <div className="upl-banner upl-banner-success animate-fade-in-up">
              <div className="upl-banner-success-row">
                <CheckCircle2 size={16} color="#4ADE80" />
                <p className="upl-banner-success-title">Excel organizado y cargado</p>
              </div>
              <p className="upl-banner-success-sub">El archivo organizado quedó seleccionado. Continúa abajo para generar el PowerPoint.</p>
            </div>
          )}

          {status === 'success' && (
            <div className="upl-banner upl-banner-success-big animate-fade-in-up">
              <div className="upl-banner-success-row">
                <div className="upl-banner-success-iconwrap">
                  <CheckCircle2 size={16} color="#4ADE80" />
                </div>
                <p className="upl-banner-success-title">¡PowerPoint generado exitosamente!</p>
              </div>
              <div className="upl-banner-success-stats">
                {[
                  { icon: Layers, label: 'Modo', value: ORGANIZE_MODES.find(m => m.id === (stats?.mode ?? orgMode))?.label ?? 'Mixto' },
                  { icon: Sparkles, label: 'Archivo', value: stats ? stats.fileName : (file?.name ?? 'Procesado') },
                ].map(({ icon: Ic, label, value }) => (
                  <div key={label} className="upl-stat">
                    <Ic size={12} color="rgba(134,239,172,0.5)" />
                    <span className="upl-stat-label">{label}:</span>
                    <span className="upl-stat-value">{value}</span>
                  </div>
                ))}
              </div>
              <div className="upl-banner-success-actions">
                <button
                  type="button"
                  onClick={() => {
                    setStatus('idle');
                    setStats(null);
                    setAudit(null);
                  }}
                  className="upl-btn-secondary press-on-active"
                >
                  <RefreshCw size={12} />
                  Generar otra
                </button>
                {audit ? (
                  <button
                    type="button"
                    onClick={() => setShowAudit(true)}
                    className="upl-btn-ghost press-on-active"
                  >Ver detalles</button>
                ) : null}
              </div>
            </div>
          )}

          {showAudit && audit ? (
            <AuditModal
              audit={audit as Parameters<typeof AuditModal>[0]['audit']}
              onClose={() => setShowAudit(false)}
            />
          ) : null}

          {/* Tiny system row at the very bottom — collapsed by default */}
          <div className="upl-syslink-row">
            <button
              type="button"
              onClick={() => setShowHealth(h => !h)}
              className="upl-syslink"
            >
              {showHealth ? 'Ocultar' : 'Diagnóstico del backend'}
            </button>
          </div>

          {showHealth && (
            <div className="upl-health-row animate-fade-in">
              <div>
                <p className="upl-health-label">Estado del backend</p>
                <p className={`upl-health-status ${backendHealth ? (backendHealth.ok ? 'is-ok' : 'is-bad') : ''}`}>
                  {backendHealth?.message || 'Aún no revisado'}
                </p>
              </div>
              <button
                type="button"
                onClick={handleHealthCheck}
                disabled={isCheckingHealth || isLoading}
                className="upl-health-btn press-on-active"
              >
                {isCheckingHealth ? <Loader2 size={14} style={{ animation: 'spin 1s linear infinite' }} /> : <Shield size={14} />}
                Revisar
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Advanced drawer (hosts AIControlPanel) */}
      <AdvancedDrawer
        open={showAdvanced}
        onClose={() => setShowAdvanced(false)}
        title="Panel avanzado · IA"
      >
        <AIControlPanel
          file={file}
          isOrganizedFile={Boolean(file?.name?.startsWith('ORGANIZADO_'))}
          originalFileName={originalFileName}
          onPromptChange={setUserPrompt}
          onFocusChange={(f) => setOrgMode(f === 'text' ? 'mixed' : f)}
          onContextChange={setPresentationContext}
        />
      </AdvancedDrawer>

      <style>{`
        .upl-root {
          width: 100%;
          display: flex;
          flex-direction: column;
        }

        .upl-card {
          background: var(--c-bg-elevated);
          backdrop-filter: blur(24px) saturate(180%);
          -webkit-backdrop-filter: blur(24px) saturate(180%);
          border-radius: var(--r-2xl);
          border: 1px solid var(--c-border);
          overflow: hidden;
          box-shadow: var(--shadow-xl), inset 0 0 1px rgba(255, 255, 255, 0.10);
        }

        /* Header */
        .upl-header {
          padding: clamp(0.75rem, 1.4vw, 1rem) clamp(0.95rem, 1.8vw, 1.25rem);
          border-bottom: 1px solid var(--c-divider);
          background: linear-gradient(180deg, rgba(255, 255, 255, 0.03), rgba(255, 255, 255, 0.01));
          display: flex; align-items: center; justify-content: space-between; gap: 1rem;
          flex-wrap: wrap;
        }
        .upl-head-left { min-width: 0; flex: 1; }
        .upl-eyebrow {
          color: white; font-weight: 800; font-size: 0.92rem;
          letter-spacing: -0.01em; line-height: 1.2;
        }
        .upl-tagline {
          color: var(--c-text-tertiary); font-size: 0.72rem;
          margin-top: 2px; line-height: 1.35;
        }
        @media (max-width: 480px) {
          .upl-tagline { display: none; }
        }
        .upl-adv-trigger {
          display: inline-flex; align-items: center; gap: 0.35rem;
          padding: 0.45rem 0.7rem;
          border-radius: var(--r-pill);
          background: rgba(124, 58, 237, 0.12);
          border: 1px solid rgba(124, 58, 237, 0.30);
          color: #C4B5FD;
          font-size: 0.74rem; font-weight: 700;
          transition: all var(--t-base) var(--ease-out);
          flex-shrink: 0;
        }
        .upl-adv-trigger:hover { background: rgba(124, 58, 237, 0.20); color: white; }

        /* Body */
        .upl-body {
          padding: clamp(0.85rem, 1.6vw, 1.25rem);
          display: flex;
          flex-direction: column;
          gap: clamp(0.7rem, 1.4vw, 0.95rem);
        }

        /* Progress (loading state) */
        .upl-progress-card {
          background: rgba(0, 0, 0, 0.30);
          border-radius: var(--r-xl);
          padding: 1rem 1.05rem;
          border: 1px solid rgba(59, 130, 246, 0.15);
        }
        .upl-progress-steps { display: flex; flex-direction: column; gap: 0.45rem; margin-bottom: 0.85rem; }
        .upl-progress-step {
          display: flex; align-items: center; gap: 0.55rem;
          transition: opacity var(--t-slow) var(--ease-out);
        }
        .upl-progress-step.is-done    { opacity: 0.55; }
        .upl-progress-step.is-active  { opacity: 1; }
        .upl-progress-step.is-pending { opacity: 0.25; }
        .upl-progress-icon {
          width: 22px; height: 22px;
          border-radius: 6px;
          display: flex; align-items: center; justify-content: center;
          background: transparent;
          border: 1px solid rgba(255, 255, 255, 0.08);
          flex-shrink: 0;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-progress-step.is-active .upl-progress-icon { background: rgba(59, 130, 246, 0.20); border-color: rgba(59, 130, 246, 0.30); }
        .upl-progress-step.is-done .upl-progress-icon { background: rgba(74, 222, 128, 0.20); border-color: rgba(74, 222, 128, 0.30); }
        .upl-progress-label { font-size: 0.78rem; font-weight: 500; color: var(--c-text-faint); }
        .upl-progress-step.is-active .upl-progress-label { font-weight: 700; color: var(--c-brand-blue-300); }
        .upl-progress-step.is-done .upl-progress-label { color: var(--c-success-400); }
        .upl-progress-bar { height: 4px; background: rgba(255, 255, 255, 0.06); border-radius: 2px; overflow: hidden; }
        .upl-progress-bar-fill {
          height: 100%;
          background: linear-gradient(90deg, var(--c-brand-blue-500), var(--c-brand-indigo-500));
          border-radius: 2px;
          transition: width 0.8s var(--ease-out);
        }
        .upl-stream-card {
          background: rgba(255, 255, 255, 0.04);
          border: 1px solid var(--c-border);
          border-radius: var(--r-lg);
          padding: 1rem;
        }

        /* Dropzone */
        .upl-dropzone {
          border: 2px dashed var(--c-border-strong);
          border-radius: var(--r-xl);
          display: flex; flex-direction: column;
          align-items: center; justify-content: center;
          background: rgba(255, 255, 255, 0.015);
          padding: clamp(0.85rem, 2.5vw, 1.5rem) 1rem;
          min-height: clamp(98px, 14vh, 140px);
          position: relative;
          overflow: hidden;
          transition: all var(--t-base) var(--ease-out);
          cursor: pointer;
        }
        .upl-dropzone:hover:not(.has-file):not(.is-dragging) {
          border-color: rgba(96, 165, 250, 0.30);
          background: rgba(59, 130, 246, 0.04);
        }
        .upl-dropzone.is-dragging {
          border-color: var(--c-brand-blue-500);
          background: rgba(59, 130, 246, 0.08);
        }
        .upl-dropzone.has-file {
          border-color: rgba(74, 222, 128, 0.40);
          background: rgba(74, 222, 128, 0.03);
          padding: clamp(0.65rem, 1.6vw, 0.95rem) 1rem;
          cursor: default;
          min-height: auto;
        }
        .upl-dropzone-shimmer {
          position: absolute; inset: 0;
          background: linear-gradient(90deg, transparent, rgba(59, 130, 246, 0.05), transparent);
          background-size: 200% 100%;
          animation: shimmer 1.5s infinite;
          pointer-events: none;
        }
        .upl-dropzone-empty {
          display: flex; flex-direction: column; align-items: center;
          gap: 0.65rem; text-align: center; z-index: 1;
        }
        .upl-dropzone-icon {
          background: rgba(255, 255, 255, 0.06);
          border-radius: var(--r-lg);
          padding: clamp(0.6rem, 1.4vw, 0.85rem);
          border: 1px solid transparent;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-dropzone.is-dragging .upl-dropzone-icon {
          background: rgba(59, 130, 246, 0.15);
          border-color: rgba(59, 130, 246, 0.20);
          transform: scale(1.05);
        }
        .upl-dropzone-headline {
          color: var(--c-text-secondary); font-weight: 700;
          margin-bottom: 0.2rem;
          font-size: clamp(0.82rem, 1.3vw, 0.9rem);
        }
        .upl-dropzone.is-dragging .upl-dropzone-headline { color: var(--c-brand-blue-300); }
        .upl-dropzone-help {
          color: var(--c-text-muted);
          font-size: clamp(0.7rem, 1.1vw, 0.76rem); line-height: 1.4;
        }
        .upl-dropzone-link {
          color: var(--c-brand-blue-400); font-weight: 700;
          text-decoration: underline; text-underline-offset: 2px;
        }

        /* File row (when file is set) */
        .upl-file-row { display: flex; align-items: center; gap: 0.75rem; width: 100%; }
        .upl-file-icon {
          background: rgba(74, 222, 128, 0.12);
          border-radius: var(--r-lg);
          padding: 0.5rem;
          border: 1px solid rgba(74, 222, 128, 0.18);
          flex-shrink: 0;
        }
        .upl-file-meta { flex: 1; min-width: 0; }
        .upl-file-name {
          color: white; font-weight: 700;
          font-size: clamp(0.8rem, 1.3vw, 0.88rem);
          margin-bottom: 0.12rem;
          overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }
        .upl-file-size { color: var(--c-text-muted); font-size: 0.7rem; }
        .upl-file-change {
          color: var(--c-text-tertiary);
          background: rgba(255, 255, 255, 0.05);
          border: 1px solid var(--c-border);
          font-size: 0.7rem; font-weight: 600;
          padding: 0.4rem 0.65rem;
          border-radius: var(--r-md);
          display: flex; align-items: center; gap: 0.3rem;
          flex-shrink: 0;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-file-change:hover { background: rgba(255, 255, 255, 0.10); color: white; }

        /* Discrete organizer */
        .upl-organizer-toggle {
          display: inline-flex; align-items: center; gap: 0.35rem;
          padding: 0.45rem 0.6rem;
          background: transparent; border: none;
          color: var(--c-text-muted);
          font-size: 0.72rem; font-weight: 600;
          cursor: pointer; align-self: flex-start;
          transition: color var(--t-base) var(--ease-out);
        }
        .upl-organizer-toggle:hover { color: var(--c-text-secondary); }

        .upl-organizer-box {
          padding: 0.85rem 0.95rem;
          background: rgba(255, 255, 255, 0.025);
          border: 1px solid var(--c-divider);
          border-radius: var(--r-lg);
          display: flex; flex-direction: column; gap: 0.6rem;
        }
        .upl-organizer-help { color: var(--c-text-tertiary); font-size: 0.74rem; line-height: 1.45; }
        .upl-organizer-modes {
          display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.4rem;
        }
        @media (min-width: 520px) {
          .upl-organizer-modes { grid-template-columns: repeat(4, 1fr); }
        }
        .upl-org-mode {
          padding: 0.5rem 0.6rem;
          background: rgba(255, 255, 255, 0.02);
          border: 1px solid var(--c-border);
          border-radius: var(--r-md);
          color: var(--c-text-secondary);
          display: flex; flex-direction: column; gap: 2px;
          text-align: left; cursor: pointer;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-org-mode:hover { background: rgba(255, 255, 255, 0.05); color: white; border-color: var(--c-border-strong); }
        .upl-org-mode.is-active {
          background: rgba(74, 222, 128, 0.10);
          border-color: rgba(74, 222, 128, 0.35);
          color: #86EFAC;
        }
        .upl-org-mode-label { font-size: 0.78rem; font-weight: 700; }
        .upl-org-mode-desc { font-size: 0.66rem; opacity: 0.75; }
        .upl-organizer-foot { color: var(--c-text-muted); font-size: 0.68rem; font-style: italic; }

        .upl-organize-cta {
          display: inline-flex; align-items: center; gap: 0.35rem;
          align-self: flex-start;
          padding: 0.45rem 0.7rem;
          background: rgba(74, 222, 128, 0.06);
          border: 1px solid rgba(74, 222, 128, 0.18);
          border-radius: var(--r-pill);
          color: #86EFAC;
          font-size: 0.7rem; font-weight: 700;
          cursor: pointer;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-organize-cta:hover { background: rgba(74, 222, 128, 0.14); color: white; }

        /* Banners */
        .upl-banner { padding: 0.85rem 1rem; border-radius: var(--r-lg); }
        .upl-banner-error {
          background: rgba(239, 68, 68, 0.08);
          border: 1px solid rgba(239, 68, 68, 0.20);
          display: flex; gap: 0.6rem; align-items: flex-start;
        }
        .upl-banner-error-icon { flex-shrink: 0; margin-top: 2px; }
        .upl-banner-error-title { color: var(--c-error-300); font-size: 0.82rem; font-weight: 700; margin-bottom: 0.2rem; }
        .upl-banner-error-msg { color: rgba(252, 165, 165, 0.74); font-size: 0.75rem; line-height: 1.5; }
        .upl-banner-retry {
          background: rgba(220, 38, 38, 0.08);
          border: 1px solid rgba(220, 38, 38, 0.30);
          color: var(--c-error-300);
        }
        .upl-banner-retry-title { font-weight: 700; margin-bottom: 0.3rem; font-size: 0.85rem; }
        .upl-banner-retry-action { font-size: 0.78rem; opacity: 0.8; margin-bottom: 0.6rem; }
        .upl-banner-retry-actions { display: flex; gap: 0.5rem; }
        .upl-btn-retry {
          padding: 0.4rem 0.8rem;
          background: rgba(124, 58, 237, 0.20);
          border: 1px solid rgba(124, 58, 237, 0.40);
          border-radius: 8px;
          color: var(--c-brand-violet-300);
          font-size: 0.78rem; font-weight: 700;
          transition: background var(--t-base) var(--ease-out);
        }
        .upl-btn-retry:hover { background: rgba(124, 58, 237, 0.30); }
        .upl-btn-cancel {
          padding: 0.4rem 0.8rem;
          background: transparent;
          border: 1px solid var(--c-border-strong);
          border-radius: 8px;
          color: var(--c-text-tertiary);
          font-size: 0.78rem;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-btn-cancel:hover { background: rgba(255, 255, 255, 0.05); color: white; }

        .upl-banner-success {
          background: rgba(74, 222, 128, 0.06);
          border: 1px solid rgba(74, 222, 128, 0.20);
        }
        .upl-banner-success-row { display: flex; gap: 0.6rem; align-items: center; margin-bottom: 0.4rem; }
        .upl-banner-success-title { color: var(--c-success-400); font-size: 0.86rem; font-weight: 800; }
        .upl-banner-success-sub { color: var(--c-text-tertiary); font-size: 0.74rem; margin-left: 1.45rem; line-height: 1.5; }
        .upl-banner-success-big {
          background: rgba(74, 222, 128, 0.06);
          border: 1px solid rgba(74, 222, 128, 0.20);
          padding: 0.95rem 1rem;
        }
        .upl-banner-success-iconwrap {
          background: rgba(74, 222, 128, 0.15);
          border-radius: 8px;
          padding: 0.35rem;
        }
        .upl-banner-success-stats {
          display: flex; gap: 0.9rem; flex-wrap: wrap;
          margin-top: 0.55rem; padding-top: 0.55rem;
          border-top: 1px solid rgba(74, 222, 128, 0.10);
        }
        .upl-stat { display: flex; align-items: center; gap: 0.35rem; }
        .upl-stat-label { color: var(--c-text-muted); font-size: 0.68rem; }
        .upl-stat-value { color: var(--c-success-400); font-size: 0.72rem; font-weight: 700; }
        .upl-banner-success-actions { display: flex; align-items: center; gap: 0.6rem; margin-top: 0.7rem; flex-wrap: wrap; }
        .upl-btn-secondary {
          padding: 0.45rem 0.78rem;
          background: rgba(74, 222, 128, 0.10);
          border: 1px solid rgba(74, 222, 128, 0.20);
          border-radius: 8px;
          color: var(--c-success-400);
          font-size: 0.75rem; font-weight: 700;
          display: flex; align-items: center; gap: 0.3rem;
          transition: background var(--t-base) var(--ease-out);
        }
        .upl-btn-secondary:hover { background: rgba(74, 222, 128, 0.18); }
        .upl-btn-ghost {
          background: rgba(255, 255, 255, 0.04);
          border: 1px solid var(--c-border-strong);
          border-radius: 8px;
          padding: 0.4rem 0.7rem;
          color: var(--c-text-tertiary);
          font-size: 0.72rem;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-btn-ghost:hover { background: rgba(255, 255, 255, 0.10); color: white; }

        /* System link row */
        .upl-syslink-row { display: flex; justify-content: flex-end; }
        .upl-syslink {
          background: none; border: none; padding: 0;
          color: var(--c-text-muted);
          font-size: 0.66rem; cursor: pointer;
          transition: color var(--t-base) var(--ease-out);
        }
        .upl-syslink:hover { color: var(--c-text-secondary); text-decoration: underline; }

        .upl-health-row {
          display: flex; align-items: center; justify-content: space-between;
          gap: 0.75rem; flex-wrap: wrap;
          padding: 0.7rem 0.85rem;
          background: rgba(255, 255, 255, 0.03);
          border: 1px solid var(--c-divider);
          border-radius: var(--r-md);
        }
        .upl-health-label { color: var(--c-text-secondary); font-size: 0.72rem; font-weight: 700; margin-bottom: 0.15rem; }
        .upl-health-status { color: var(--c-text-muted); font-size: 0.7rem; }
        .upl-health-status.is-ok { color: var(--c-success-400); }
        .upl-health-status.is-bad { color: var(--c-error-300); }
        .upl-health-btn {
          padding: 0.45rem 0.7rem;
          background: rgba(59, 130, 246, 0.12);
          border: 1px solid rgba(59, 130, 246, 0.20);
          border-radius: var(--r-md);
          color: var(--c-brand-blue-300);
          font-size: 0.72rem; font-weight: 700;
          display: flex; align-items: center; gap: 0.35rem;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-health-btn:hover:not(:disabled) { background: rgba(59, 130, 246, 0.20); }
        .upl-health-btn:disabled { color: var(--c-text-faint); cursor: not-allowed; }
      `}</style>
    </div>
  );
}
