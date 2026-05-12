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
      key: 'socya-institucional',
      name: 'Socya Institucional',
      primary_hex: '#087062',   // teal Socya
      accent_hex: '#69BE28',    // logo-green
      text_hex: '#1A1A1A',
      bg_hex: '#FFFFFF',
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
    } catch (err: unknown) {
      // Pasamos el error original al usuario en vez de un genérico — así puede
      // distinguir entre archivo corrupto, password-protected, etc.
      const detail = err instanceof Error && err.message
        ? `: ${err.message}`
        : '.';
      setFile(null);
      setStatus('error');
      setErrorMessage(`No pudimos leer el contenido del Excel${detail}`);
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
        // Si el server no nos dio un PipelineError tipado, hacemos uno con
        // contexto del status code para que el usuario sepa qué hacer.
        const fallback: PipelineErrorPayload = (() => {
          if (res.status === 503) return {
            code: 'PYTHON_RUNTIME_ERROR',
            message: 'El servicio está temporalmente no disponible (Python no responde).',
            user_action: 'retry',
          };
          if (res.status === 413) return {
            code: 'EXCEL_INVALID',
            message: 'El archivo Excel es demasiado grande para procesar.',
            user_action: 'upload_smaller',
          };
          if (res.status === 400) return {
            code: 'EXCEL_INVALID',
            message: 'El archivo Excel no es válido o falta información.',
            user_action: 'upload_again',
          };
          if (res.status >= 500) return {
            code: 'PYTHON_RUNTIME_ERROR',
            message: `Error del servidor (HTTP ${res.status}). Inténtalo de nuevo.`,
            user_action: 'retry',
          };
          return {
            code: 'PYTHON_RUNTIME_ERROR',
            message: `No se pudo iniciar la generación (HTTP ${res.status}).`,
            user_action: 'retry',
          };
        })();
        setRetryError(errPayload ?? fallback);
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
                // Defer revoke — Safari/Firefox necesitan que la URL siga
                // viva unos ms para que el download inicie correctamente.
                setTimeout(() => URL.revokeObjectURL(url), 1000);
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
    } catch (err: unknown) {
      // TypeError es típicamente fallo de red o conexión cortada;
      // otros errores (parser SSE, etc.) son del cliente — siempre logueamos
      // el detalle a consola para debug y damos un mensaje útil al usuario.
      console.error('[generate-pptx] error en cliente:', err);
      const isNetwork = err instanceof TypeError;
      setRetryError({
        code: 'PYTHON_RUNTIME_ERROR' as const,
        message: isNetwork
          ? 'Se perdió la conexión con el servidor. Verifica tu red e inténtalo de nuevo.'
          : err instanceof Error && err.message
            ? `Error en el navegador: ${err.message}`
            : 'Algo salió mal en el navegador.',
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
      // Defer revoke — algunos navegadores cancelan la descarga si la URL
      // se revoca demasiado pronto tras el click.
      setTimeout(() => URL.revokeObjectURL(url), 1000);
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
                          ? <CheckCircle2 size={12} color="#69BE28" />
                          : isActive
                            ? <StepIcon size={12} color="#087062" style={{ animation: 'spin 2s linear infinite' }} />
                            : <StepIcon size={12} color="rgba(26,26,26,0.25)" />}
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
                    <UploadCloud size={28} color={isDragActive ? '#087062' : 'rgba(26,26,26,0.40)'} />
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
                    <FileSpreadsheet size={20} color="#69BE28" />
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
              <AlertCircle size={16} color="#D43838" className="upl-banner-error-icon" />
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
                <CheckCircle2 size={16} color="#69BE28" />
                <p className="upl-banner-success-title">Excel organizado y cargado</p>
              </div>
              <p className="upl-banner-success-sub">El archivo organizado quedó seleccionado. Continúa abajo para generar el PowerPoint.</p>
            </div>
          )}

          {status === 'success' && (
            <div className="upl-banner upl-banner-success-big animate-fade-in-up">
              <div className="upl-banner-success-row">
                <div className="upl-banner-success-iconwrap">
                  <CheckCircle2 size={16} color="#69BE28" />
                </div>
                <p className="upl-banner-success-title">¡PowerPoint generado exitosamente!</p>
              </div>
              <div className="upl-banner-success-stats">
                {[
                  { icon: Layers, label: 'Modo', value: ORGANIZE_MODES.find(m => m.id === (stats?.mode ?? orgMode))?.label ?? 'Mixto' },
                  { icon: Sparkles, label: 'Archivo', value: stats ? stats.fileName : (file?.name ?? 'Procesado') },
                ].map(({ icon: Ic, label, value }) => (
                  <div key={label} className="upl-stat">
                    <Ic size={12} color="#087062" />
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

        /* ── Card: superficie blanca, borde sutil, sombra suave ── */
        .upl-card {
          background: var(--c-bg-elevated);
          border-radius: var(--r-lg);
          border: 1px solid var(--c-border);
          overflow: hidden;
          box-shadow: var(--shadow-md);
        }

        /* ── Header: white con teal accent, family Jost ── */
        .upl-header {
          padding: clamp(0.85rem, 1.6vw, 1.15rem) clamp(1rem, 1.8vw, 1.35rem);
          border-bottom: 1px solid var(--c-divider);
          background: linear-gradient(180deg, rgba(8, 112, 98, 0.025), transparent);
          display: flex; align-items: center; justify-content: space-between; gap: 1rem;
          flex-wrap: wrap;
        }
        .upl-head-left { min-width: 0; flex: 1; }
        .upl-eyebrow {
          color: var(--c-text-primary);
          font-family: var(--font-heading);
          font-weight: 700;
          font-size: 0.95rem;
          letter-spacing: -0.01em; line-height: 1.2;
        }
        .upl-tagline {
          color: var(--c-text-tertiary); font-size: 0.74rem;
          margin-top: 3px; line-height: 1.4;
        }
        @media (max-width: 480px) {
          .upl-tagline { display: none; }
        }
        .upl-adv-trigger {
          display: inline-flex; align-items: center; gap: 0.4rem;
          padding: 0.5rem 0.85rem;
          border-radius: var(--r-md);
          background: var(--c-bg-elevated);
          border: 1.5px solid var(--c-primary);
          color: var(--c-primary);
          font-family: var(--font-heading);
          font-size: 0.72rem; font-weight: 700;
          letter-spacing: 0.04em; text-transform: uppercase;
          transition: all var(--t-base) var(--ease-out);
          flex-shrink: 0;
        }
        .upl-adv-trigger:hover {
          background: var(--c-accent-green);
          color: var(--c-primary-dark);
          border-color: var(--c-primary-dark);
        }

        /* ── Body ── */
        .upl-body {
          padding: clamp(0.95rem, 1.8vw, 1.35rem);
          display: flex;
          flex-direction: column;
          gap: clamp(0.75rem, 1.4vw, 1rem);
        }

        /* ── Progress (loading state) — light surface ── */
        .upl-progress-card {
          background: var(--c-bg-tinted);
          border-radius: var(--r-lg);
          padding: 1rem 1.05rem;
          border: 1px solid var(--c-border);
        }
        .upl-progress-steps { display: flex; flex-direction: column; gap: 0.5rem; margin-bottom: 0.85rem; }
        .upl-progress-step {
          display: flex; align-items: center; gap: 0.55rem;
          transition: opacity var(--t-slow) var(--ease-out);
        }
        .upl-progress-step.is-done    { opacity: 0.65; }
        .upl-progress-step.is-active  { opacity: 1; }
        .upl-progress-step.is-pending { opacity: 0.35; }
        .upl-progress-icon {
          width: 22px; height: 22px;
          border-radius: var(--r-sm);
          display: flex; align-items: center; justify-content: center;
          background: var(--c-bg-elevated);
          border: 1px solid var(--c-border);
          flex-shrink: 0;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-progress-step.is-active .upl-progress-icon {
          background: rgba(8, 112, 98, 0.12);
          border-color: rgba(8, 112, 98, 0.35);
        }
        .upl-progress-step.is-done .upl-progress-icon {
          background: rgba(105, 190, 40, 0.16);
          border-color: rgba(105, 190, 40, 0.40);
        }
        .upl-progress-label { font-size: 0.78rem; font-weight: 500; color: var(--c-text-tertiary); }
        .upl-progress-step.is-active .upl-progress-label { font-weight: 700; color: var(--c-primary); }
        .upl-progress-step.is-done .upl-progress-label { color: var(--c-logo-green); }
        .upl-progress-bar {
          height: 4px;
          background: var(--c-divider);
          border-radius: 2px;
          overflow: hidden;
        }
        .upl-progress-bar-fill {
          height: 100%;
          background: linear-gradient(90deg, var(--c-primary), var(--c-logo-green));
          border-radius: 2px;
          transition: width 0.8s var(--ease-out);
        }
        .upl-stream-card {
          background: var(--c-bg-tinted);
          border: 1px solid var(--c-border);
          border-radius: var(--r-lg);
          padding: 1rem;
        }

        /* ── Dropzone ── */
        .upl-dropzone {
          border: 2px dashed var(--c-border-strong);
          border-radius: var(--r-lg);
          display: flex; flex-direction: column;
          align-items: center; justify-content: center;
          background: var(--c-bg-neutral);
          padding: clamp(0.95rem, 2.5vw, 1.6rem) 1rem;
          min-height: clamp(98px, 14vh, 140px);
          position: relative;
          overflow: hidden;
          transition: all var(--t-base) var(--ease-out);
          cursor: pointer;
        }
        .upl-dropzone:hover:not(.has-file):not(.is-dragging) {
          border-color: var(--c-primary);
          background: var(--c-bg-tinted);
        }
        .upl-dropzone.is-dragging {
          border-color: var(--c-primary);
          background: rgba(8, 112, 98, 0.08);
        }
        .upl-dropzone.has-file {
          border-style: solid;
          border-color: var(--c-logo-green);
          background: var(--c-accent-green);
          padding: clamp(0.7rem, 1.6vw, 1rem) 1rem;
          cursor: default;
          min-height: auto;
        }
        .upl-dropzone-shimmer {
          position: absolute; inset: 0;
          background: linear-gradient(90deg, transparent, rgba(8, 112, 98, 0.08), transparent);
          background-size: 200% 100%;
          animation: shimmer 1.5s infinite;
          pointer-events: none;
        }
        .upl-dropzone-empty {
          display: flex; flex-direction: column; align-items: center;
          gap: 0.7rem; text-align: center; z-index: 1;
        }
        .upl-dropzone-icon {
          background: var(--c-bg-elevated);
          border-radius: var(--r-md);
          padding: clamp(0.65rem, 1.4vw, 0.9rem);
          border: 1px solid var(--c-border);
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-dropzone.is-dragging .upl-dropzone-icon {
          background: rgba(8, 112, 98, 0.15);
          border-color: var(--c-primary);
          transform: scale(1.05);
        }
        .upl-dropzone-headline {
          color: var(--c-text-primary);
          font-family: var(--font-heading);
          font-weight: 700;
          margin-bottom: 0.2rem;
          font-size: clamp(0.84rem, 1.4vw, 0.95rem);
        }
        .upl-dropzone.is-dragging .upl-dropzone-headline { color: var(--c-primary); }
        .upl-dropzone-help {
          color: var(--c-text-tertiary);
          font-size: clamp(0.72rem, 1.2vw, 0.78rem); line-height: 1.5;
        }
        .upl-dropzone-link {
          color: var(--c-primary); font-weight: 700;
          text-decoration: underline; text-underline-offset: 2px;
        }

        /* ── File row (when file is set) ── */
        .upl-file-row { display: flex; align-items: center; gap: 0.85rem; width: 100%; }
        .upl-file-icon {
          background: var(--c-bg-elevated);
          border-radius: var(--r-md);
          padding: 0.55rem;
          border: 1px solid var(--c-logo-green);
          flex-shrink: 0;
          display: flex; align-items: center; justify-content: center;
        }
        .upl-file-meta { flex: 1; min-width: 0; }
        .upl-file-name {
          color: var(--c-text-primary);
          font-family: var(--font-heading);
          font-weight: 700;
          font-size: clamp(0.84rem, 1.3vw, 0.92rem);
          margin-bottom: 0.18rem;
          overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }
        .upl-file-size { color: var(--c-text-secondary); font-size: 0.72rem; }
        .upl-file-change {
          color: var(--c-text-secondary);
          background: var(--c-bg-elevated);
          border: 1px solid var(--c-border-strong);
          font-size: 0.7rem; font-weight: 600;
          padding: 0.45rem 0.7rem;
          border-radius: var(--r-md);
          display: flex; align-items: center; gap: 0.3rem;
          flex-shrink: 0;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-file-change:hover {
          background: var(--c-primary);
          color: white;
          border-color: var(--c-primary);
        }

        /* ── Discrete organizer ── */
        .upl-organizer-toggle {
          display: inline-flex; align-items: center; gap: 0.4rem;
          padding: 0.45rem 0.65rem;
          background: transparent; border: none;
          color: var(--c-text-tertiary);
          font-size: 0.74rem; font-weight: 600;
          cursor: pointer; align-self: flex-start;
          transition: color var(--t-base) var(--ease-out);
        }
        .upl-organizer-toggle:hover { color: var(--c-primary); }

        .upl-organizer-box {
          padding: 0.95rem 1.05rem;
          background: var(--c-bg-tinted);
          border: 1px solid var(--c-border);
          border-radius: var(--r-md);
          display: flex; flex-direction: column; gap: 0.65rem;
        }
        .upl-organizer-help { color: var(--c-text-secondary); font-size: 0.76rem; line-height: 1.5; }
        .upl-organizer-modes {
          display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.45rem;
        }
        @media (min-width: 520px) {
          .upl-organizer-modes { grid-template-columns: repeat(4, 1fr); }
        }
        .upl-org-mode {
          padding: 0.55rem 0.7rem;
          background: var(--c-bg-elevated);
          border: 1px solid var(--c-border);
          border-radius: var(--r-md);
          color: var(--c-text-secondary);
          display: flex; flex-direction: column; gap: 2px;
          text-align: left; cursor: pointer;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-org-mode:hover {
          background: var(--c-accent-green);
          color: var(--c-primary);
          border-color: var(--c-primary);
        }
        .upl-org-mode.is-active {
          background: var(--c-primary);
          border-color: var(--c-primary);
          color: white;
        }
        .upl-org-mode-label { font-size: 0.78rem; font-weight: 700; font-family: var(--font-heading); }
        .upl-org-mode-desc { font-size: 0.66rem; opacity: 0.85; }
        .upl-organizer-foot { color: var(--c-text-muted); font-size: 0.7rem; font-style: italic; }

        .upl-organize-cta {
          display: inline-flex; align-items: center; gap: 0.4rem;
          align-self: flex-start;
          padding: 0.5rem 0.85rem;
          background: var(--c-bg-elevated);
          border: 1.5px solid var(--c-logo-green);
          border-radius: var(--r-md);
          color: var(--c-logo-green);
          font-family: var(--font-heading);
          font-size: 0.7rem; font-weight: 700;
          letter-spacing: 0.04em; text-transform: uppercase;
          cursor: pointer;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-organize-cta:hover {
          background: var(--c-logo-green);
          color: white;
        }

        /* ── Banners ── */
        .upl-banner { padding: 0.95rem 1.1rem; border-radius: var(--r-md); }
        .upl-banner-error {
          background: #FEF2F2;
          border: 1px solid rgba(212, 56, 56, 0.30);
          display: flex; gap: 0.65rem; align-items: flex-start;
        }
        .upl-banner-error-icon { flex-shrink: 0; margin-top: 2px; }
        .upl-banner-error-title {
          color: var(--c-error-300);
          font-family: var(--font-heading);
          font-size: 0.84rem; font-weight: 700; margin-bottom: 0.2rem;
        }
        .upl-banner-error-msg { color: var(--c-text-secondary); font-size: 0.76rem; line-height: 1.5; }
        .upl-banner-retry {
          background: #FEF2F2;
          border: 1px solid rgba(212, 56, 56, 0.40);
          color: var(--c-error-300);
        }
        .upl-banner-retry-title {
          font-family: var(--font-heading);
          font-weight: 700; margin-bottom: 0.3rem; font-size: 0.86rem;
        }
        .upl-banner-retry-action { font-size: 0.78rem; color: var(--c-text-secondary); margin-bottom: 0.65rem; }
        .upl-banner-retry-actions { display: flex; gap: 0.55rem; }
        .upl-btn-retry {
          padding: 0.5rem 0.9rem;
          background: var(--c-primary);
          border: 1px solid var(--c-primary);
          border-radius: var(--r-md);
          color: white;
          font-family: var(--font-heading);
          font-size: 0.74rem; font-weight: 700;
          letter-spacing: 0.04em; text-transform: uppercase;
          transition: background var(--t-base) var(--ease-out);
        }
        .upl-btn-retry:hover { background: var(--c-primary-dark); }
        .upl-btn-cancel {
          padding: 0.5rem 0.9rem;
          background: transparent;
          border: 1px solid var(--c-border-strong);
          border-radius: var(--r-md);
          color: var(--c-text-secondary);
          font-family: var(--font-heading);
          font-size: 0.74rem; font-weight: 700;
          letter-spacing: 0.04em; text-transform: uppercase;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-btn-cancel:hover {
          background: var(--c-bg-tinted);
          color: var(--c-text-primary);
          border-color: var(--c-primary);
        }

        .upl-banner-success {
          background: var(--c-accent-green);
          border: 1px solid rgba(105, 190, 40, 0.40);
        }
        .upl-banner-success-row { display: flex; gap: 0.65rem; align-items: center; margin-bottom: 0.4rem; }
        .upl-banner-success-title {
          color: var(--c-primary-dark);
          font-family: var(--font-heading);
          font-size: 0.88rem; font-weight: 800;
        }
        .upl-banner-success-sub { color: var(--c-text-secondary); font-size: 0.75rem; margin-left: 1.5rem; line-height: 1.5; }
        .upl-banner-success-big {
          background: var(--c-accent-green);
          border: 1px solid rgba(105, 190, 40, 0.40);
          padding: 1rem 1.1rem;
        }
        .upl-banner-success-iconwrap {
          background: var(--c-logo-green);
          border-radius: var(--r-md);
          padding: 0.4rem;
          color: white;
          display: flex; align-items: center; justify-content: center;
        }
        .upl-banner-success-iconwrap svg { color: white !important; stroke: white; }
        .upl-banner-success-stats {
          display: flex; gap: 1rem; flex-wrap: wrap;
          margin-top: 0.6rem; padding-top: 0.6rem;
          border-top: 1px solid rgba(105, 190, 40, 0.30);
        }
        .upl-stat { display: flex; align-items: center; gap: 0.4rem; }
        .upl-stat-label { color: var(--c-text-tertiary); font-size: 0.7rem; }
        .upl-stat-value { color: var(--c-primary); font-size: 0.74rem; font-weight: 700; font-family: var(--font-heading); }
        .upl-banner-success-actions { display: flex; align-items: center; gap: 0.65rem; margin-top: 0.8rem; flex-wrap: wrap; }
        .upl-btn-secondary {
          padding: 0.5rem 0.85rem;
          background: var(--c-primary);
          border: 1px solid var(--c-primary);
          border-radius: var(--r-md);
          color: white;
          font-family: var(--font-heading);
          font-size: 0.72rem; font-weight: 700;
          letter-spacing: 0.04em; text-transform: uppercase;
          display: flex; align-items: center; gap: 0.35rem;
          transition: background var(--t-base) var(--ease-out);
        }
        .upl-btn-secondary:hover { background: var(--c-primary-dark); }
        .upl-btn-ghost {
          background: transparent;
          border: 1.5px solid var(--c-primary);
          border-radius: var(--r-md);
          padding: 0.5rem 0.85rem;
          color: var(--c-primary);
          font-family: var(--font-heading);
          font-size: 0.72rem; font-weight: 700;
          letter-spacing: 0.04em; text-transform: uppercase;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-btn-ghost:hover { background: var(--c-accent-green); color: var(--c-primary-dark); }

        /* ── System link row ── */
        .upl-syslink-row { display: flex; justify-content: flex-end; }
        .upl-syslink {
          background: none; border: none; padding: 0;
          color: var(--c-text-muted);
          font-size: 0.68rem; cursor: pointer;
          transition: color var(--t-base) var(--ease-out);
        }
        .upl-syslink:hover { color: var(--c-primary); text-decoration: underline; }

        .upl-health-row {
          display: flex; align-items: center; justify-content: space-between;
          gap: 0.75rem; flex-wrap: wrap;
          padding: 0.75rem 0.9rem;
          background: var(--c-bg-tinted);
          border: 1px solid var(--c-border);
          border-radius: var(--r-md);
        }
        .upl-health-label { color: var(--c-text-primary); font-size: 0.74rem; font-weight: 700; margin-bottom: 0.15rem; }
        .upl-health-status { color: var(--c-text-tertiary); font-size: 0.72rem; }
        .upl-health-status.is-ok { color: var(--c-logo-green); font-weight: 700; }
        .upl-health-status.is-bad { color: var(--c-error-300); font-weight: 700; }
        .upl-health-btn {
          padding: 0.5rem 0.85rem;
          background: var(--c-primary);
          border: 1px solid var(--c-primary);
          border-radius: var(--r-md);
          color: white;
          font-family: var(--font-heading);
          font-size: 0.7rem; font-weight: 700;
          letter-spacing: 0.04em; text-transform: uppercase;
          display: flex; align-items: center; gap: 0.4rem;
          transition: all var(--t-base) var(--ease-out);
        }
        .upl-health-btn:hover:not(:disabled) { background: var(--c-primary-dark); }
        .upl-health-btn:disabled { opacity: 0.5; cursor: not-allowed; }
      `}</style>
    </div>
  );
}
