import Link from 'next/link';
import { ArrowLeft, RefreshCw } from 'lucide-react';
import { getRuntimeDependencyStatus } from '@/utils/server-runtime';

export const dynamic = 'force-dynamic';
export const runtime = 'nodejs';

export const metadata = {
  title: 'Estado del sistema · Socya PPTX',
  description: 'Diagnóstico en vivo de las dependencias del generador.',
};

interface CheckRow {
  label: string;
  ok: boolean;
  detail: string;
}

async function aiKeyConfigured(): Promise<boolean> {
  // Treat as configured if env var has a non-empty trimmed value.
  const key = (process.env.OPENROUTER_API_KEY || '').trim();
  return key.length > 16;
}

export default async function StatusPage() {
  const status = await getRuntimeDependencyStatus(true);
  const aiKey = await aiKeyConfigured();

  const checks: CheckRow[] = [
    {
      label: 'Python runtime',
      ok: status.python.ok,
      detail: status.python.ok
        ? `Detectado: ${status.python.version ?? 'versión desconocida'}`
        : `No detectado: ${status.python.error ?? 'sin información'}`,
    },
    {
      label: 'Pipeline Socya (`socya_pipeline`)',
      ok: status.scripts.pipeline,
      detail: status.scripts.pipeline
        ? 'Módulo importable y dependencias presentes'
        : 'No se pudo importar — dependencias faltantes (pandas, openpyxl, python-pptx, matplotlib, requests)',
    },
    {
      label: 'Plantilla institucional',
      ok: status.scripts.template,
      detail: status.scripts.template
        ? 'Plantilla_Presentacion_Socya disponible'
        : 'Plantilla no encontrada en la raíz del proyecto',
    },
    {
      label: 'Asistente IA (OpenRouter)',
      ok: aiKey,
      detail: aiKey
        ? 'API key configurada en el entorno'
        : 'Falta OPENROUTER_API_KEY — la planificación con IA no funcionará',
    },
  ];

  const allGreen = checks.every((c) => c.ok);
  const overallTone: 'ok' | 'degraded' | 'down' = allGreen
    ? 'ok'
    : checks[0].ok && checks[1].ok
      ? 'degraded'
      : 'down';
  const overallLabel = overallTone === 'ok'
    ? 'Operativo'
    : overallTone === 'degraded'
      ? 'Degradado'
      : 'Caído';

  const checkedAt = new Date(status.checkedAt).toLocaleString('es-CO', {
    dateStyle: 'medium', timeStyle: 'medium',
  });

  return (
    <main className="status-shell">
      <div className="status-card">
        <header className="status-head">
          <div>
            <Link href="/" className="status-back">
              <ArrowLeft size={14} />
              Volver al generador
            </Link>
            <h1 className="status-title">Estado del sistema</h1>
            <p className="status-sub">
              Diagnóstico en vivo de las dependencias requeridas para generar PowerPoints.
            </p>
          </div>
          <span className={`status-pill is-${overallTone}`}>
            {overallLabel}
          </span>
        </header>

        <section className="status-list">
          {checks.map((c) => (
            <div key={c.label} className={`status-row is-${c.ok ? 'ok' : 'bad'}`}>
              <div className="status-row-dot" aria-hidden />
              <div className="status-row-text">
                <p className="status-row-label">{c.label}</p>
                <p className="status-row-detail">{c.detail}</p>
              </div>
              <span className="status-row-badge">{c.ok ? 'OK' : 'FALLA'}</span>
            </div>
          ))}
        </section>

        <footer className="status-foot">
          <span>Última verificación: <strong>{checkedAt}</strong></span>
          <form action="/status" method="get">
            <button type="submit" className="status-refresh">
              <RefreshCw size={13} />
              Volver a comprobar
            </button>
          </form>
        </footer>
      </div>

      <style>{`
        .status-shell {
          min-height: 100dvh;
          background: var(--c-bg-neutral);
          color: var(--c-text-primary);
          padding: clamp(1rem, 4vw, 3rem) 1rem;
          display: flex;
          flex-direction: column;
          align-items: center;
          font-family: var(--font-sans);
        }
        .status-card {
          width: 100%;
          max-width: 640px;
          background: var(--c-bg-elevated);
          border: 1px solid var(--c-border);
          border-radius: var(--r-lg);
          box-shadow: var(--shadow-md);
          overflow: hidden;
        }
        .status-head {
          padding: clamp(1rem, 2.4vw, 1.5rem);
          border-bottom: 1px solid var(--c-divider);
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 1rem;
          flex-wrap: wrap;
        }
        .status-back {
          display: inline-flex;
          align-items: center;
          gap: 0.4rem;
          color: var(--c-text-tertiary);
          font-size: 0.74rem;
          text-decoration: none;
          margin-bottom: 0.6rem;
          transition: color 0.18s;
        }
        .status-back:hover { color: var(--c-primary); }
        .status-title {
          font-family: var(--font-heading);
          font-size: 1.35rem;
          font-weight: 800;
          letter-spacing: -0.01em;
          margin: 0;
        }
        .status-sub {
          color: var(--c-text-secondary);
          font-size: 0.82rem;
          margin-top: 0.3rem;
          line-height: 1.5;
          max-width: 36rem;
        }
        .status-pill {
          display: inline-flex;
          align-items: center;
          padding: 0.45rem 0.95rem;
          border-radius: var(--r-pill);
          font-family: var(--font-heading);
          font-size: 0.74rem;
          font-weight: 800;
          letter-spacing: 0.06em;
          text-transform: uppercase;
          flex-shrink: 0;
        }
        .status-pill.is-ok {
          background: var(--c-accent-green);
          color: var(--c-primary-dark);
          border: 1px solid rgba(105, 190, 40, 0.40);
        }
        .status-pill.is-degraded {
          background: #FEF7E0;
          color: #8A6800;
          border: 1px solid rgba(243, 196, 0, 0.45);
        }
        .status-pill.is-down {
          background: #FEF2F2;
          color: var(--c-error-300);
          border: 1px solid rgba(212, 56, 56, 0.45);
        }

        .status-list {
          display: flex;
          flex-direction: column;
        }
        .status-row {
          display: flex;
          align-items: center;
          gap: 0.85rem;
          padding: 0.95rem 1.2rem;
          border-bottom: 1px solid var(--c-divider);
        }
        .status-row:last-child { border-bottom: none; }
        .status-row-dot {
          width: 10px;
          height: 10px;
          border-radius: 50%;
          flex-shrink: 0;
        }
        .status-row.is-ok .status-row-dot {
          background: var(--c-logo-green);
          box-shadow: 0 0 0 4px rgba(105, 190, 40, 0.18);
        }
        .status-row.is-bad .status-row-dot {
          background: var(--c-error-300);
          box-shadow: 0 0 0 4px rgba(212, 56, 56, 0.18);
        }
        .status-row-text { flex: 1; min-width: 0; }
        .status-row-label {
          font-family: var(--font-heading);
          font-weight: 700;
          font-size: 0.86rem;
          color: var(--c-text-primary);
        }
        .status-row-detail {
          color: var(--c-text-secondary);
          font-size: 0.74rem;
          line-height: 1.45;
          margin-top: 2px;
        }
        .status-row-badge {
          font-family: var(--font-heading);
          font-size: 0.66rem;
          font-weight: 800;
          letter-spacing: 0.05em;
          padding: 0.3rem 0.55rem;
          border-radius: var(--r-pill);
          flex-shrink: 0;
        }
        .status-row.is-ok .status-row-badge {
          background: var(--c-accent-green);
          color: var(--c-primary-dark);
        }
        .status-row.is-bad .status-row-badge {
          background: #FEF2F2;
          color: var(--c-error-300);
        }

        .status-foot {
          padding: 0.95rem 1.2rem;
          border-top: 1px solid var(--c-divider);
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 0.7rem;
          flex-wrap: wrap;
          color: var(--c-text-tertiary);
          font-size: 0.72rem;
          background: var(--c-bg-tinted);
        }
        .status-foot strong {
          color: var(--c-text-primary);
          font-weight: 700;
        }
        .status-refresh {
          display: inline-flex;
          align-items: center;
          gap: 0.4rem;
          padding: 0.5rem 0.85rem;
          background: var(--c-bg-elevated);
          border: 1px solid var(--c-border-strong);
          border-radius: var(--r-md);
          color: var(--c-text-primary);
          font-family: var(--font-heading);
          font-size: 0.72rem;
          font-weight: 700;
          letter-spacing: 0.04em;
          text-transform: uppercase;
          cursor: pointer;
          transition: all 0.18s;
        }
        .status-refresh:hover {
          background: var(--c-primary);
          color: white;
          border-color: var(--c-primary);
        }
      `}</style>
    </main>
  );
}
