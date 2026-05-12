import ExcelUploader from '@/components/ExcelUploader';
import { Sparkles } from 'lucide-react';

export const metadata = {
  title: 'Socya PPTX Generator – Excel a PowerPoint Inteligente',
  description: 'Convierte cualquier archivo Excel en una presentación profesional con gráficos, tablas, KPIs y badges — todo con detección automática de datos.',
};

export default function Home() {
  return (
    <main className="page-shell">
      {/* Subtle ambient teal wash — anchored to brand, no dark orbs */}
      <div className="bg-wash bg-wash-top" aria-hidden />
      <div className="bg-wash bg-wash-bottom" aria-hidden />

      {/* Brand strip · 3px logo-green divider · Socya pattern (Manual DAPC08) */}
      <header className="page-header animate-fade-in-up">
        <div className="brand">
          <div className="brand-mark" aria-hidden>
            {/* Stylized "S" leaf — minimal Socya glyph */}
            <svg viewBox="0 0 32 32" width="22" height="22" fill="none">
              <path
                d="M22.5 9.4c-1.6-1.7-4-2.6-6.6-2.6-3.5 0-6.4 2.1-6.4 4.7 0 2.4 2 3.7 5.7 4.5l1.6.4c4.5 1 7.1 2.7 7.1 6.1 0 3.6-3.6 6.2-8.4 6.2-3.4 0-6.3-1.4-7.9-3.5"
                stroke="#69BE28" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round"
              />
              <circle cx="24" cy="6.5" r="1.8" fill="#69BE28" />
            </svg>
          </div>
          <div className="brand-text">
            <p className="brand-title">Socya <span className="brand-title-accent">PPTX</span></p>
            <p className="brand-sub">Generador inteligente · v4.0</p>
          </div>
        </div>

        <div className="status-pill" role="status">
          <Sparkles size={11} />
          <span>Detección inteligente activa</span>
        </div>
      </header>

      <div className="socya-divider" aria-hidden />

      {/* Main interactive surface */}
      <section className="uploader-shell animate-fade-in-up stagger-1">
        <ExcelUploader />
      </section>

      <style>{`
        .page-shell {
          min-height: 100dvh;
          width: 100%;
          max-width: 100vw;
          background: var(--c-bg-neutral);
          display: flex;
          flex-direction: column;
          align-items: center;
          padding: clamp(0.85rem, 2vw, 1.4rem);
          font-family: var(--font-sans);
          color: var(--c-text-primary);
          position: relative;
          overflow-x: clip;
        }

        /* Subtle teal wash — replaces dark orbs. Light, on-brand, never noisy. */
        .bg-wash {
          position: absolute;
          inset: 0;
          pointer-events: none;
          opacity: 0.55;
        }
        .bg-wash-top {
          background:
            radial-gradient(60% 35% at 18% 0%, rgba(8, 112, 98, 0.10) 0%, transparent 70%),
            radial-gradient(45% 30% at 92% 4%, rgba(105, 190, 40, 0.10) 0%, transparent 70%);
        }
        .bg-wash-bottom {
          background:
            radial-gradient(50% 35% at 80% 100%, rgba(8, 112, 98, 0.08) 0%, transparent 70%);
        }

        .page-header {
          width: 100%;
          max-width: 960px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: var(--space-3);
          margin-bottom: 0;
          padding: 0.85rem 0.25rem;
          z-index: 10;
        }
        .brand { display: flex; align-items: center; gap: 0.7rem; }
        .brand-mark {
          background: white;
          border-radius: var(--r-md);
          padding: 0.35rem 0.45rem;
          display: flex; align-items: center; justify-content: center;
          border: 1px solid var(--c-border);
          box-shadow: var(--shadow-sm);
        }
        .brand-title {
          color: var(--c-text-primary);
          font-family: var(--font-heading);
          font-weight: 800;
          font-size: 1rem;
          letter-spacing: -0.02em;
          line-height: 1.05;
        }
        .brand-title-accent {
          color: var(--c-primary);
        }
        .brand-sub {
          color: var(--c-text-tertiary);
          font-family: var(--font-heading);
          font-size: 0.66rem;
          letter-spacing: 0.04em;
          line-height: 1.2;
          margin-top: 2px;
          text-transform: uppercase;
        }

        .status-pill {
          display: inline-flex; align-items: center;
          gap: var(--space-2);
          padding: 0.34rem 0.65rem;
          border-radius: var(--r-pill);
          background: var(--c-accent-green);
          border: 1px solid rgba(105, 190, 40, 0.30);
          color: var(--c-primary);
          font-family: var(--font-heading);
          font-size: 0.65rem;
          font-weight: 700;
          letter-spacing: 0.06em;
          text-transform: uppercase;
        }
        @media (max-width: 480px) {
          .brand-sub, .status-pill { display: none; }
        }

        .socya-divider {
          width: 100%;
          max-width: 960px;
          margin-bottom: clamp(1rem, 2vw, 1.4rem);
          z-index: 10;
        }

        .uploader-shell {
          width: 100%;
          max-width: 960px;
          z-index: 10;
        }
      `}</style>
    </main>
  );
}
