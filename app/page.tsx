import ExcelUploader from '@/components/ExcelUploader';
import { BarChart3, Sparkles } from 'lucide-react';

export const metadata = {
  title: 'Socya PPTX Generator – Excel a PowerPoint Inteligente',
  description: 'Convierte cualquier archivo Excel en una presentación profesional con gráficos, tablas, KPIs y badges — todo con detección automática de datos.',
};

export default function Home() {
  return (
    <main className="page-shell">
      {/* Ambient background (decorative, non-interactive) */}
      <div className="bg-grid" aria-hidden />
      <div className="bg-orb orb-blue" aria-hidden />
      <div className="bg-orb orb-green" aria-hidden />
      <div className="bg-topline" aria-hidden />

      {/* Compact brand strip — no hero, no duplicated subtitle */}
      <header className="page-header animate-fade-in-up">
        <div className="brand">
          <div className="brand-mark" aria-hidden>
            <BarChart3 size={20} color="white" />
          </div>
          <div className="brand-text">
            <p className="brand-title">Socya PPTX Generator</p>
            <p className="brand-sub">v4.0 · Hermes-driven</p>
          </div>
        </div>

        <div className="status-pill" role="status">
          <Sparkles size={11} />
          <span>Detección inteligente activa</span>
        </div>
      </header>

      {/* Main interactive surface */}
      <section className="uploader-shell animate-fade-in-up stagger-1">
        <ExcelUploader />
      </section>

      <style>{`
        .page-shell {
          min-height: 100dvh;
          width: 100%;
          max-width: 100vw;
          background: var(--c-bg-deep);
          display: flex;
          flex-direction: column;
          align-items: center;
          padding: clamp(0.75rem, 2vw, 1.25rem);
          font-family: var(--font-sans);
          position: relative;
          overflow-x: clip;
        }

        .bg-grid {
          position: absolute; inset: 0;
          background-image:
            linear-gradient(rgba(255, 255, 255, 0.012) 1px, transparent 1px),
            linear-gradient(90deg, rgba(255, 255, 255, 0.012) 1px, transparent 1px);
          background-size: 60px 60px;
          pointer-events: none;
        }
        .bg-orb {
          position: absolute;
          border-radius: 50%;
          pointer-events: none;
          will-change: transform;
        }
        .orb-blue {
          top: -16rem; right: -8rem;
          width: 42rem; height: 42rem;
          background: radial-gradient(circle, rgba(59, 130, 246, 0.10) 0%, transparent 55%);
          animation: float-orb 20s ease-in-out infinite;
        }
        .orb-green {
          bottom: -12rem; left: -10rem;
          width: 36rem; height: 36rem;
          background: radial-gradient(circle, rgba(74, 222, 128, 0.07) 0%, transparent 55%);
          animation: float-orb 25s ease-in-out infinite reverse;
        }
        .bg-topline {
          position: absolute; top: 0; left: 0; right: 0; height: 1px;
          background: linear-gradient(90deg,
            transparent,
            rgba(59, 130, 246, 0.30),
            rgba(99, 102, 241, 0.30),
            transparent);
          pointer-events: none;
        }

        .page-header {
          width: 100%;
          max-width: 960px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: var(--space-3);
          margin-bottom: clamp(0.85rem, 1.6vw, 1.15rem);
          z-index: 10;
        }
        .brand { display: flex; align-items: center; gap: 0.65rem; }
        .brand-mark {
          background: linear-gradient(135deg, var(--c-brand-blue-700), var(--c-brand-blue-500));
          border-radius: var(--r-md);
          padding: 0.45rem 0.5rem;
          display: flex; align-items: center; justify-content: center;
          box-shadow: 0 4px 18px rgba(59, 130, 246, 0.28),
                      inset 0 1px 0 rgba(255, 255, 255, 0.15);
          border: 1px solid rgba(59, 130, 246, 0.30);
        }
        .brand-title {
          color: var(--c-text-primary);
          font-weight: 800; font-size: 0.92rem;
          letter-spacing: -0.02em; line-height: 1.1;
        }
        .brand-sub {
          color: var(--c-text-muted);
          font-size: 0.66rem;
          letter-spacing: 0.02em;
          line-height: 1.2;
          margin-top: 1px;
        }
        .status-pill {
          display: inline-flex; align-items: center;
          gap: var(--space-2);
          padding: 0.32rem 0.6rem;
          border-radius: var(--r-pill);
          background: rgba(74, 222, 128, 0.08);
          border: 1px solid rgba(74, 222, 128, 0.18);
          color: var(--c-success-400);
          font-size: 0.66rem; font-weight: 700;
          letter-spacing: 0.02em;
        }
        @media (max-width: 480px) {
          .brand-sub, .status-pill { display: none; }
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
