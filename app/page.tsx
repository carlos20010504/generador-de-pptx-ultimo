import ExcelUploader from '@/components/ExcelUploader';
import { BarChart3 } from 'lucide-react';

export const metadata = {
  title: 'Socya PPTX Generator – Excel a PowerPoint Inteligente',
  description: 'Convierte cualquier archivo Excel en una presentación profesional con gráficos, tablas, KPIs y badges — todo con detección automática de datos.',
};

const FEATURES = [
  { emoji: '1', label: 'Organiza', desc: 'Prepara el Excel' },
  { emoji: '2', label: 'Genera', desc: 'Crea el PowerPoint' },
  { emoji: '✨', label: 'Flujo guiado', desc: 'Menos fricción' },
];

export default function Home() {
  return (
    <main style={{
      minHeight: '100dvh',
      background: '#060D18',
      display: 'flex', flexDirection: 'column', alignItems: 'center',
      padding: '1rem 1rem 0.8rem',
      fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif",
      position: 'relative', overflow: 'hidden',
    }}>

      {/* ── Background Effects ── */}
      {/* Grid pattern */}
      <div style={{
        position: 'absolute', inset: 0,
        backgroundImage: `
          linear-gradient(rgba(255,255,255,0.015) 1px, transparent 1px),
          linear-gradient(90deg, rgba(255,255,255,0.015) 1px, transparent 1px)
        `,
        backgroundSize: '60px 60px',
        pointerEvents: 'none',
      }} />

      {/* Gradient orbs */}
      <div style={{
        position: 'absolute', top: '-20rem', right: '-10rem',
        width: '55rem', height: '55rem', borderRadius: '50%',
        background: 'radial-gradient(circle, rgba(59,130,246,0.12) 0%, transparent 55%)',
        pointerEvents: 'none', animation: 'float-orb 20s ease-in-out infinite',
      }} />
      <div style={{
        position: 'absolute', bottom: '-15rem', left: '-12rem',
        width: '48rem', height: '48rem', borderRadius: '50%',
        background: 'radial-gradient(circle, rgba(74,222,128,0.08) 0%, transparent 55%)',
        pointerEvents: 'none', animation: 'float-orb 25s ease-in-out infinite reverse',
      }} />
      <div style={{
        position: 'absolute', top: '40%', left: '50%', transform: 'translate(-50%, -50%)',
        width: '40rem', height: '40rem', borderRadius: '50%',
        background: 'radial-gradient(circle, rgba(99,102,241,0.06) 0%, transparent 50%)',
        pointerEvents: 'none',
      }} />

      {/* Top fade */}
      <div style={{
        position: 'absolute', top: 0, left: 0, right: 0, height: '1px',
        background: 'linear-gradient(90deg, transparent, rgba(59,130,246,0.3), rgba(99,102,241,0.3), transparent)',
        pointerEvents: 'none',
      }} />

      {/* ── Logo / Brand ── */}
      <div
        style={{
          display: 'flex', alignItems: 'center', gap: '0.75rem',
          marginBottom: '0.85rem', zIndex: 10,
        }}
        className="animate-fade-in-up"
      >
        <div style={{
          background: 'linear-gradient(135deg, #1E40AF, #3B82F6)',
          borderRadius: '14px', padding: '0.65rem',
          display: 'flex',
          boxShadow: '0 4px 24px rgba(59,130,246,0.35)',
          border: '1px solid rgba(59,130,246,0.3)',
        }}>
          <BarChart3 size={24} color="white" />
        </div>
        <div>
          <p style={{
            color: 'rgba(255,255,255,0.92)', fontWeight: 800,
            fontSize: '1.08rem', margin: 0, letterSpacing: '-0.02em',
          }}>
            Socya PPTX Generator
          </p>
          <p style={{
            color: 'rgba(255,255,255,0.3)', fontSize: '0.72rem',
            margin: 0, letterSpacing: '0.02em',
          }}>
            Motor de Generación Automática v4.0
          </p>
        </div>
      </div>

      {/* ── Headline ── */}
      <header
        style={{
          textAlign: 'center', marginBottom: '0.7rem',
          zIndex: 10, maxWidth: '640px',
        }}
        className="animate-fade-in-up stagger-1"
      >
        <div style={{
          display: 'inline-flex',
          alignItems: 'center',
          gap: '0.45rem',
          padding: '0.45rem 0.8rem',
          borderRadius: '999px',
          background: 'rgba(74,222,128,0.08)',
          border: '1px solid rgba(74,222,128,0.16)',
          color: '#86EFAC',
          fontSize: '0.7rem',
          fontWeight: 800,
          marginBottom: '0.65rem',
        }}>
          Flujo recomendado: organizar Excel y luego generar PPTX
        </div>
        <h1 style={{
          fontSize: 'clamp(1.5rem, 3vw, 2.2rem)', fontWeight: 900,
          color: 'white', margin: '0 0 0.45rem',
          letterSpacing: '-0.035em', lineHeight: 1.12,
        }}>
          De{' '}
          <span style={{ color: '#4ADE80' }}>Excel</span>
          {' '}a{' '}
          <span style={{
            background: 'linear-gradient(90deg, #60A5FA, #818CF8, #A78BFA)',
            WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent',
          }}>
            PowerPoint
          </span>
          {' '}inteligente
        </h1>
        <p style={{
          color: 'rgba(255,255,255,0.40)', fontSize: '0.8rem',
          lineHeight: 1.45, margin: 0,
        }}>
          Un flujo simple para que sea claro qué hacer primero:
          <strong style={{ color: 'rgba(255,255,255,0.68)' }}> organizar el Excel</strong> y luego
          <strong style={{ color: 'rgba(255,255,255,0.68)' }}> generar el PowerPoint</strong>.
        </p>
      </header>

      {/* ── Uploader Component ── */}
      <div
        style={{ width: '100%', maxWidth: '760px', zIndex: 10 }}
        className="animate-fade-in-up stagger-2"
      >
        <ExcelUploader />
      </div>

      {/* ── Feature Grid ── */}
      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          justifyContent: 'center',
          gap: '0.5rem 0.6rem',
          marginTop: '0.65rem', zIndex: 10,
          maxWidth: '520px',
        }}
        className="animate-fade-in-up stagger-3"
      >
        {FEATURES.map((f) => (
          <div
            key={f.label}
            style={{
              display: 'flex', alignItems: 'center', gap: '0.4rem',
              padding: '0.32rem 0.58rem',
              background: 'rgba(255,255,255,0.03)',
              borderRadius: '999px',
              border: '1px solid rgba(255,255,255,0.06)',
              transition: 'all 0.2s',
            }}
          >
            <span style={{ fontSize: '0.85rem' }}>{f.emoji}</span>
            <div>
              <p style={{
                color: 'rgba(255,255,255,0.55)', fontWeight: 700,
                fontSize: '0.67rem', margin: 0, lineHeight: 1.2,
              }}>
                {f.label}
              </p>
            </div>
          </div>
        ))}
      </div>
    </main>
  );
}
