"use client";
import React from 'react';
import { FileSpreadsheet, Database, Brain, Check, Palette, Loader2 } from 'lucide-react';

const PHASES = [
  { key: 'parsing',    icon: FileSpreadsheet, label: 'Leyendo Excel' },
  { key: 'inventory',  icon: Database,        label: 'Construyendo inventario' },
  { key: 'planning',   icon: Brain,           label: 'Consultando IA' },
  { key: 'validating', icon: Check,           label: 'Validando datos' },
  { key: 'rendering',  icon: Palette,         label: 'Renderizando PPTX' },
] as const;

interface Props {
  currentPhase: typeof PHASES[number]['key'] | 'done' | 'error' | null;
  message?: string;
}

export default function GenerationProgress({ currentPhase, message }: Props) {
  const idx = PHASES.findIndex(p => p.key === currentPhase);
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.45rem' }}>
      {PHASES.map((p, i) => {
        const done = idx > i || currentPhase === 'done';
        const active = idx === i;
        const Icon = p.icon;
        return (
          <div key={p.key} style={{
            display: 'flex', alignItems: 'center', gap: '0.6rem',
            padding: '0.55rem 0.75rem',
            borderRadius: 'var(--r-md)',
            background: active
              ? 'rgba(8, 112, 98, 0.08)'
              : done
                ? 'rgba(105, 190, 40, 0.06)'
                : 'transparent',
            border: active
              ? '1px solid rgba(8, 112, 98, 0.25)'
              : done
                ? '1px solid rgba(105, 190, 40, 0.20)'
                : '1px solid transparent',
            color: done
              ? 'var(--c-logo-green)'
              : active
                ? 'var(--c-primary)'
                : 'var(--c-text-muted)',
            fontSize: '0.78rem',
            fontWeight: active || done ? 700 : 500,
            transition: 'all 0.22s',
          }}>
            {active
              ? <Loader2 size={14} className="spin" style={{ animation: 'spin 1s linear infinite' }} />
              : done
                ? <Check size={14} strokeWidth={3} />
                : <Icon size={14} />}
            <span>{p.label}</span>
            {active && message && (
              <span style={{
                marginLeft: 'auto',
                color: 'var(--c-text-tertiary)',
                fontSize: '0.7rem',
                fontWeight: 500,
              }}>
                {message}
              </span>
            )}
          </div>
        );
      })}
    </div>
  );
}
