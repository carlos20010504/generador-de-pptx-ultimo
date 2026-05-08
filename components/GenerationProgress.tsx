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
    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
      {PHASES.map((p, i) => {
        const done = idx > i || currentPhase === 'done';
        const active = idx === i;
        const Icon = p.icon;
        return (
          <div key={p.key} style={{
            display: 'flex', alignItems: 'center', gap: '0.6rem',
            padding: '0.5rem 0.7rem', borderRadius: '8px',
            background: active ? 'rgba(124,58,237,0.12)' : 'transparent',
            color: done ? '#86EFAC' : active ? '#A78BFA' : 'rgba(255,255,255,0.35)',
            fontSize: '0.78rem',
          }}>
            {active ? <Loader2 size={14} className="spin" /> : <Icon size={14} />}
            <span>{p.label}</span>
            {active && message && (
              <span style={{ marginLeft: 'auto', opacity: 0.7, fontSize: '0.7rem' }}>
                {message}
              </span>
            )}
          </div>
        );
      })}
    </div>
  );
}
