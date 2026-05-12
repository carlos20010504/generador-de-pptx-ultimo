"use client";
import React from 'react';
import { X } from 'lucide-react';
import { aiStatusBadge, AIStatus } from '@/utils/ai-status';

const REASON_LABELS: Record<string, string> = {
  block_ref_not_found: 'la IA referenció un bloque que no existe',
  single_dominant_category: 'una categoría dominaba >85% (gráfica sin sentido)',
  all_zero_values: 'todos los valores eran cero',
  too_few_categories: 'no había suficientes categorías para una gráfica',
  too_few_rows: 'la tabla tenía muy pocas filas',
  low_fill_ratio: 'demasiadas celdas vacías',
  all_bullets_failed_provenance: 'ningún bullet citaba datos del Excel',
  missing_required_field: 'faltó un campo obligatorio',
  no_columns_after_filter: 'no quedaron columnas tras el filtrado',
  no_rows_after_filter: 'no quedaron filas tras el filtrado',
  all_zero: 'todos los valores eran cero',
};

interface Audit {
  model_used?: string;
  cache_hit?: boolean;
  fallback_chain_steps?: { from: string; reason: string }[];
  slides_planned: number;
  slides_validated: number;
  slides_dropped: { type: string; reason: string; block_ref?: string }[];
  bullets_dropped: number;
}

export default function AuditModal({ audit, onClose }: {
  audit: Audit; onClose: () => void
}) {
  const status = aiStatusBadge({
    model: audit.model_used,
    cache_hit: audit.cache_hit,
    fallback_steps: audit.fallback_chain_steps as AIStatus['fallback_steps'],
  });

  // Tone color en clave Socya — todos sobre fondo claro
  const toneColor =
    status.tone === 'cache' ? 'var(--c-logo-green)'
    : status.tone === 'warn' ? '#8A6800'
    : 'var(--c-primary)';

  return (
    <div
      role="dialog"
      aria-modal="true"
      style={{
        position: 'fixed', inset: 0,
        background: 'rgba(18, 60, 73, 0.42)',
        backdropFilter: 'blur(4px)',
        display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100,
      }}
      onClick={onClose}
    >
      <div onClick={e => e.stopPropagation()} style={{
        background: 'var(--c-bg-elevated)',
        border: '1px solid var(--c-border)',
        borderRadius: 'var(--r-lg)',
        padding: '1.5rem',
        maxWidth: '560px', width: '92%',
        color: 'var(--c-text-primary)',
        maxHeight: '80vh', overflowY: 'auto',
        boxShadow: 'var(--shadow-lg)',
      }}>
        <div style={{
          display: 'flex', justifyContent: 'space-between',
          alignItems: 'center', marginBottom: '1rem',
        }}>
          <h2 style={{
            fontSize: '1.05rem', margin: 0,
            fontFamily: 'var(--font-heading)', fontWeight: 800,
            color: 'var(--c-text-primary)',
          }}>Detalles de la generación</h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="Cerrar"
            style={{
              background: 'var(--c-bg-elevated)',
              border: '1px solid var(--c-border-strong)',
              borderRadius: 'var(--r-md)',
              width: '1.95rem', height: '1.95rem',
              color: 'var(--c-text-secondary)',
              cursor: 'pointer',
              padding: 0,
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              transition: 'all 0.18s',
            }}
            onMouseEnter={e => {
              e.currentTarget.style.background = 'var(--c-primary)';
              e.currentTarget.style.color = 'white';
              e.currentTarget.style.borderColor = 'var(--c-primary)';
            }}
            onMouseLeave={e => {
              e.currentTarget.style.background = 'var(--c-bg-elevated)';
              e.currentTarget.style.color = 'var(--c-text-secondary)';
              e.currentTarget.style.borderColor = 'var(--c-border-strong)';
            }}
          >
            <X size={16} />
          </button>
        </div>

        <p style={{
          fontSize: '0.78rem', margin: '0 0 1rem',
          color: toneColor,
          fontFamily: 'var(--font-heading)',
          fontWeight: 700, letterSpacing: '0.04em',
        }}>
          {status.label}
        </p>

        <div style={{
          fontSize: '0.82rem',
          display: 'flex', gap: '1.2rem',
          marginBottom: '1rem', flexWrap: 'wrap',
          color: 'var(--c-text-secondary)',
          paddingBottom: '1rem',
          borderBottom: '1px solid var(--c-divider)',
        }}>
          <span>
            <strong style={{ color: 'var(--c-primary)', fontFamily: 'var(--font-heading)' }}>
              {audit.slides_validated}
            </strong> slides en el PPT
          </span>
          <span>
            <strong style={{ color: 'var(--c-text-primary)', fontFamily: 'var(--font-heading)' }}>
              {audit.slides_dropped.length}
            </strong> descartados
          </span>
          <span>
            <strong style={{ color: 'var(--c-text-primary)', fontFamily: 'var(--font-heading)' }}>
              {audit.bullets_dropped}
            </strong> bullets descartados
          </span>
        </div>

        {audit.slides_dropped.length > 0 && (
          <>
            <h3 style={{
              fontSize: '0.78rem', margin: '0 0 0.55rem',
              fontFamily: 'var(--font-heading)',
              fontWeight: 700, textTransform: 'uppercase',
              letterSpacing: '0.06em',
              color: 'var(--c-text-secondary)',
            }}>
              Slides que omitimos
            </h3>
            <ul style={{
              fontSize: '0.78rem', color: 'var(--c-text-secondary)',
              paddingLeft: '1.2rem', margin: 0,
              lineHeight: 1.6,
            }}>
              {audit.slides_dropped.map((d, i) => (
                <li key={i} style={{ marginBottom: '0.35rem' }}>
                  Slide tipo{' '}
                  <strong style={{ color: 'var(--c-text-primary)' }}>{d.type}</strong>:{' '}
                  {REASON_LABELS[d.reason] || d.reason}
                  {d.block_ref ? (
                    <span style={{ color: 'var(--c-text-tertiary)' }}> (bloque {d.block_ref})</span>
                  ) : ''}
                </li>
              ))}
            </ul>
          </>
        )}
      </div>
    </div>
  );
}
