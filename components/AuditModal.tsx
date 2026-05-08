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
  return (
    <div
      role="dialog"
      aria-modal="true"
      style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100,
      }}
      onClick={onClose}
    >
      <div onClick={e => e.stopPropagation()} style={{
        background: '#0F172A', borderRadius: '14px', padding: '1.5rem',
        maxWidth: '560px', width: '92%', color: 'white',
        maxHeight: '80vh', overflowY: 'auto',
      }}>
        <div style={{
          display: 'flex', justifyContent: 'space-between',
          alignItems: 'center', marginBottom: '1rem',
        }}>
          <h2 style={{ fontSize: '1.1rem', margin: 0 }}>Detalles de la generación</h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="Cerrar"
            style={{
              background: 'none', border: 'none', color: 'white', cursor: 'pointer',
              padding: 0, display: 'flex', alignItems: 'center',
            }}
          >
            <X size={18} />
          </button>
        </div>

        <p style={{
          fontSize: '0.85rem', margin: '0 0 1rem',
          color: status.tone === 'cache' ? '#86EFAC' : status.tone === 'warn' ? '#FCD34D' : '#A78BFA',
        }}>
          {status.label}
        </p>

        <div style={{
          fontSize: '0.8rem', display: 'flex', gap: '1rem',
          marginBottom: '1rem', flexWrap: 'wrap',
        }}>
          <span><strong>{audit.slides_validated}</strong> slides en el PPT</span>
          <span><strong>{audit.slides_dropped.length}</strong> descartados</span>
          <span><strong>{audit.bullets_dropped}</strong> bullets descartados</span>
        </div>

        {audit.slides_dropped.length > 0 && (
          <>
            <h3 style={{ fontSize: '0.85rem', margin: '0 0 0.5rem' }}>
              Slides que omitimos
            </h3>
            <ul style={{
              fontSize: '0.78rem', color: 'rgba(255,255,255,0.7)',
              paddingLeft: '1.2rem', margin: 0,
            }}>
              {audit.slides_dropped.map((d, i) => (
                <li key={i} style={{ marginBottom: '0.3rem' }}>
                  Slide tipo <strong>{d.type}</strong>: {REASON_LABELS[d.reason] || d.reason}
                  {d.block_ref ? ` (bloque ${d.block_ref})` : ''}
                </li>
              ))}
            </ul>
          </>
        )}
      </div>
    </div>
  );
}
