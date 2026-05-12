# UI cleanup — reducir carga visual sin tocar UX

**Fecha:** 2026-05-12
**Decisión:** rediseño moderado (B + selección de A) — flujo lineal único.

## Problema

Después de añadir `OnboardingPanel` y `PlanPreviewPanel`, la página principal acumula:
- 3 fuentes de "sugerencias" (Onboarding.suggestions, AIControlPanel.suggestions, AIControlPanel.promptHints).
- 2 listas de slides previstas (PlanPreviewPanel + AIControlPanel.recommendedSlides).
- 3 selectores de "cómo lo quieres" (step tabs, mode selector, AIControlPanel.focus).
- Header + banner contextual + hero gigante repiten el mismo mensaje.
- Sidebar fija de 360px que el usuario rara vez toca, ocupa columna constante.

## Objetivo

Pantalla limpia con jerarquía clara, UX preservada (mismos endpoints, misma lógica, misma capacidad de configurar), una sola CTA primaria visible.

## Cambios

### Layout

Antes: hero + card con tabs + 2 paneles secuenciales + 3 cards de modo + 2 CTAs + sidebar fija.
Después: header compacto + card único con flujo lineal + drawer de avanzado (oculto por defecto).

### Componentes

- **NUEVO** `components/PreparePanel.tsx` — fusiona `OnboardingPanel` + `PlanPreviewPanel`. Tres secciones internas:
  1. **Detección** (siempre abierta, compacta): 4 chips de stats + top-3 KPIs inline + warnings warn/error solamente (info se agrupa).
  2. **Plan de slides** (siempre abierta): lista de 1 línea por slide, click expande detalles, checkbox toggle, contador X/Y, pills resumen.
  3. **Refinar** (cerrada por defecto, "+ Refinar prompt y contexto"): textarea de prompt + sugerencias inline + audiencia/tema/modo en una fila.

- **MODIFICADO** `components/ExcelUploader.tsx`:
  - Elimina step tabs Organizar/Generar.
  - Elimina banner contextual del step.
  - Elimina mode selector como sección visible (vive dentro de PreparePanel.Refinar).
  - Elimina CTA "Revisar estructura" (PlanPreview ya cumple esa función).
  - Elimina sidebar columna fija → reemplaza por botón flotante "⚙️ Avanzado" que abre drawer.
  - Mantiene: dropzone, error/retry banners, success banner, audit modal, animaciones, todo el flujo de generación SSE.
  - Flujo de "organizar" pasa a link discreto dentro del dropzone ("¿Excel desordenado? Organízalo primero").

- **NUEVO** `components/AdvancedDrawer.tsx` — wrapper que muestra/oculta `AIControlPanel` desde el lado derecho. AIControlPanel sin cambios internos.

- **MODIFICADO** `app/page.tsx`:
  - Elimina `<section className="hero">` (líneas 37-49).
  - Header queda: brand + status pill solamente.
  - Tagline corto pasa al header interno del card.

### Carga de datos

`PreparePanel` dispara `/api/quick-summary` y `/api/preview-plan` en paralelo con `Promise.allSettled` para que un fallo no bloquee al otro. Los `useEffect` con `fileKey` se preservan.

### Lo que NO cambia

- Endpoints (`quick-summary`, `preview-plan`, `generate-pptx`, `excel-intelligence`, `advanced-generate`, `health`).
- Lógica de pipeline Python.
- Estados existentes (`status`, `progressPhase`, `retryError`, `audit`, etc.).
- Componentes `GenerationProgress`, `AuditModal`, `AIControlPanel` (interno).
- Validaciones de archivo, error codes, SSE handling.

## Métricas esperadas

- ExcelUploader: ~1800 → ~1100 líneas.
- Secciones visibles con archivo cargado: ~7 → ~3.
- Componentes en pantalla simultáneos: 4 (uploader+onboarding+preview+sidebar) → 2 (uploader+drawer-toggle).

## Riesgos

- **Drawer en mobile**: debe ser bottom-sheet o full-screen. Verificar con Playwright en 375px.
- **AIControlPanel auto-load**: hoy carga sugerencias al subir archivo. Si el drawer está cerrado, igual debe cargar (cache caliente para cuando el usuario abra).
- **Modo selector dentro de Refinar**: usuario que hoy está acostumbrado a verlo arriba debe poder encontrarlo. Mitigación: mostrar el modo activo como pill en el header del PreparePanel.

## Verificación

- Playwright screenshots: desktop 1280px y mobile 375px del flujo completo (upload → prepare → drawer → generate).
- Smoke test manual: subir Excel, togglear slides, abrir drawer, cambiar tema, generar.
- `tsc --noEmit` y `next build` deben pasar sin errores.
