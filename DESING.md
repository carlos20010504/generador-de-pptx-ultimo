---
name: Socya
description: >
  Sistema de identidad visual oficial de Fundacion Socya — consultor y operador
  socioambiental lider en Colombia. Basado en el Manual de Identidad Corporativa
  (Codigo DAPC08, Version 01, Junio 2014) y la paleta digital extendida del sitio
  web socya.org.co.
version: alpha

colors:
  # Colores oficiales del logosimbolo (Manual quierDAPC08, p.7)
  logo-green: "#69BE28"
  logo-gray: "#4D4F53"

  # Paleta digital extendida (sitio web socya.org.co)
  primary: "#087062"
  primary-dark: "#123C49"
  sky: "#00A0DF"
  lime: "#80C612"
  yellow: "#F3C400"
  orange: "#FF8300"

  # Neutros y superficies
  on-primary: "#FFFFFF"
  surface: "#FFFFFF"
  neutral: "#F5F5F5"
  on-surface: "#1A1A1A"
  text-secondary: "#4D4F53"
  border: "#D9D9D9"
  accent-green: "#EEF7E6"

typography:
  # Fuente oficial piezas graficas: Futura (Manual DAPC08 p.14)
  # Alternativa web (Futura no es Google Font): Jost o Nunito Sans
  # Para documentos internos: Calibri (Manual DAPC08 p.14)
  h1:
    fontFamily: "Futura, Jost, sans-serif"
    fontSize: "2.5rem"
    fontWeight: "800"
    lineHeight: "1.15"
    letterSpacing: "-0.01em"
  h2:
    fontFamily: "Futura, Jost, sans-serif"
    fontSize: "2rem"
    fontWeight: "700"
    lineHeight: "1.2"
  h3:
    fontFamily: "Futura, Jost, sans-serif"
    fontSize: "1.5rem"
    fontWeight: "700"
    lineHeight: "1.3"
  h4:
    fontFamily: "Futura, Jost, sans-serif"
    fontSize: "1.125rem"
    fontWeight: "600"
    lineHeight: "1.4"
  body-lg:
    fontFamily: "Calibri, Open Sans, sans-serif"
    fontSize: "1.125rem"
    fontWeight: "400"
    lineHeight: "1.7"
  body-md:
    fontFamily: "Calibri, Open Sans, sans-serif"
    fontSize: "1rem"
    fontWeight: "400"
    lineHeight: "1.6"
  body-sm:
    fontFamily: "Calibri, Open Sans, sans-serif"
    fontSize: "0.875rem"
    fontWeight: "400"
    lineHeight: "1.5"
  label:
    fontFamily: "Futura, Jost, sans-serif"
    fontSize: "0.75rem"
    fontWeight: "700"
    letterSpacing: "0.08em"
  nav:
    fontFamily: "Futura, Jost, sans-serif"
    fontSize: "0.9rem"
    fontWeight: "500"

rounded:
  sm: "4px"
  md: "8px"
  lg: "16px"
  xl: "24px"
  full: "9999px"

spacing:
  xs: "4px"
  sm: "8px"
  md: "16px"
  lg: "24px"
  xl: "40px"
  2xl: "64px"
  3xl: "96px"

components:
  button-primary:
    backgroundColor: "{colors.primary}"
    textColor: "{colors.on-primary}"
    typography: "{typography.label}"
    rounded: "{rounded.md}"
    padding: "12px 28px"
  button-primary-hover:
    backgroundColor: "{colors.primary-dark}"
    textColor: "{colors.on-primary}"
  button-secondary:
    backgroundColor: "{colors.logo-green}"
    textColor: "{colors.on-primary}"
    typography: "{typography.label}"
    rounded: "{rounded.md}"
    padding: "12px 28px"
  button-ghost:
    backgroundColor: "transparent"
    textColor: "{colors.primary}"
    rounded: "{rounded.md}"
    padding: "12px 28px"
  button-ghost-hover:
    backgroundColor: "{colors.accent-green}"
    textColor: "{colors.primary-dark}"
  nav-link:
    textColor: "{colors.on-surface}"
    typography: "{typography.nav}"
  nav-link-active:
    textColor: "{colors.primary}"
  card:
    backgroundColor: "{colors.surface}"
    rounded: "{rounded.lg}"
    padding: "32px"
  card-teal:
    backgroundColor: "{colors.accent-green}"
    rounded: "{rounded.lg}"
    padding: "32px"
  badge:
    backgroundColor: "{colors.accent-green}"
    textColor: "{colors.primary}"
    rounded: "{rounded.full}"
    padding: "4px 12px"
    typography: "{typography.label}"
  badge-sky:
    backgroundColor: "{colors.sky}"
    textColor: "{colors.on-primary}"
    rounded: "{rounded.full}"
    padding: "4px 12px"
    typography: "{typography.label}"
  footer:
    backgroundColor: "{colors.primary}"
    textColor: "{colors.on-primary}"
  footer-link:
    textColor: "{colors.on-primary}"
  divider:
    backgroundColor: "{colors.logo-green}"
    height: "3px"
---

## Overview

**SOCIAL + AMBIENTAL = SOCYA.** El nombre es un acronimo que expresa la razon
de ser de la organizacion: la union de lo social y lo ambiental. Su simbolo es
una figura humana formada por hojas verdes que concentra ambos mundos en un
solo trazo organico.

La identidad visual es moderna, dinamica y llena de vida (cita directa del
manual de marca). No busca la solemnidad institucional pesada; busca movimiento,
cercania y proposito. El verde es el color de la naturaleza: armonia, crecimiento,
exuberancia, fertilidad y frescura.

En medios digitales la paleta se amplia con un espectro cromatico vibrante
(teal oscuro, azul cielo, lima, amarillo, naranja) que representa las distintas
unidades y programas de la organizacion, siempre anclado en los dos colores
normativos del logosimbolo.

## Colors

### Colores del logosimbolo (normativos — Manual DAPC08 p.7)

Estos dos colores son los unicos autorizados para reproducir el logosimbolo
en piezas propias de Socya. Su reproduccion es de obligatorio cumplimiento con
la mayor fidelidad posible, independientemente del soporte de impresion.

- **Logo Green (#69BE28) — Pantone 368 / C63 M0 Y97 K0 / R105 G190 B40:**
  El verde del simbolo de marca (figura humana-hoja y punto). Representa
  naturaleza, crecimiento y vida. Tambien se usa como color de titulares en
  piezas graficas y como linea divisora en el pie de pagina.

- **Logo Gray (#4D4F53) — Pantone Cool Gray 11 / C48 M36 Y24 K66 / R77 G79 B83:**
  El gris del logotipo (letras "Socya"). Representa solidez y profesionalismo.
  Usado en el texto del logotipo y como color de cuerpo en piezas formales.

**Regla critica del manual (pagina 13):** El logosimbolo a color solo puede
usarse sobre fondo blanco puro. Sobre otros colores, unicamente en piezas de
terceras marcas donde Socya no controla el diseno. Esta restriccion nunca debe
violarse en piezas propias.

### Paleta digital extendida (socya.org.co)

- **Primary (#087062):** Verde azulado profundo. Color dominante del sitio web,
  fondos de hero, secciones institucionales y footer. Evoca profundidad y
  confianza.
- **Primary Dark (#123C49):** Azul pizarra oscuro. Para hover states y variantes
  de mayor contraste.
- **Sky (#00A0DF):** Azul cielo vibrante. Accion secundaria, iconos interactivos,
  enlaces informativos, badges agua/aire.
- **Lime (#80C612):** Verde lima (cercano a logo-green, optimizado para pantalla).
  Indicadores de progreso, graficas, estados positivos, economia circular.
- **Yellow (#F3C400):** Amarillo solar. KPIs destacados, alertas informativas,
  secciones de logros y estadisticas.
- **Orange (#FF8300):** Naranja. CTAs de maxima urgencia, Negocios Circulares,
  elementos de alta energia visual.

Los colores extendidos son acentos tematicos, nunca fondos dominantes. El
color `primary` (#087062) domina el UI; `logo-green` y `logo-gray` son
exclusivos del logosimbolo.

## Typography

Socya define una jerarquia tipografica dual segun el contexto (Manual DAPC08 p.14):

**Futura** es la tipografia oficial para todas las piezas graficas de comunicacion:
titulos, subtitulos, copy y textos destacados en material impreso, presentaciones
y publicaciones digitales. Es sans serif geometrica con extensa familia de pesos
(Light, Book, Medium, Bold, ExtraBold, Heavy y variantes Condensed y Oblique).
Aporta modernidad y caracter diferenciador. En entornos web donde Futura no esta
disponible, usar **Jost** o **Nunito Sans** como sustituto geometrico.

**Calibri** es la tipografia oficial para todos los documentos que se desarrollen
al interior de Socya: informes, memorandos, plantillas internas, correos y
documentos de gestion.

Nunca mezclar Futura con otra fuente de titulares en la misma pieza. El
contraste de peso dentro de la familia (ExtraBold en titulos / Light en cuerpos)
genera jerarquia sin necesidad de fuentes adicionales.

## Layout

Reticulado de 12 columnas, contenedor maximo 1280px, padding lateral 80px en
desktop y 24px en mobile.

Las secciones alternan fondos surface blanco y neutral (#F5F5F5) para crear
ritmo visual. Secciones de alto impacto institucional (hero, footer, CTA
principal) usan primary (#087062) como fondo con texto blanco.

Espaciado entre secciones: 3xl (96px) en desktop, 2xl (64px) en mobile.
Gap interno de tarjetas: lg (24px).

## Elevation & Depth

Sistema de elevacion minimalista — la profundidad es funcional, no decorativa:

- Nivel 0: Sin sombra. Tarjetas sobre fondos neutros en reposo.
- Nivel 1: box-shadow 0 2px 8px rgba(0,0,0,0.07) — hover de tarjetas.
- Nivel 2: box-shadow 0 6px 20px rgba(0,0,0,0.12) — modales y dropdowns.

Prohibido usar sombras sobre el logosimbolo (regla explicita del manual).
No usar difuminados ni efectos de profundidad sobre la marca.

## Shapes

El simbolo organico del logo (curvas de hojas, figura humana fluida) informa
la preferencia visual por formas redondeadas en todo el sistema:

- Botones: rounded.md (8px)
- Tarjetas: rounded.lg (16px)
- Badges y etiquetas: rounded.full (pill)
- Inputs de formulario: rounded.md (8px)
- Imagenes en tarjetas: rounded.lg solo en esquinas superiores

Evitar esquinas completamente rectas (0px) en elementos interactivos. Evitar
radios excesivos (>24px) en contenedores de texto extenso.

## Components

**Boton primario:** Fondo primary (#087062), texto blanco, Futura Bold 0.75rem
uppercase, letter-spacing 0.08em, padding 12x28px, rounded.md. Hover a
primary-dark (#123C49).

**Boton secundario:** Fondo logo-green (#69BE28), texto blanco. Para acciones
afirmativas en contextos donde el verde del logo es apropiado (secciones
ambientales, confirmaciones positivas).

**Boton ghost:** Sin fondo, borde 1.5px primary, texto primary. Hover con fondo
accent-green. Para acciones secundarias junto a un boton primario.

**Divisor de seccion:** Linea horizontal 3px en logo-green (#69BE28). Aparece
al final del header, separando secciones y en el footer (patron del manual).

**Tarjeta de unidad de negocio:** Fondo blanco, rounded.lg, padding 32px,
imagen superior con radio en esquinas superiores, titulo en Futura Bold (h3),
descripcion en Calibri (body-md). Sin sombra en reposo, nivel 1 en hover.

**Navegacion:** Fondo blanco, logo color alineado a la izquierda, links en
on-surface con typography.nav (Futura Medium), activo/hover en primary.
Mobile: hamburger con overlay semi-transparente.

**Footer:** Fondo primary (#087062), logo version blanca, texto y links en
blanco. Grid 4 columnas desktop, 2 tablet, 1 mobile.

**Badge tematico:** Pill redondeado por programa: verde (#69BE28+blanco) para
ambiental, cielo (#00A0DF+blanco) para agua/infraestructura, amarillo
(#F3C400+oscuro) para social/comunitario, naranja (#FF8300+blanco) para
negocios circulares.

## Do's and Don'ts

### Correcto

- Logosimbolo a color unicamente sobre fondo blanco.
- Logosimbolo en escala de grises sobre fondos de color (Pantone Cool Gray 11
  y Pantone Cool Gray 8 segun el manual).
- Reproducir los colores del logo con la mayor fidelidad: Pantone 368 para el
  verde, Pantone Cool Gray 11 para el gris.
- Respetar el area de seguridad del logosimbolo: margen minimo equivalente a
  la mitad de la altura de la letra "a" del logotipo en todos sus lados.
- Tamano minimo del logosimbolo en impresion: 1.5 cm de ancho.
- Futura para piezas graficas, Calibri para documentos internos.
- Validar contraste WCAG AA en toda combinacion texto/fondo.

### Incorrecto (prohibiciones explicitas del Manual DAPC08 p.10)

- No cambiar la posicion del simbolo respecto al logotipo.
- No variar las proporciones del simbolo ni del logotipo entre si.
- No usar el logosimbolo como marca de agua.
- No aplicar sombreados al logosimbolo.
- No usar difuminados de ningun tipo sobre el logosimbolo.
- No ubicar el logosimbolo sobre colores de bajo contraste.
- No modificar la composicion cromatica del logosimbolo (no recolorear).
- No usar el logosimbolo a color sobre fondos distintos al blanco en piezas propias.
- No usar los colores extendidos (sky, lime, yellow, orange) como fondos
  dominantes en secciones principales — son acentos tematicos.
