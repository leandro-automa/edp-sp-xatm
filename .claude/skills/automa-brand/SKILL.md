---
name: automa-brand
description: >
  Apply Automa's official brand guidelines (September 2023 brandbook) and cultural values to any UI,
  document, PDF, presentation, or artifact built for Automa. Use this skill whenever working on an
  Automa project, building or reviewing UI components for Automa, generating documents or exports
  with Automa branding, auditing brand compliance, or whenever the user mentions Automa colors,
  Automa fonts, Automa brand, or brand guidelines. Also trigger when the user asks to evaluate,
  review, or audit visual consistency of an Automa application. This skill overrides the generic
  brand-guidelines skill for all Automa contexts.
---

# Automa Brand Guidelines

Automa is a DNV company specializing in SCADA/SPCS solutions for the electrical sector (Power & Utilities). Their tagline is "Power & Utilities" and their hashtag is **#energiadainovação**. Every piece of UI, documentation, and communication must reflect their identity: technology with a human touch, innovation as DNA, and excellence in everything.

## 1. Color System

Colors are the most important property of Automa's identity. They must always be present in communications to create a recognizable, consistent brand.

### Primary Colors (must always be present)

| Name | Hex | RGB | Usage |
|---|---|---|---|
| **Azul Escuro Automa** | `#1A3180` | 26, 49, 128 | Primary / navigation / dark backgrounds / headings on dark surfaces |
| **Azul Claro Automa** | `#1142FF` | 17, 66, 255 | Secondary / accent buttons / section headers / interactive elements |
| **Verde Automa** | `#75C800` | 117, 200, 0 | Success states / accent highlights / active indicators / call-to-action |

### Primary Color Tints

Use tints for hover states, disabled states, backgrounds, borders, and subtle variations.

| Base Color | 80% | 60% | 40% | 20% |
|---|---|---|---|---|
| Azul Escuro | `#485A99` | `#7683B3` | — | — |
| Azul Claro | `#4168FF` | `#708EFF` | — | — |
| Verde | `#91D333` | `#ACDE66` | — | — |

### Secondary Colors (restricted use)

These exist only for infographics, illustrations, decorative images, and data visualization where the primary palette is insufficient. They must never replace or overshadow the primary blue/green identity.

| Color | Hex | Use case |
|---|---|---|
| Red | `#D61B1A` | Error states, critical alerts, charts |
| Orange | `#E77A0C` | Warning states, charts |
| Yellow | `#FFCE12` | Caution, highlights in charts |
| Dark Green | `#007D05` | Environmental/sustainability context |
| Cyan | `#04EBFF` | Technology/digital accents |
| Purple | `#800E74` | Charts, categorization |
| Violet | `#A438BC` | Charts, categorization |

### Color Rules

- **NEVER use gradient colors.** All colors must be flat/solid. This is a hard rule — no linear-gradient, no radial-gradient, no color transitions.
- People must always recognize blue and green as Automa's primary colors in any piece.
- On dark backgrounds (Azul Escuro), use white text and Verde for accents.
- On Azul Claro backgrounds, use white text.
- On Verde backgrounds, use white text with the negative logo.
- Avoid placing the colored logo on gray, magenta, yellow, lime, or purple backgrounds — these reduce legibility.

## 2. Typography

### Font Families

| Context | Font | Weights |
|---|---|---|
| **Headings & titles** | Urbanist | Light (300), Regular (400), Bold (700) + Italics |
| **Body text** | Montserrat | Light (300), Regular (400), Bold (700) + Italics |
| **System fallback** | Verdana | Regular, Bold + Italics |

### How to Apply

- All page titles, section headers, card titles, and navigation labels use **Urbanist**.
- All paragraph text, form labels, table content, descriptions, and UI body copy use **Montserrat**.
- When Urbanist/Montserrat are unavailable (e.g., system emails, PowerPoint on machines without the fonts), fall back to **Verdana**.
- Load Urbanist and Montserrat from Google Fonts for web applications.

## 3. Logo Usage

### Versions

- **Preferred**: Full-color logo on white background.
- **Negative**: White logo on Verde (#75C800), Azul Claro (#1142FF), or Azul Escuro (#1A3180) backgrounds.
- **Monochrome**: Black on white, or white on black — only when color printing is unavailable.
- **With tagline**: "Power & Utilities" aligned right below the logo. Requires specific art files.
- **Symbol only**: The stylized "A" can be used alone for small applications (< 35mm print / < 160px digital), app icons, or avatars where the name is already nearby.

### Minimum Sizes

| Version | Print | Digital |
|---|---|---|
| Logo only | 18 mm | 80 px |
| With tagline (stacked) | 35 mm | 160 px |
| With tagline (horizontal) | 50 mm | 220 px |

### Reserve Area

Maintain clear space around the logo equal to the height of the lowercase letter "a" in the wordmark. No other elements should invade this space.

### Prohibited Uses

Never alter proportions, change the official colors, rotate, distort, apply graphic effects (shadows, outlines, 3D), reconstruct/redraw, or use an unapproved version of the logo.

## 4. Graphic Elements

Automa's visual system uses a **curve-shaped graphic** derived from the logo symbol — a swooping shape that evokes the "A" letterform. This creates a proprietary, recognizable look across all communications.

### Rules

- Use only ONE graphic element per piece/page.
- Never rotate the graphic elements.
- Never place graphics in corners — they should flow naturally from edges.
- The graphic can be used for image clipping masks, section dividers, and decorative backgrounds.
- Use brand colors (blue and green curves) for the graphic elements.

## 5. Design Principles

- **Clean and professional**: Minimal visual noise, generous whitespace, clear hierarchy.
- **Flat design**: Solid colors only, no gradients anywhere.
- **Maximum legibility**: Text contrast must be high; avoid busy backgrounds under text.
- **Consistent**: Every piece should be immediately recognizable as Automa.
- **Modern**: Embrace current design trends while maintaining brand identity.

## 6. Cultural Values

Automa's five cultural pillars should inform the tone and feel of interfaces and communications. They don't dictate specific UI patterns, but they should be reflected in how the product feels to use.

### Pessoas (People) — "Gente é tudo para a gente"
- Empathetic UX: clear error messages, helpful empty states, forgiving form validation.
- Respect user time: fast loads, progressive disclosure, sensible defaults.

### Colaboração (Collaboration) — "Somos um time único"
- Teamwork-oriented features: shared views, collaborative workflows, visibility across roles.
- The UI should make collective success easy.

### Inovação (Innovation) — "Inovação está em nosso DNA"
- Embrace modern UI patterns, keep the app feeling current.
- Continuous improvement mindset: the UI should feel like it's evolving.

### Excelência (Excellence) — "Excelência em tudo que fazemos"
- Pixel-perfect attention to detail: consistent spacing, aligned elements, polished interactions.
- Accountability in the UI: clear audit trails, timestamps, ownership indicators.

### Confiança (Trust) — "Confiança acima de tudo"
- Transparency: show system status, sync state, data freshness.
- Ethics: secure handling of data, clear privacy indicators.
- Autonomy: trust users with powerful features, don't over-restrict.

## 7. Brand Compliance Checklist

When auditing or building for Automa, verify:

- [ ] Primary colors (#1A3180, #1142FF, #75C800) are dominant and recognizable
- [ ] No gradients anywhere — all colors are flat/solid
- [ ] Headings use Urbanist font family
- [ ] Body text uses Montserrat font family
- [ ] Verdana is set as the system fallback
- [ ] Logo is used correctly (right version for the background, proper clear space)
- [ ] Only one graphic element per page/view
- [ ] Secondary colors are used sparingly and only for data visualization/decoration
- [ ] Color contrast meets accessibility standards (WCAG AA minimum)
- [ ] The overall feel is clean, modern, professional, and distinctly "Automa"
- [ ] Dark nav/sidebar uses Azul Escuro (#1A3180) with white text
- [ ] Success/active states use Verde (#75C800)
- [ ] Interactive elements (buttons, links) use Azul Claro (#1142FF)
- [ ] Error states use red (#D61B1A), warnings use orange (#E77A0C)

## 8. Web/App Implementation Notes

### CSS Custom Properties (recommended)

```css
:root {
  /* Primary */
  --automa-azul-escuro: #1A3180;
  --automa-azul-claro: #1142FF;
  --automa-verde: #75C800;

  /* Primary tints */
  --automa-azul-escuro-80: #485A99;
  --automa-azul-escuro-60: #7683B3;
  --automa-azul-claro-80: #4168FF;
  --automa-azul-claro-60: #708EFF;
  --automa-verde-80: #91D333;
  --automa-verde-60: #ACDE66;

  /* Secondary (charts/decoration only) */
  --automa-red: #D61B1A;
  --automa-orange: #E77A0C;
  --automa-yellow: #FFCE12;

  /* Typography */
  --font-heading: 'Urbanist', Verdana, sans-serif;
  --font-body: 'Montserrat', Verdana, sans-serif;
}
```

### Vuetify Theme Mapping

For Vue 3 + Vuetify 3 projects:

```typescript
{
  themes: {
    light: {
      colors: {
        primary: '#1A3180',    // Azul Escuro — nav, dark backgrounds
        secondary: '#1142FF',  // Azul Claro — buttons, accent
        accent: '#75C800',     // Verde — success, highlights
        success: '#75C800',
        error: '#D61B1A',
        warning: '#E77A0C',
        info: '#1142FF',
      }
    }
  }
}
```

### PDF/Document Generation

When generating PDFs or documents:
- Header/footer backgrounds: Azul Escuro (#1A3180) with white text
- Section title accents: Azul Claro (#1142FF)
- Highlight/status indicators: Verde (#75C800)
- Body font: Montserrat (embed if possible)
- Title font: Urbanist (embed if possible)
- Include the Automa logo in the correct version for the background color
