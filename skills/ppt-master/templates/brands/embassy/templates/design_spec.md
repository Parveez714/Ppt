---
brand_id: embassy
kind: brand
summary: Embassy Developments — investor updates, internal town halls, customer/buyer pitches, partner and corporate decks, project intros
primary_color: "#000000"
---

# Embassy Developments Brand Specification

> Identity-only preset. Pages are composed freely under these constraints.

> **SVG authoring hard rule:** Never use `text-transform` in any SVG `<text>` or `<tspan>` element — it is an unregistered property that blocks PPTX export. Write all uppercase text with the string uppercased directly in the SVG source.
> Governing principle: **Quiet Authority** — restrained, uncluttered, confident.
> "Embassy doesn't settle at the benchmark; it quietly redefines it."

## I. Brand Overview

| Property | Value |
|---|---|
| Brand Name | Embassy Developments |
| Use Cases | Investor updates, town halls, buyer pitches, partner decks, project intros |
| Tone | Honest. Thoughtful. Direct. Frank Curator. |
| Canvas | 1280×720 (16:9) default |

## II. Color Scheme

| Role | HEX | Use |
|---|---|---|
| primary / ink | `#000000` | Primary text, dark backgrounds, logo on light |
| bg / paper | `#FFFFFF` | Primary background, text on dark |
| grey-mid | `#666567` | Secondary text, captions, dividers, labels |
| grey-light | `#AAA9A7` | Hairlines, subtle fills, disabled/secondary detail |
| off-white | `#F4F3F1` | Optional warm panel background — use sparingly |
| accent / orange | `#F58220` | ≤20% per slide — one highlighted word, one rule, one chart series |

**Background strategy:**
- Content slides → light (white `#FFFFFF` or off-white `#F4F3F1`, black text `#000000`)
- Cover and section dividers → dark (pure black `#000000`, white text `#FFFFFF`, logo reversed to white)
- Closing / Thank You slide → dark (pure black `#000000`, white text `#FFFFFF`, logo reversed to white)
- Orange appears as a single moment per slide only — never as a background or large fill

**Banned:** gradient backgrounds, glows, drop shadows, decorative borders, rounded accent-cards, additional hues or tints.

## III. Typography

| Role | Family | Weight | Size (16:9) | Tracking | Case |
|---|---|---|---|---|---|
| display (cover) | Poppins | 500 | 64–80px | +1px | Write text IN UPPERCASE in SVG |
| slide title | Poppins | 500 | 44–52px | +1.5px | Write text IN UPPERCASE in SVG |
| section-divider title | Poppins | 300 | 56–64px | +2px | Write text IN UPPERCASE in SVG |
| subtitle / kicker | Poppins | 600 | 18–22px | +3px | Write text IN UPPERCASE in SVG |
| body | Poppins | 400 | 22–26px | 0 | Sentence case |
| body small / caption | Poppins | 400 | 16–20px | +0.3px | Sentence case |
| big stat / figure | Poppins | 300 | 80–120px | 0 | — |
| footer / page number | Poppins | 400 | 14–16px | +0.5px | Sentence case / uppercase |

Line-height: ~1.15 for headings, ~1.45 for body. Never set body below 16px.

**Headings echo the logo:** Write heading text content IN UPPERCASE directly in the SVG `<text>` element — do NOT use `text-transform: uppercase` CSS.

## IV. Logo Placement & Clear Space

- **Placement:** Top-Right corner of the slide on ALL slides.
- **Top-right bounds (1280x720 canvas):** `x="1090" y="28" width="170" height="39"` (keeping 80px right margin safe zone).
- **Background matching:**
  - White logo (`embassy_logo_white.svg`) on dark slides (Cover, Section dividers, Closing/Thank You).
  - Black logo (`embassy_logo_black.svg`) on light slides (Content slides).
- **Safe Zone:** Keep all text and shapes clear of `x >= 1060` in the top area `y <= 90` to avoid crowding the logo.

## V. Slide-Type Recipes (Must Include Front Page & Thank You Page)

1. **Front Page / Cover Slide (Slide 1):**
   - Background: Pure Black (`#000000`).
   - Logo: Top-right corner (white logo).
   - Display Title: 64–80px, Poppins 500, Pure White (`#FFFFFF`), UPPERCASE.
   - Subtitle: 20–24px, Poppins 400, Grey Mid (`#666567`) or Accent Orange (`#F58220`).
   - Metadata: Date / Author in footer, 14–16px, `#666567`.

2. **Standard Content Slides (Slides 2 to N-1):**
   - Background: Pure White (`#FFFFFF`).
   - Logo: Top-right corner (black logo).
   - Kicker / Category: 16–18px, Poppins 600, Accent Orange (`#F58220`), UPPERCASE.
   - Slide Title: 44–50px, Poppins 500, Pure Black (`#000000`), UPPERCASE.
   - Body & Visuals: Clean grid, cards, stats, or metrics using `#000000`, `#666567`, `#AAA9A7`, and subtle `#F58220` highlights.
   - Footer: Dynamic copyright line bottom-left, page number bottom-right.

3. **Closing / Thank You Slide (Last Slide):**
   - Background: Pure Black (`#000000`).
   - Logo: Top-right corner (white logo).
   - Headline: "THANK YOU", 64–80px, Poppins 500, Pure White (`#FFFFFF`), UPPERCASE.
   - Contact / Corporate details: 20–24px, Poppins 400, Grey Mid (`#666567`).
   - Footer: Dynamic copyright line bottom-left, pure white/grey.

## VI. Dynamic Copyright & Footer

- Dynamic copyright text: `© 2026. Embassy Developments Limited. All rights reserved.`
- Placed bottom-left (`x="80" y="680"`), font-size 14px, Poppins 400, color `#666567` (or `#AAA9A7` on black background).
