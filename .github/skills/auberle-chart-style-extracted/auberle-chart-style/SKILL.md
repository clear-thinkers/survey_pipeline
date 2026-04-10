---
name: auberle-chart-style
description: >
  Apply the Auberle / 412 Youth Zone report visualization style whenever the user asks
  to make a chart, table, or data visual "in Auberle style", "in report style", or
  "consistent with the Youth Zone report". Also trigger when the user says "use our
  house style", "reuse the template", or "same style as before" in the context of
  data visualization. Produces exportable PNG files that match the organization's
  established visual identity.
---

# Auberle Report Visualization Style Guide

Use this skill to produce charts, tables, and data visuals that match the design
language established in the 412 Youth Zone annual survey reports.

---

## Core Design Principles

- **White background** — always `background: #fff`, never transparent or colored containers
- **Clean and minimal** — no gradients, no shadows, no decorative effects
- **Flat, editorial feel** — generous whitespace, thin borders, restrained color use
- **Font stack** — `'Segoe UI', Arial, sans-serif` (system sans-serif, not display fonts)
- **Two font weights only** — 400 (body) and 600–700 (labels, headings, highlighted values)

---

## Color Palette

Use color sparingly and only to encode meaning.

| Role              | Hex       | Usage                                      |
|--------------------|-----------|--------------------------------------------|
| **Blue**          | `#185FA5` | Primary bar/accent color; also used for max highlight when min/max coding is on |
| **Orange**        | `#C2600A` | Secondary accent; also used for min highlight when min/max coding is on |
| **Dark text**     | `#111`    | Title, current-period values               |
| **Body text**     | `#333`    | Table cell values, general data            |
| **Muted text**    | `#888` / `#999` | Column headers, n-counts, footnotes  |
| **Border light**  | `#ebebeb` | Row dividers                               |
| **Border mid**    | `#e0e0e0` | Section dividers                           |
| **Border strong** | `#111`    | Header underline (2px)                     |
| **Sparkline line**| `#bbb`    | Trend line stroke                          |
| **Dot neutral**   | `#ccc` / `#555` | Non-highlighted sparkline dots       |

**Min/max color coding is opt-in.** Do NOT apply blue/orange highlights for high/low values unless the user explicitly asks for it (e.g. "highlight the min and max" or "color-code highs and lows"). By default, use `#185FA5` uniformly for all bars.
---

## Typography

| Element              | Size  | Weight | Color   |
|----------------------|-------|--------|---------|
| Table title          | 13px  | 700    | `#111`  |
| Column header (upper)| 11px  | 500    | `#888`, uppercase, letter-spacing 0.04em |
| Column header (lower)| 12px  | 500–600| `#444`  |
| n-count row          | 11px  | 400    | `#999`  |
| Data cells           | 13–14px| 400   | `#333`  |
| Current-period cells | 13–14px| 500–600| `#111` |
| Highlighted min/max  | same  | 600–700| blue / orange (only when explicitly requested) |
| Footnote             | 11–12px| 400   | `#aaa`  |

**Title casing:** Title Case for chart/table titles (e.g. "Satisfaction Ratings for Youth Coaches Over Time").
All other labels use sentence case.

---

## Table Structure

### Header pattern for multi-period tables
```
[ Row label col ]  [ Merged header: "% Often or all of the time"  ]  [ Trend ]
[ "My coach…"   ]  [ Sep-19 | Mar-22 | Feb-23 | Feb-24 | Feb-25 | Mar-26 ]  [ Trend ]
```
- The spanning label ("% Often or all of the time") uses `colspan` to merge across all data columns
- The sub-header row has `border-bottom: 2px solid #111`
- The n-count row sits between the header and data rows, separated by `border-bottom: 1px solid #e0e0e0`
- Data rows use `border-bottom: 0.5px solid #ebebeb`; last row has no border

### Color coding rule (opt-in only)
Only apply min/max color coding when the user explicitly requests it.
When requested:
```python
mx = max(row_data)
mn = min(row_data)
# Color the max bar/cell #185FA5, the min bar/cell #C2600A, everything else #185FA5
```
By default, all bars use `#185FA5` uniformly.

---

## Sparkline Trend Lines

Include a sparkline column in multi-period charts when using matplotlib. Render as a small inset axes or a separate subplot column (width ~0.6 inches).

```python
# Standard sparkline pattern (matplotlib)
ax_spark.plot(x_vals, y_vals, color='#bbb', linewidth=1.5)
ax_spark.scatter(x_vals[:-1], y_vals[:-1], color='#ccc', s=12, zorder=3)
ax_spark.scatter([x_vals[-1]], [y_vals[-1]], color='#555', s=28, zorder=4)
# Last dot: #185FA5 if max, #C2600A if min (only when min/max coding is on), else #555
ax_spark.axis('off')
```

---

## Chart Types & When to Use

| Type | Use case | Key style notes |
|------|----------|-----------------|
| **Data table + sparklines** | Multi-period longitudinal data (coach ratings, banking trends) | See table structure above |
| **Simple bar chart** | Single-period comparisons (job barriers, reasons for leaving) | Horizontal bars, `#185FA5` fill, `#ebebeb` gridlines only, no axis borders |
| **Grouped bar chart** | Comparisons across 2–3 subgroups | Blue + orange as the two primary group colors |
| **Donut / pie** | Part-of-whole (housing status, license status) | Blue dominant slice, orange secondary, gray for remaining |
| **Stacked bar** | Composition over time or across groups | Blue + orange + `#bbb` neutral |

---

## Output

All output is a **PNG image file** written directly with matplotlib:

```python
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

fig.savefig(str(out_path), dpi=150, bbox_inches="tight")
plt.close(fig)
```

- Save to `output/412YZ/charts/` using the `chart_NN_description.png` naming convention.
- Standard figure width: 7–8 inches. Height: sized to content (allow ~0.52 in per bar row).
- Do NOT produce HTML, CSS, or intermediate files.

---

## Quick-Reference Prompt Pattern

When a user says:

> "Make a [chart type] of [data] in Auberle style"

Apply this checklist:
- [ ] White background (`fig.patch.set_facecolor('white')`)
- [ ] Title in Title Case, bold, `#111`, 13px
- [ ] Merged span header if multi-period
- [ ] All bars `#185FA5` by default — blue/orange min/max only if explicitly requested
- [ ] Sparkline column if multi-period
- [ ] Footnote/subtitle text in `#aaa`, 9–11px
- [ ] Save as PNG via `fig.savefig(..., dpi=150, bbox_inches='tight')`; no HTML output

---

## Example Invocations

- *"Make a horizontal bar chart of job barriers in Auberle style"*
- *"Make a housing status donut chart in Auberle style"*
- *"Same style as before — employment by age grouped bars"*
- *"Rebuild the banking table with sparklines, Auberle style, export PNG"*
