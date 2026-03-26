---
name: Excel Analysis Format Standard
description: Mandatory formatting rules for all Excel analysis - row heights, fonts, colors, charts, navigation, number formats, currency conventions.
type: feedback
---

# Excel Analysis Formatting Standard

**Why:** Consistent, professional Excel output for institutional-grade analysis. Maximizes readability and navigation.

**How to apply:** Apply to ALL Excel files generated for analysis, models, dashboards.

---

## Layout & Structure

### Row Heights
- Keep uniform/default throughout
- Never vary row heights between sections

### Row 1
- Insert new row if data already exists
- Always leave empty for presentation spacing
- Empty rows between blocks/sections for visual separation

### Column A (Navigation)
- Insert new column A if data already exists
- Keep narrow (~3 characters)
- Wingdings font
- Place "v" markers at each section start — enables quick Ctrl+↑/↓ scrolling between sections
- End with "v" marker + "End Model" label in Column B

### Freeze Panes
- Freeze at **header row** (not necessarily row 1 — find where column headers are)
- Freeze at **column right before data starts** (not necessarily column A — could be B or C if labels exist)
- Result: headers stay visible while scrolling; navigation + label columns stay visible while scrolling right

### Visibility
- Never hide rows or columns
- Use grouping only if needed

### Cell Spanning
- Never use Merge & Center
- Use Center Across Selection for spanning headers
- Use Wrap Text to display long content clearly

## Typography

### Font
- **Aptos** (default)
- **Calibri** fallback if Aptos unavailable
- 11pt throughout - no smaller sizes

### Text Color (Wall Street Data Source Convention)
Color-code cells by data source:
- **Blue (#0000FF):** Hardcoded values / manual inputs
- **Black (#000000):** Formulas / calculated values
- **Green (#006400):** References to other sheets in same workbook
- **Purple (#800080):** References to other workbooks (external links)

### Casing
- Avoid ALL CAPS
- Use Title Case for headers, section labels, chart titles

## Colors

### Fill Colors (Bright Pastels)
- Green: #C6EFCE
- Yellow: #FFE699 (also used for assumptions)
- Blue: #DEEBF7
- Orange: #FCE4D6

### Assumption Cells
- **Yellow fill (#FFE699)** + **medium outside border**
- Applied to hardcoded numeric inputs in rows labeled "Assumption", "Input", "Driver", "Sensitivity", "Scenario"
- Makes key model inputs instantly visible

### Avoid
- Dark fills with white text

## Number Formatting

### Alignment
- Numbers: always right-aligned
- Column headers for numbers: also right-aligned to match

### Percentages
- Always format as percentage (0.0%)
- Never display as plain decimal numbers

### Currency Convention
- Vietnamese Dong: VNDk, VNDmm, VNDbn
- US Dollars: US$k, US$mm, US$bn
- Never capitalized M or B (use mm, bn)

## Charts

### Size
- Always 3 inches tall x 5 inches wide (216pt x 360pt)

### Fonts
- Font: Aptos
- Title: 12pt bold
- Axis labels: 10pt
- Data labels: 10pt

### Gridlines
- Always remove major gridlines

### Data Labels
- Always show value labels on all series
- Hide axes when data labels make them redundant

### Percentages in Charts
- Format as 0.0%
- Never display as plain numbers

## Punctuation

- Never use long dashes (em dashes)
- Use short dashes (hyphens) instead
