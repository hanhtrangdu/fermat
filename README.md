# Fermat

> *"I have discovered a truly marvelous proof of this, which this margin is too narrow to contain."*
> — Pierre de Fermat, 1637

Like the legendary mathematician who transformed raw numbers into elegant theorems, **Fermat** transforms raw documents into polished, institutional-grade deliverables.

Pierre de Fermat was a 17th-century French mathematician who, despite being an amateur (he was a lawyer by profession), made foundational contributions to number theory, probability, and analytic geometry. His famous "Last Theorem" — scribbled in a margin — took 358 years to prove.

This tool takes a similar approach: **aggressive, opinionated formatting** that enforces Wall Street standards without asking permission.

---

## What It Does

| Input | Output |
|-------|--------|
| `.docx` | Margins, fonts, headings, bullets, tables, filler removal, hedge strengthening |
| `.xlsx` | Row heights, fonts, pastel fills, navigation columns, chart sizing, number formats |

**Philosophy:** Documents should look like they came from a top-tier investment bank, not a college essay.

---

## Installation

```bash
# Clone
git clone https://github.com/1e9-investments/fermat.git
cd fermat

# Set up Python environment
python3 -m venv .venv
source .venv/bin/activate
pip install python-docx openpyxl
```

---

## Usage

```bash
./fermat.sh <filepath> [--no-backup]
```

**Examples:**
```bash
# Format a Word document (auto-backup created)
./fermat.sh report.docx

# Format an Excel workbook without backup
./fermat.sh model.xlsx --no-backup
```

Backups are saved to `./backups/` with timestamps.

---

## Word Formatting (`.docx`)

### Page Setup
- **Margins:** 0.75" all sides
- **Font:** 11pt Aptos (body), 12pt+ for headings

### Color Palette (Forest Green — Wall St Standard)
| Element | Color |
|---------|-------|
| H1 | `#1A3C2B` Deep forest |
| H2 | `#2D6A4F` Mid forest |
| H3 | `#40916C` Medium green |
| Sub-labels | `#52796F` Muted sage (italic) |
| Sources | `#595959` Gray |

### Tables
- Header row: Dark forest green fill, white text
- Alternating rows: Light gray `#F5F5F5`
- Key metrics (Revenue, EBITDA, PAT): Sage highlight `#C2D69B`

### Style Enforcement
**Removes filler phrases:**
- "It is worth noting that..."
- "Basically...", "Essentially...", "Actually..."
- "In order to...", "Due to the fact that..."

**Strengthens hedges:**
- "might be" → "is"
- "seems to be" → "is"
- "could potentially" → "can"

---

## Excel Formatting (`.xlsx`)

### Typography
- **Font:** 11pt Aptos
- **Headers:** Bold, 12pt for titles

### Text Color (Wall Street Data Source Convention)
| Color | Hex | Meaning |
|-------|-----|---------|
| **Blue** | `#0000FF` | Hardcoded values / manual inputs |
| **Black** | `#000000` | Formulas / calculated values |
| **Green** | `#006400` | References to other sheets |
| **Purple** | `#800080` | References to other workbooks |

### Fill Colors (Bright Pastels Only)
| Color | Hex | Use |
|-------|-----|-----|
| Green | `#C6EFCE` | Positive indicators |
| Yellow | `#FFE699` | Assumptions, highlights |
| Blue | `#DEEBF7` | Headers, neutral |
| Orange | `#FCE4D6` | Alerts |

### Assumption Cells
- **Yellow fill** (`#FFE699`) + **medium outside border**
- Applied to hardcoded inputs in "Assumption", "Input", "Driver", "Sensitivity" rows
- Makes key model drivers instantly visible

**No dark fills.** Ever.

### Number Formats
- **Currency:** Accounting format, red negatives
- **Ratios:** `0.0x` with red negatives
- **Percentages:** `0.0%` with red negatives

### Structure
- **Navigation column:** Column A, Wingdings "v" markers at each section (Ctrl+↑/↓ to jump)
- **Freeze panes:** Header row + column before data starts (keeps headers & labels visible)
- **Column widths:** 12 for numeric, 25 for labels, 3 for nav
- **Charts:** 3" tall × 5" wide, standardized

---

## File Structure

```
fermat/
├── fermat.sh          # Entry point
├── format_word.py     # Word document formatter
├── format_excel.py    # Excel workbook formatter
├── standards/         # Formatting specifications
│   ├── word.md
│   └── excel.md
├── backups/           # Auto-generated backups
└── .venv/             # Python virtual environment
```

---

## Requirements

- Python 3.8+
- `python-docx` — Word document manipulation
- `openpyxl` — Excel workbook manipulation

---

## License

MIT

---

*"The precision of mathematics applied to the chaos of corporate documents."*
