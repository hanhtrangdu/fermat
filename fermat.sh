#!/bin/bash
# Fermat - Aggressive document formatter for 1e9-Investments
# Usage: fermat.sh <filepath> [--no-backup]

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$SCRIPT_DIR/.venv/bin/python"
BACKUP_DIR="$SCRIPT_DIR/backups"

if [ -z "$1" ]; then
    echo "Usage: fermat.sh <filepath> [--no-backup]"
    echo "Supported: .docx, .xlsx"
    exit 1
fi

FILEPATH="$1"
NO_BACKUP="${2:-}"

if [ ! -f "$FILEPATH" ]; then
    echo "Error: File not found: $FILEPATH"
    exit 1
fi

# Get file extension
EXT="${FILEPATH##*.}"
EXT_LOWER=$(echo "$EXT" | tr '[:upper:]' '[:lower:]')

# Create backup unless --no-backup
if [ "$NO_BACKUP" != "--no-backup" ]; then
    TIMESTAMP=$(date +%Y%m%d_%H%M%S)
    BASENAME=$(basename "$FILEPATH")
    BACKUP_PATH="$BACKUP_DIR/${BASENAME%.${EXT}}_${TIMESTAMP}.${EXT}"
    mkdir -p "$BACKUP_DIR"
    cp "$FILEPATH" "$BACKUP_PATH"
    echo "Backup: $BACKUP_PATH"
fi

case "$EXT_LOWER" in
    docx)
        echo "Formatting Word document..."
        "$VENV" "$SCRIPT_DIR/format_word.py" "$FILEPATH"
        ;;
    xlsx)
        echo "Formatting Excel workbook..."
        "$VENV" "$SCRIPT_DIR/format_excel.py" "$FILEPATH"
        ;;
    doc)
        echo "Error: .doc format not supported. Convert to .docx first."
        exit 1
        ;;
    xls)
        echo "Error: .xls format not supported. Convert to .xlsx first."
        exit 1
        ;;
    *)
        echo "Error: Unsupported file type: .$EXT"
        echo "Supported: .docx, .xlsx"
        exit 1
        ;;
esac

echo "Done: $FILEPATH"
