#!/bin/bash
# sync_to_mac.sh — Pull OneNote archive folders from Cerebro to local Dropbox.
#
# Usage:
#   ./sync_to_mac.sh          — sync 02_Markdown and 03_Summaries (default)
#   ./sync_to_mac.sh all      — sync all three folders including 01_Raw_Audit
#   ./sync_to_mac.sh 01       — sync 01_Raw_Audit only
#   ./sync_to_mac.sh 02       — sync 02_Markdown only
#   ./sync_to_mac.sh 03       — sync 03_Summaries only
#   ./sync_to_mac.sh 02 03    — sync any combination

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

REMOTE_USER="lab-user"
REMOTE_HOST="10.254.254.48"
REMOTE_BASE="Onenote-archivist/onenote_audit"
LOCAL_BASE="/Users/james/Dropbox/Business/OneNote"

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

BOLD="\033[1m"
GREEN="\033[0;32m"
CYAN="\033[0;36m"
YELLOW="\033[0;33m"
RESET="\033[0m"

log()  { echo -e "${CYAN}▶ $*${RESET}"; }
ok()   { echo -e "${GREEN}✓ $*${RESET}"; }
warn() { echo -e "${YELLOW}⚠ $*${RESET}"; }

sync_folder() {
    local folder="$1"
    local src="${REMOTE_USER}@${REMOTE_HOST}:${REMOTE_BASE}/${folder}/"
    local dst="${LOCAL_BASE}/${folder}/"

    log "Syncing ${folder} ..."
    mkdir -p "$dst"
    rsync -az \
        --progress \
        --human-readable \
        --delete \
        --ignore-errors \
        --exclude="*.pyc" \
        --exclude="__pycache__/" \
        --exclude=".DS_Store" \
        "$src" "$dst"

    local exit_code=$?
    if [ $exit_code -eq 0 ]; then
        ok "${folder} complete — ${dst}"
    else
        warn "${folder} finished with errors (exit code ${exit_code})"
    fi
    return $exit_code
}

# ---------------------------------------------------------------------------
# PRE-SYNC: FIX TRAILING SPACES IN NAMES ON CEREBRO
# ---------------------------------------------------------------------------

fix_trailing_spaces() {
    log "Checking for filenames with trailing spaces on Cerebro ..."
    ssh "${REMOTE_USER}@${REMOTE_HOST}" bash <<'ENDSSH'
set -e
BASE="$HOME/Onenote-archivist/onenote_audit"

# Rename directories first (deepest first to avoid path conflicts)
find "$BASE" -depth -type d -name "* " | while IFS= read -r path; do
    dir=$(dirname "$path")
    old=$(basename "$path")
    new="${old%" "}"       # strip one trailing space; loop handles multiples
    new="${new%"${new##*[! ]}"}"  # strip ALL trailing spaces
    new=$(echo "$old" | sed 's/[[:space:]]*$//')
    if [ "$old" != "$new" ]; then
        echo "  Renaming dir:  '$old'  →  '$new'"
        mv "$path" "$dir/$new"
    fi
done

# Then rename files
find "$BASE" -type f -name "* " | while IFS= read -r path; do
    dir=$(dirname "$path")
    old=$(basename "$path")
    new=$(echo "$old" | sed 's/[[:space:]]*$//')
    if [ "$old" != "$new" ]; then
        echo "  Renaming file: '$old'  →  '$new'"
        mv "$path" "$dir/$new"
    fi
done

# Also handle files with trailing space before extension (e.g. "Harris .md")
find "$BASE" -type f | while IFS= read -r path; do
    dir=$(dirname "$path")
    old=$(basename "$path")
    # Strip spaces before the last dot (e.g. "Harris .md" → "Harris.md")
    new=$(echo "$old" | sed 's/[[:space:]]*\(\.[^.]*\)$/\1/')
    if [ "$old" != "$new" ]; then
        echo "  Renaming file: '$old'  →  '$new'"
        mv "$path" "$dir/$new"
    fi
done
ENDSSH
    ok "Filename check complete."
    echo ""
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

echo ""
echo -e "${BOLD}============================================================${RESET}"
echo -e "${BOLD}  OneNote Archive Sync: Cerebro → Mac${RESET}"
echo -e "${BOLD}============================================================${RESET}"
echo -e "  Source : ${REMOTE_USER}@${REMOTE_HOST}:~/${REMOTE_BASE}/"
echo -e "  Dest   : ${LOCAL_BASE}/"
echo ""

# Determine which folders to sync
if [ $# -eq 0 ]; then
    FOLDERS=("01_Raw_Audit" "02_Markdown" "03_Summaries")
elif [ "$1" = "all" ]; then
    FOLDERS=("01_Raw_Audit" "02_Markdown" "03_Summaries")
else
    FOLDERS=()
    for arg in "$@"; do
        case "$arg" in
            01) FOLDERS+=("01_Raw_Audit") ;;
            02) FOLDERS+=("02_Markdown") ;;
            03) FOLDERS+=("03_Summaries") ;;
            01_Raw_Audit|02_Markdown|03_Summaries) FOLDERS+=("$arg") ;;
            *) warn "Unknown folder argument: $arg (use 01, 02, 03, or all)"; exit 1 ;;
        esac
    done
fi

echo -e "  Folders: ${FOLDERS[*]}"
echo ""

fix_trailing_spaces

START=$(date +%s)
ERRORS=0

for folder in "${FOLDERS[@]}"; do
    sync_folder "$folder"
    [ $? -ne 0 ] && ERRORS=$((ERRORS + 1))
    echo ""
done

END=$(date +%s)
ELAPSED=$((END - START))
MINS=$((ELAPSED / 60))
SECS=$((ELAPSED % 60))

echo -e "${BOLD}============================================================${RESET}"
if [ $ERRORS -eq 0 ]; then
    echo -e "${GREEN}${BOLD}  Sync complete — ${MINS}m ${SECS}s${RESET}"
else
    echo -e "${YELLOW}${BOLD}  Sync complete with ${ERRORS} error(s) — ${MINS}m ${SECS}s${RESET}"
fi
echo -e "${BOLD}============================================================${RESET}"
echo ""
