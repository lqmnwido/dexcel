#!/usr/bin/env bash
#
# install.sh — Dexcel installer for macOS / Linux
#
# Install:
#   curl -fsSL https://raw.githubusercontent.com/luqmanhafiz81/dexcel/main/install.sh | sh
#
# Run after install:
#   dexcel

set -uo pipefail

RELEASE_BASE_URL="${DEXCEL_RELEASE_URL:-https://github.com/luqmanhafiz81/dexcel/releases/latest/download}"
MIN_PYTHON_MINOR=8

INSTALL_DIR="$HOME/.dexcel"
APP_DIR="$INSTALL_DIR/app"
VENV_DIR="$INSTALL_DIR/venv"
LOG_DIR="$INSTALL_DIR/logs"
BIN_DIR="$HOME/.local/bin"
SHIM_PATH="$BIN_DIR/dexcel"

TIMESTAMP="$(date +%Y%m%d_%H%M%S)"
mkdir -p "$LOG_DIR"
INSTALL_LOG="$LOG_DIR/install_${TIMESTAMP}.log"

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*" >> "$INSTALL_LOG"
}

info() {
    echo "$*"
    log "$*"
}

fail() {
    echo ""
    echo "Error: $*"
    echo "Log: $INSTALL_LOG"
    log "FATAL: $*"
    exit 1
}

INSTALL_COMPLETE=0
HAD_PRIOR_INSTALL=0
[ -x "$VENV_DIR/bin/python" ] && HAD_PRIOR_INSTALL=1

cleanup_on_failure() {
    if [ "$INSTALL_COMPLETE" -ne 1 ] && [ "$HAD_PRIOR_INSTALL" -ne 1 ]; then
        log "Rolling back incomplete fresh install: removing $INSTALL_DIR"
        rm -rf "$INSTALL_DIR"
    fi
}

trap cleanup_on_failure EXIT

info "Installing Dexcel..."
log "Install dir: $INSTALL_DIR"

OS="unknown"

if [[ "${OSTYPE:-}" == "linux-gnu"* ]]; then
    OS="linux"
elif [[ "${OSTYPE:-}" == "darwin"* ]]; then
    OS="macos"
fi

log "Detected OS: $OS"

python_is_usable() {
    command -v "$1" >/dev/null 2>&1 || return 1

    "$1" -c "
import sys
ok = sys.version_info[0] == 3 and sys.version_info[1] >= $MIN_PYTHON_MINOR
sys.exit(0 if ok else 1)
" 2>/dev/null
}

find_python() {
    for candidate in python3 python; do
        if python_is_usable "$candidate"; then
            echo "$candidate"
            return 0
        fi
    done

    return 1
}

PYTHON_BIN="$(find_python || true)"

if [ -z "$PYTHON_BIN" ]; then
    info "Python 3.${MIN_PYTHON_MINOR}+ not found. Installing Python..."

    if [ "$OS" = "linux" ]; then
        if command -v apt >/dev/null 2>&1; then
            sudo apt update >>"$INSTALL_LOG" 2>&1 || fail "apt update failed"
            sudo apt install -y python3 python3-pip python3-venv unixodbc unixodbc-dev >>"$INSTALL_LOG" 2>&1 \
                || fail "apt install failed"
        elif command -v dnf >/dev/null 2>&1; then
            sudo dnf install -y python3 python3-pip unixODBC unixODBC-devel >>"$INSTALL_LOG" 2>&1 \
                || fail "dnf install failed"
        elif command -v yum >/dev/null 2>&1; then
            sudo yum install -y python3 python3-pip unixODBC unixODBC-devel >>"$INSTALL_LOG" 2>&1 \
                || fail "yum install failed"
        elif command -v pacman >/dev/null 2>&1; then
            sudo pacman -Sy --noconfirm python python-pip unixodbc >>"$INSTALL_LOG" 2>&1 \
                || fail "pacman install failed"
        else
            fail "No supported package manager found. Install Python 3.${MIN_PYTHON_MINOR}+ manually."
        fi
    elif [ "$OS" = "macos" ]; then
        if ! command -v brew >/dev/null 2>&1; then
            info "Homebrew not found. Installing Homebrew..."

            /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" >>"$INSTALL_LOG" 2>&1 \
                || fail "Homebrew installation failed"

            if [ -x /opt/homebrew/bin/brew ]; then
                eval "$(/opt/homebrew/bin/brew shellenv)"
            fi

            if [ -x /usr/local/bin/brew ]; then
                eval "$(/usr/local/bin/brew shellenv)"
            fi
        fi

        brew install python unixodbc >>"$INSTALL_LOG" 2>&1 || fail "brew install failed"
    else
        fail "Unsupported OS. Install Python manually and try again."
    fi

    PYTHON_BIN="$(find_python || true)"
    [ -n "$PYTHON_BIN" ] || fail "Python installed but not found on PATH."
else
    if [ "$OS" = "linux" ]; then
        if command -v apt >/dev/null 2>&1; then
            sudo apt install -y unixodbc unixodbc-dev >>"$INSTALL_LOG" 2>&1 || true
        elif command -v dnf >/dev/null 2>&1; then
            sudo dnf install -y unixODBC unixODBC-devel >>"$INSTALL_LOG" 2>&1 || true
        elif command -v yum >/dev/null 2>&1; then
            sudo yum install -y unixODBC unixODBC-devel >>"$INSTALL_LOG" 2>&1 || true
        elif command -v pacman >/dev/null 2>&1; then
            sudo pacman -Sy --noconfirm unixodbc >>"$INSTALL_LOG" 2>&1 || true
        fi
    elif [ "$OS" = "macos" ] && command -v brew >/dev/null 2>&1; then
        brew install unixodbc >>"$INSTALL_LOG" 2>&1 || true
    fi
fi

info "Using Python: $("$PYTHON_BIN" --version 2>&1)"

mkdir -p "$APP_DIR"

download() {
    local url="$1"
    local dest="$2"

    if command -v curl >/dev/null 2>&1; then
        curl -fsSL "$url" -o "$dest"
    elif command -v wget >/dev/null 2>&1; then
        wget -q "$url" -O "$dest"
    else
        fail "Neither curl nor wget found."
    fi
}

info "Downloading Dexcel application files..."

download "$RELEASE_BASE_URL/db_to_excel.py" "$APP_DIR/db_to_excel.py.new" >>"$INSTALL_LOG" 2>&1 \
    || fail "Failed to download db_to_excel.py"

download "$RELEASE_BASE_URL/requirements.txt" "$APP_DIR/requirements.txt.new" >>"$INSTALL_LOG" 2>&1 \
    || fail "Failed to download requirements.txt"

mv "$APP_DIR/db_to_excel.py.new" "$APP_DIR/db_to_excel.py"
mv "$APP_DIR/requirements.txt.new" "$APP_DIR/requirements.txt"

venv_is_healthy() {
    [ -x "$VENV_DIR/bin/python" ] && "$VENV_DIR/bin/python" -c "import sys" >/dev/null 2>&1
}

if venv_is_healthy; then
    info "Existing Dexcel environment found. Updating it."
else
    [ -d "$VENV_DIR" ] && rm -rf "$VENV_DIR"

    info "Creating isolated Python environment..."
    "$PYTHON_BIN" -m venv "$VENV_DIR" >>"$INSTALL_LOG" 2>&1 \
        || fail "Failed to create virtual environment."
fi

VENV_PY="$VENV_DIR/bin/python"
venv_is_healthy || fail "Virtual environment is not working."

"$VENV_PY" -m pip install --upgrade pip >>"$INSTALL_LOG" 2>&1 \
    || fail "Failed to upgrade pip."

info "Installing Dexcel dependencies..."
"$VENV_PY" -m pip install -r "$APP_DIR/requirements.txt" >>"$INSTALL_LOG" 2>&1 \
    || fail "Failed to install dependencies."

"$VENV_PY" -c "import pandas, openpyxl" >>"$INSTALL_LOG" 2>&1 \
    || fail "Core packages failed verification."

declare -A DRIVER_MODULE=(
    ["MySQL/MariaDB"]="pymysql"
    ["PostgreSQL"]="psycopg2"
    ["SQL Server"]="pyodbc"
    ["Oracle"]="oracledb"
)

WORKING=()
BROKEN=()

for label in "${!DRIVER_MODULE[@]}"; do
    module="${DRIVER_MODULE[$label]}"

    if "$VENV_PY" -c "import $module" >>"$INSTALL_LOG" 2>&1; then
        WORKING+=("$label")
    else
        BROKEN+=("$label")
    fi
done

mkdir -p "$BIN_DIR"

cat > "$SHIM_PATH" << EOF
#!/usr/bin/env bash
exec "$VENV_PY" "$APP_DIR/db_to_excel.py" "\$@"
EOF

chmod +x "$SHIM_PATH"

PATH_LINE='export PATH="$HOME/.local/bin:$PATH"'
NEEDS_PATH_NOTE=0

add_path_to_rc() {
    local rc="$1"

    [ -f "$rc" ] || touch "$rc"

    if ! grep -qF "$PATH_LINE" "$rc" 2>/dev/null; then
        {
            echo ""
            echo "# Added by Dexcel installer"
            echo "$PATH_LINE"
        } >> "$rc"

        log "Added PATH update to $rc"
    fi
}

case "${SHELL:-}" in
    */zsh)
        add_path_to_rc "$HOME/.zshrc"
        ;;
    */bash)
        add_path_to_rc "$HOME/.bashrc"

        if [ "$OS" = "macos" ]; then
            add_path_to_rc "$HOME/.bash_profile"
        fi
        ;;
    *)
        add_path_to_rc "$HOME/.profile"
        ;;
esac

case ":$PATH:" in
    *":$BIN_DIR:"*)
        ;;
    *)
        NEEDS_PATH_NOTE=1
        ;;
esac

INSTALL_COMPLETE=1
trap - EXIT

echo ""
echo "Dexcel installed successfully."

if [ "${#WORKING[@]}" -gt 0 ]; then
    echo "Working database drivers: ${WORKING[*]}"
fi

if [ "${#BROKEN[@]}" -gt 0 ]; then
    echo "Unavailable database drivers: ${BROKEN[*]}"
    echo "See log: $INSTALL_LOG"
fi

echo ""

if [ "$NEEDS_PATH_NOTE" -eq 1 ]; then
    echo "Open a NEW terminal, then run:"
    echo "  dexcel"
else
    echo "Run:"
    echo "  dexcel"
fi