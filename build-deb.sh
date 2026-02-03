#!/bin/bash
#
# Build Debian package for Resistance Measuring application
#

set -e

# Package metadata
PACKAGE_NAME="resistance-measuring"
VERSION="1.3"
REVISION="1"
ARCH="all"
MAINTAINER="Resistance Measuring Team <noreply@example.com>"
DESCRIPTION="Resistance measuring software for product testing"
LONG_DESCRIPTION="Product tester software that works with jig tester hardware (Arduino).
 Used in production lines to validate products by measuring resistance values.
 Results are displayed with green (Pass) or red (Fail) indicators."

# Directories
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BUILD_DIR="${SCRIPT_DIR}/${PACKAGE_NAME}_${VERSION}-${REVISION}_${ARCH}"

# Check prerequisites
check_prerequisites() {
    local missing=()

    if ! command -v dpkg-deb &> /dev/null; then
        missing+=("dpkg-deb (install: sudo apt install dpkg)")
    fi

    if ! command -v convert &> /dev/null; then
        missing+=("convert (install: sudo apt install imagemagick)")
    fi

    if [ ${#missing[@]} -ne 0 ]; then
        echo "Error: Missing required tools:"
        for tool in "${missing[@]}"; do
            echo "  - $tool"
        done
        exit 1
    fi

    if [ ! -f "${SCRIPT_DIR}/tester.pyw" ]; then
        echo "Error: tester.pyw not found in ${SCRIPT_DIR}"
        exit 1
    fi

    if [ ! -f "${SCRIPT_DIR}/favicon.ico" ]; then
        echo "Error: favicon.ico not found in ${SCRIPT_DIR}"
        exit 1
    fi
}

# Create package directory structure
create_structure() {
    echo "Creating package structure..."

    rm -rf "${BUILD_DIR}"

    mkdir -p "${BUILD_DIR}/DEBIAN"
    mkdir -p "${BUILD_DIR}/usr/bin"
    mkdir -p "${BUILD_DIR}/usr/share/${PACKAGE_NAME}"
    mkdir -p "${BUILD_DIR}/usr/share/applications"
    mkdir -p "${BUILD_DIR}/usr/share/doc/${PACKAGE_NAME}"

    # Create icon directories for various sizes
    for size in 16 32 48 64 128 256; do
        mkdir -p "${BUILD_DIR}/usr/share/icons/hicolor/${size}x${size}/apps"
    done
}

# Generate control file
create_control_file() {
    echo "Creating control file..."

    cat > "${BUILD_DIR}/DEBIAN/control" << EOF
Package: ${PACKAGE_NAME}
Version: ${VERSION}-${REVISION}
Section: electronics
Priority: optional
Architecture: ${ARCH}
Depends: python3, python3-tk, python3-openpyxl, python3-serial
Maintainer: ${MAINTAINER}
Description: ${DESCRIPTION}
 ${LONG_DESCRIPTION}
EOF
}

# Create launcher script
create_launcher() {
    echo "Creating launcher script..."

    cat > "${BUILD_DIR}/usr/bin/${PACKAGE_NAME}" << 'EOF'
#!/bin/bash
#
# Launcher for Resistance Measuring application
#

APP_NAME="resistance-measuring"
APP_DIR="/usr/share/${APP_NAME}"
CONFIG_DIR="${XDG_CONFIG_HOME:-$HOME/.config}/${APP_NAME}"
DATA_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/${APP_NAME}"

# Create config and data directories if they don't exist
mkdir -p "${CONFIG_DIR}"
mkdir -p "${DATA_DIR}"

# Change to data directory so CSV files are stored there
cd "${DATA_DIR}"

# Copy default config if it doesn't exist
if [ ! -f "${CONFIG_DIR}/config.json" ]; then
    if [ -f "${APP_DIR}/config.json" ]; then
        cp "${APP_DIR}/config.json" "${CONFIG_DIR}/"
    fi
fi

# Create symlink to config in data directory if it doesn't exist
if [ ! -e "${DATA_DIR}/config.json" ]; then
    ln -sf "${CONFIG_DIR}/config.json" "${DATA_DIR}/config.json"
fi

# Run the application
exec python3 "${APP_DIR}/tester.pyw" "$@"
EOF

    chmod 755 "${BUILD_DIR}/usr/bin/${PACKAGE_NAME}"
}

# Create .desktop file
create_desktop_file() {
    echo "Creating .desktop file..."

    cat > "${BUILD_DIR}/usr/share/applications/${PACKAGE_NAME}.desktop" << EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=Resistance Measuring
GenericName=Resistance Tester
Comment=${DESCRIPTION}
Exec=${PACKAGE_NAME}
Icon=${PACKAGE_NAME}
Terminal=false
Categories=Utility;Electronics;
Keywords=resistance;measuring;tester;arduino;
EOF
}

# Convert icons
convert_icons() {
    echo "Converting icons..."

    for size in 16 32 48 64 128 256; do
        convert "${SCRIPT_DIR}/favicon.ico" \
            -resize "${size}x${size}" \
            -background none \
            -gravity center \
            -extent "${size}x${size}" \
            "${BUILD_DIR}/usr/share/icons/hicolor/${size}x${size}/apps/${PACKAGE_NAME}.png"
    done
}

# Copy application files
copy_files() {
    echo "Copying application files..."

    # Copy main application
    cp "${SCRIPT_DIR}/tester.pyw" "${BUILD_DIR}/usr/share/${PACKAGE_NAME}/"

    # Copy documentation
    if [ -f "${SCRIPT_DIR}/README.md" ]; then
        cp "${SCRIPT_DIR}/README.md" "${BUILD_DIR}/usr/share/doc/${PACKAGE_NAME}/"
    fi
}

# Set permissions
set_permissions() {
    echo "Setting permissions..."

    # Ensure proper ownership and permissions
    find "${BUILD_DIR}" -type d -exec chmod 755 {} \;
    find "${BUILD_DIR}" -type f -exec chmod 644 {} \;
    chmod 755 "${BUILD_DIR}/usr/bin/${PACKAGE_NAME}"
}

# Build the package
build_package() {
    echo "Building package..."

    dpkg-deb --build --root-owner-group "${BUILD_DIR}"

    echo ""
    echo "Package built successfully: ${BUILD_DIR}.deb"
    echo ""
    echo "To inspect the package:"
    echo "  dpkg -I ${BUILD_DIR}.deb"
    echo ""
    echo "To install the package:"
    echo "  sudo dpkg -i ${BUILD_DIR}.deb"
    echo "  sudo apt-get install -f  # Install dependencies if needed"
    echo ""
    echo "To uninstall:"
    echo "  sudo dpkg -r ${PACKAGE_NAME}"
}

# Main
main() {
    echo "Building ${PACKAGE_NAME} ${VERSION}-${REVISION} Debian package..."
    echo ""

    check_prerequisites
    create_structure
    create_control_file
    create_launcher
    create_desktop_file
    convert_icons
    copy_files
    set_permissions
    build_package
}

main "$@"
