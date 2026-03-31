#!/usr/bin/env bash
set -euo pipefail

APP_NAME="题库录入服务"
DIST_DIR="dist"
APP_PATH="${DIST_DIR}/${APP_NAME}.app"

if [[ ! -d "${APP_PATH}" ]]; then
  echo "Missing app bundle: ${APP_PATH}" >&2
  exit 1
fi

VERSION="${1:-}"
if [[ -z "${VERSION}" ]]; then
  VERSION="unknown"
fi

BUILD_DIR="build"
DMG_STAGE="${BUILD_DIR}/dmg-stage"
DMG_NAME="${APP_NAME}-macOS-${VERSION}.dmg"

rm -rf "${DMG_STAGE}"
mkdir -p "${DMG_STAGE}"

cp -R "${APP_PATH}" "${DMG_STAGE}/"
ln -sf /Applications "${DMG_STAGE}/Applications"

rm -f "${DMG_NAME}"
hdiutil create -volname "${APP_NAME}" -srcfolder "${DMG_STAGE}" -ov -format UDZO "${DMG_NAME}"
echo "Created ${DMG_NAME}"

