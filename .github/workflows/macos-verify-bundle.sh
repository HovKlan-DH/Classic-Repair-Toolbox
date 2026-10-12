#!/usr/bin/env bash
# Verifies one architecture's macOS release (build-and-release.yml, macOS job) and fails the
# release unless BOTH apps a user can end up with are what macOS needs to let CRT on the local
# network (reported 2026-10-12: "No route to host" with "Local Network" switched on):
#
#   installer - the app the Setup .pkg installs (a fresh install)
#   update    - the app Velopack's in-app updater installs from the full .nupkg (every user who
#               already has CRT)
#
# For each: CRT's bundle ID, executable, X.Y.Z version and NSLocalNetworkUsageDescription in the
# Info.plist; the data pack in Contents/Resources, not Contents/MacOS; the executable's own
# LC_UUID (macos-macho-uuid.py); a signature under the bundle ID that binds the Info.plist, with
# the hardened runtime and exactly the entitlements in Assets/macOS/CRT.entitlements. The two
# apps' Info.plist, executable and resource seal must be the same bytes.
#
# codesign's deep check must pass for the installer. For the update it is only reported: a zip
# keeps no extended attributes, and codesign keeps the signature of every non-code file under
# Contents/MacOS (the .NET assemblies) in one. Whether macOS's Local Network check minds that
# is not known - see "Release process" in .claude/CLAUDE.md.
#
# Finally the installed app is STARTED with one of Velopack's fast-exit hooks
# (--veloapp-obsolete <version>): the .NET runtime and the JIT run under the hardened runtime,
# Velopack exits with 0 before Avalonia opens a window. An x64 app is skipped on a runner
# without Rosetta.
#
# Usage: macos-verify-bundle.sh <arm64|x64> <velopack output folder>
# Environment: PACK_ID, EXE_UNIX, VERSION, MAC_BUNDLE_ID, MAC_ENTITLEMENTS, EXPECTED_UUID
set -euo pipefail
shopt -s nullglob

ARCH="$1"
OUTPUT_DIR="$2"
: "${PACK_ID:?}" "${EXE_UNIX:?}" "${VERSION:?}" "${MAC_BUNDLE_ID:?}" "${MAC_ENTITLEMENTS:?}" "${EXPECTED_UUID:?}"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORK="${RUNNER_TEMP:-/tmp}/verify-macos-$ARCH"
APP_NAME="Classic Repair Toolbox.app"

fail() {
  echo "::error::$ARCH: $*"
  exit 1
}

rm -rf "$WORK"
mkdir -p "$WORK"

# ----------------------------------------------------------------------------------------------
# The two apps
# ----------------------------------------------------------------------------------------------
PKG="$OUTPUT_DIR/$PACK_ID-osx-$ARCH-Setup.pkg"
[ -f "$PKG" ] || fail "no installer at $PKG"
pkgutil --expand-full "$PKG" "$WORK/pkg" || fail "pkgutil could not expand $PKG"
APPS=("$WORK"/pkg/*/Payload/*.app)
[ ${#APPS[@]} -eq 1 ] || fail "expected one .app in $PKG, found ${#APPS[@]}"
INSTALLED="${APPS[0]}"

# Velopack's updater unpacks lib/app/ of the full .nupkg as the bundle and turns each
# "<name>.__symlink" file, which holds the link's target, back into the link.
NUPKG="$OUTPUT_DIR/$PACK_ID-$VERSION-osx-$ARCH-full.nupkg"
[ -f "$NUPKG" ] || fail "no update package at $NUPKG"
unzip -q "$NUPKG" -d "$WORK/nupkg" || fail "could not unzip $NUPKG"
[ -d "$WORK/nupkg/lib/app/Contents" ] || fail "no lib/app/Contents in $NUPKG"
mkdir -p "$WORK/update"
UPDATED="$WORK/update/$APP_NAME"
mv "$WORK/nupkg/lib/app" "$UPDATED"
while IFS= read -r -d '' stored; do
  ln -s "$(cat "$stored")" "${stored%.__symlink}"
  rm "$stored"
done < <(find "$UPDATED" -name '*.__symlink' -print0)

# ----------------------------------------------------------------------------------------------
# What each must hold
# ----------------------------------------------------------------------------------------------
plist_value() {
  plutil -extract "$2" raw "$1" 2>/dev/null || echo "<missing>"
}

check_app() {
  local what="$1" app="$2"
  local plist="$app/Contents/Info.plist"
  local exe="$app/Contents/MacOS/$EXE_UNIX"
  local value signature

  echo "== $ARCH $what: $app"
  [ -f "$plist" ] || fail "$what: no Contents/Info.plist"
  [ -f "$exe" ] || fail "$what: no Contents/MacOS/$EXE_UNIX"

  value=$(plist_value "$plist" CFBundleIdentifier)
  [ "$value" = "$MAC_BUNDLE_ID" ] || fail "$what: CFBundleIdentifier is '$value', expected '$MAC_BUNDLE_ID'"
  value=$(plist_value "$plist" CFBundleExecutable)
  [ "$value" = "$EXE_UNIX" ] || fail "$what: CFBundleExecutable is '$value', expected '$EXE_UNIX'"
  value=$(plist_value "$plist" CFBundleVersion)
  [ "$value" = "${VERSION%%-*}" ] || fail "$what: CFBundleVersion is '$value', expected '${VERSION%%-*}'"
  value=$(plist_value "$plist" NSLocalNetworkUsageDescription)
  [ "$value" != "<missing>" ] || fail "$what: Info.plist has no NSLocalNetworkUsageDescription"

  [ -d "$app/Contents/Resources/Data" ] || fail "$what: no data pack in Contents/Resources/Data"
  [ ! -e "$app/Contents/MacOS/Data" ] || fail "$what: a data pack in Contents/MacOS - it belongs in Contents/Resources"

  value=$(python3 "$SCRIPT_DIR/macos-macho-uuid.py" read "$exe") || fail "$what: could not read the executable's LC_UUID"
  [ "$value" = "$EXPECTED_UUID" ] || fail "$what: the executable's LC_UUID is $value, expected $EXPECTED_UUID"

  if ! signature=$(codesign --display --verbose=2 "$app" 2>&1); then
    echo "$signature"
    fail "$what: the app bundle is not signed"
  fi
  echo "$signature"
  grep -qxF "Identifier=$MAC_BUNDLE_ID" <<< "$signature" || fail "$what: not signed under '$MAC_BUNDLE_ID'"
  grep -q '^Info.plist entries=' <<< "$signature" || fail "$what: the signature does not bind the Info.plist"
  grep -q '^CodeDirectory .*flags=.*runtime' <<< "$signature" || fail "$what: not signed with the hardened runtime"

  codesign --display --entitlements - --xml "$app" > "$WORK/entitlements.xml" 2>/dev/null \
    || fail "$what: could not read the signed entitlements"
  python3 - "$MAC_ENTITLEMENTS" "$WORK/entitlements.xml" <<'PY' || fail "$what: the signed entitlements differ from $MAC_ENTITLEMENTS"
import plistlib, sys
wanted = plistlib.load(open(sys.argv[1], "rb"))
signed = plistlib.load(open(sys.argv[2], "rb"))
if wanted != signed:
    print("wanted:", wanted)
    print("signed:", signed)
    sys.exit(1)
PY
}

deep_check() {
  local what="$1" app="$2" required="$3"
  if codesign --verify --deep --strict --verbose=2 "$app"; then
    echo "$ARCH $what: codesign's deep check passed"
  elif [ "$required" = required ]; then
    fail "$what: codesign's deep check fails (see the log above)"
  else
    echo "::warning::$ARCH $what: codesign's deep check fails (see the log above) - expected while the .NET assemblies sit in Contents/MacOS, as a zip keeps no extended attributes"
  fi
}

start_app() {
  local app="$1" status=0
  if [ "$ARCH" = x64 ] && [ "$(uname -m)" = arm64 ] && ! arch -x86_64 /usr/bin/true 2>/dev/null; then
    echo "::notice::$ARCH: the installed app was not started - this runner has no Rosetta"
    return
  fi
  perl -e 'alarm shift; exec @ARGV or die "exec failed: $!\n"' 120 \
    "$app/Contents/MacOS/$EXE_UNIX" --veloapp-obsolete "$VERSION" || status=$?
  [ $status -eq 0 ] || fail "installer: the app did not start and exit cleanly under the hardened runtime (exit $status)"
  echo "$ARCH installer: the app started and exited cleanly"
}

check_app installer "$INSTALLED"
check_app update "$UPDATED"

for file in Contents/Info.plist "Contents/MacOS/$EXE_UNIX" Contents/_CodeSignature/CodeResources; do
  cmp -s "$INSTALLED/$file" "$UPDATED/$file" || fail "the update package's $file differs from the installer's"
done
echo "$ARCH: the installer and the update package carry the same Info.plist, executable and seal"

deep_check installer "$INSTALLED" required
deep_check update "$UPDATED" reported
start_app "$INSTALLED"
