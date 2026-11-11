#!/usr/bin/env bash
# ###########################################################################################
# Puts a server build of main on its server version's release in HovKlan-DH/CRT-Server (owner request,
# 2026-11-11: a standing pre-release in CRT's own repository "will polute the list"), for the test server's
# updater, which reads that repository's NEWEST release's pointer every minute (ServerReleaseFeed).
#
#   VERSION=<server version> crt-server-publish-main.sh
#
# The version comes in the environment, never written into the script by the workflow (GitHub's advice against
# script injection - this step holds a token).
#
# Run from the workspace holding the three files the composite action crt-server-build leaves. GH_TOKEN must
# be the secret CRT_SERVER_REPO_TOKEN - a fine-grained token with Contents read and write on CRT-Server ONLY
# (the workflow's own token cannot write another repository).
#
# Safe to run again from the top, which the job does when a run is cancelled or fails half way:
#   - a NEW version's release is made as a DRAFT, its files put on it, and only then published as the newest
#     release - so the box never reads a pointer whose build is not there yet;
#   - an existing version's files are replaced, the build first and its pointer LAST (a newer commit of the same
#     version: the box takes it while that version is not installed - a HELD build is replaced by it - but a
#     version that FAILED to come up is not tried again by itself, whatever its checksum; that needs a new server
#     version, or `update-server --retry-failed` on the box), and its notes rewritten from VERSION.md.
# The tag is made on CRT-Server's own default branch (CRT's commits are not in that repository); the commit
# built is in the pointer, the notes and the attestation.
# ###########################################################################################
set -euo pipefail

REPO="HovKlan-DH/CRT-Server"

if ! printf '%s' "${VERSION:-}" | grep -Eq '^[0-9]+\.[0-9]+\.[0-9]+$'; then
  echo "::error::[${VERSION:-}] is no server version."
  exit 1
fi

if [ -z "${GH_TOKEN:-}" ]; then
  echo "::error::The secret CRT_SERVER_REPO_TOKEN is not set - the test server's build cannot be put on $REPO. INSTALLING.md, \"Automatic updates from GitHub\", says how to make it."
  exit 1
fi

python3 .github/workflows/crt-server-release-notes.py "$VERSION" "$GITHUB_SHA" src/CRT.Server/VERSION.md > crt-server-notes.md
cat crt-server-notes.md

if gh release view "$VERSION" --repo "$REPO" > /dev/null 2>&1; then
  gh release edit "$VERSION" --repo "$REPO" --notes-file crt-server-notes.md
else
  gh release create "$VERSION" --repo "$REPO" --draft --title "CRT.Server $VERSION" --notes-file crt-server-notes.md
fi

gh release upload "$VERSION" crt-server-linux-x64.tar.gz crt-server-linux-x64.tar.gz.sigstore.json --repo "$REPO" --clobber
gh release upload "$VERSION" crt-server-linux-x64.json --repo "$REPO" --clobber

# Published (a draft's first time) and the newest - main's versions only rise.
gh release edit "$VERSION" --repo "$REPO" --draft=false --latest
