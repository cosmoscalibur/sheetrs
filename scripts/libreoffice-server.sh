#!/usr/bin/env bash
# Starts LibreOffice (headless) listening on a socket, as the scripts that make the oracles and the
# fixtures of the tests expect (`sheetrs-io/tests/assets/oracle/make_oracle.py`,
# `sheetrs-edit/tests/assets/replace/make_fixtures.py`): port 2002, its own profile folder.
#
#   scripts/libreoffice-server.sh        # starts it in the background
#   /opt/libreoffice26.8/program/python sheetrs-io/tests/assets/oracle/make_oracle.py
set -euo pipefail
PROFILE="${LO_PROFILE:-/tmp/sheetrs-lo-profile}"
PORT="${LO_PORT:-2002}"
# `[s]ocket` keeps the pattern from matching this very command line.
if pgrep -f "[a]ccept=socket,host=localhost,port=${PORT}" >/dev/null; then
  echo "already listening on ${PORT}"; exit 0
fi
nohup soffice --headless --invisible --norestore "-env:UserInstallation=file://${PROFILE}" \
  "--accept=socket,host=localhost,port=${PORT};urp;" >/tmp/sheetrs-lo.log 2>&1 &
sleep 8
echo "listening on ${PORT} (profile ${PROFILE}, log /tmp/sheetrs-lo.log)"
