#!/usr/bin/env bash
# Send one command to a running scripts/webclient/driver.mjs; the JS for `eval` comes on stdin.
# Usage: drive.sh <port> <token> [eval|quit]    (MAXT: curl timeout in seconds, default 60)
port="$1"
token="$2"
cmd="${3:-eval}"
curl -s --max-time "${MAXT:-60}" -X POST --data-binary @- "http://127.0.0.1:$port/$token/$cmd"
echo
