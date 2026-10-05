#!/bin/bash
set -e
cd "$(dirname "$0")"
./ARRESTA_FINANCE_HUB.command || true
sleep 1
exec ./AVVIA_FINANCE_HUB.command
