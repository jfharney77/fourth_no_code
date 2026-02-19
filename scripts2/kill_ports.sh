#!/usr/bin/env bash
set -euo pipefail

# Kill any services running on ports 8000-8010

found=0

for port in $(seq 8000 8010); do
    pids=$(lsof -ti :"$port" 2>/dev/null || true)
    if [ -n "$pids" ]; then
        for pid in $pids; do
            echo "Killing PID $pid on port $port"
            kill -9 "$pid"
            found=1
        done
    fi
done

if [ "$found" -eq 0 ]; then
    echo "No services found running on ports 8000-8010"
else
    echo "Done"
fi
