#!/bin/bash
echo "=== Clip Cutter ==="
echo

pip3 install openpyxl -q 2>/dev/null

read -p "Path to MP4 folder (Enter for default 'input' folder): " INPUT_FOLDER

if [ -n "$INPUT_FOLDER" ]; then
    python3 "$(dirname "$0")/clip_cutter.py" --input "$INPUT_FOLDER"
else
    python3 "$(dirname "$0")/clip_cutter.py"
fi
