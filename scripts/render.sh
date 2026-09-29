#!/bin/bash
# usage: render.sh file.pptx outdir  (Korean locale → no CJK/Latin autospace, matches original PDF)
mkdir -p "$2"; LANG=ko_KR.UTF-8 LC_ALL=ko_KR.UTF-8 timeout 180 soffice -env:UserInstallation=file:///tmp/lo2 --headless --convert-to pdf --outdir "$2" "$1" >/dev/null 2>&1
