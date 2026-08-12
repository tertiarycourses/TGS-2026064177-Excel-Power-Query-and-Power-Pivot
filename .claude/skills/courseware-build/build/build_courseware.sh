#!/usr/bin/env bash
# Build ALL courseware for Excel Power Query and Power Pivot (TGS-2026064177)
# from the single source (course_data.py + data_domainN.py + lab_steps.py):
#   labs/            19 folders, each with a starter .xlsx (mock data + Power Query)
#   courseware/      PPT, LP.docx, LG.docx  + a PDF of each, with page-numbered TOCs
#   LG-*.md          the Learner Guide Markdown mirror
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="${COURSE_REPO:-$(cd "$HERE/../../../.." && pwd)}"
OUT="$REPO/courseware"

SOFFICE="$(command -v soffice || echo /Applications/LibreOffice.app/Contents/MacOS/soffice)"

echo "==> repo:   $REPO"
echo "==> output: $OUT"

echo "==> 1/5  labs (workbooks + mock data)"
python3 "$HERE/build_labs.py"

echo "==> 2/5  injecting Power Query into the lab workbooks"
python3 "$HERE/inject_powerquery.py"

echo "==> 3/5  slide deck + lab figures"
python3 "$HERE/build_slides.py"
python3 "$HERE/make_lab_figures.py"

echo "==> 4/5  Lesson Plan + Learner Guide (+ Markdown mirror)"
python3 "$HERE/build_lesson_plan.py"
python3 "$HERE/build_learner_guide.py"

echo "==> 5/5  PDFs + page-numbered TOCs"
cd "$OUT"
for f in *.pptx *.docx; do
  [ -e "$f" ] || continue
  "$SOFFICE" --headless --convert-to pdf --outdir . "$f" >/dev/null 2>&1
done
# LibreOffice cannot build a TOC field headless, so inject a static one using the
# page numbers from the first render, then re-render.
for base in LG LP; do
  d=$(ls "$base"-*.docx 2>/dev/null | head -1) || true
  p=$(ls "$base"-*.pdf  2>/dev/null | head -1) || true
  if [ -n "${d:-}" ] && [ -n "${p:-}" ]; then
    python3 "$HERE/inject_toc.py" "$d" "$p" 2
    "$SOFFICE" --headless --convert-to pdf --outdir . "$d" >/dev/null 2>&1
  fi
done

echo
echo "==> done"
ls -la "$OUT"
