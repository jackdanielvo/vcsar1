#!/bin/bash
# ============================================================
# VCSAR Team 1 — download History page photos to make them local
# Run this ONCE. On a Mac you can double-click it (see notes below).
# It saves every History/gallery image from vcsar1.org into
# images/gallery/ so the new site no longer depends on WordPress.
# Safe to re-run; it just re-downloads.
# ============================================================

cd "$(dirname "$0")" || exit 1
DEST="images/gallery"
BASE="https://vcsar1.org/wp-content/uploads"
mkdir -p "$DEST"

ok=0; fail=0
dl(){
  if curl -fsSL "$1" -o "$DEST/$2"; then
    echo "  ✓ $2"; ok=$((ok+1))
  else
    echo "  ✗ could not fetch $2"; fail=$((fail+1))
  fi
}

echo "Downloading into $DEST ..."

# Numbered gallery photos (grid thumbnail + full size)
for i in $(seq 1 27); do
  dl "$BASE/2025/01/SARwebsite_History-Gallery-$i-634x476.jpg" "SARwebsite_History-Gallery-$i-634x476.jpg"
  dl "$BASE/2025/01/SARwebsite_History-Gallery-$i.jpg"          "SARwebsite_History-Gallery-$i.jpg"
done

# DSC field photos (thumbnail + full)
for n in DSC_7775 DSC_7708 DSC_7739 DSC_7763 DSC_7753; do
  dl "$BASE/2024/12/$n-634x476.jpg" "$n-634x476.jpg"
  dl "$BASE/2024/12/$n-scaled.jpg"  "$n-scaled.jpg"
done

# Background scans (thumbnail + full)
for n in history-bg01 history-bg02; do
  dl "$BASE/2024/12/$n-634x476.jpg" "$n-634x476.jpg"
  dl "$BASE/2024/12/$n.jpg"         "$n.jpg"
done

# Full-size-only scans
for n in Scan10012_flooding Scan10052_onmountain Scan10042_firsttruck; do
  dl "$BASE/2024/12/$n.jpg" "$n.jpg"
done

# Story / era photos used in the history timeline + page banner
dl "$BASE/2024/12/history_onmountain2.jpg"           "history_onmountain2.jpg"
dl "$BASE/2024/12/history-oldSARbadge-800x800-1.jpg" "history-oldSARbadge-800x800-1.jpg"
dl "$BASE/2024/12/history-Scan10047_crew.jpg"        "history-Scan10047_crew.jpg"
dl "$BASE/2024/12/Scan10042_firsttruck.jpg"          "Scan10042_firsttruck.jpg"
dl "$BASE/2024/12/Scan10052_onmountain.jpg"          "Scan10052_onmountain.jpg"
dl "$BASE/2025/01/SARwebsite_History-jeep-1.jpg"     "SARwebsite_History-jeep-1.jpg"
dl "$BASE/2025/01/SARwebsite_History-hd01-1-2.jpg"   "SARwebsite_History-hd01-1-2.jpg"

echo ""
echo "Finished:  $ok downloaded, $fail failed."
echo "The History page now loads these from $DEST — no WordPress needed."
echo "You can close this window."
