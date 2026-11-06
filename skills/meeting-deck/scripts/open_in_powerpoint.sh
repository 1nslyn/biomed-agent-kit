#!/bin/bash
# Opens a copy of a deck in real PowerPoint, exports a PDF and closes only that copy.
# This is the only check that proves PowerPoint accepts the file; LibreOffice and python-pptx accept files it refuses.
# A timeout usually means a Repair prompt is waiting on screen. Close that window by hand.
# Usage: open_in_powerpoint.sh deck.pptx
set -euo pipefail

deck="${1:?usage: open_in_powerpoint.sh deck.pptx}"
[ -f "$deck" ] || { echo "no such file: $deck" >&2; exit 2; }

# PowerPoint is sandboxed: it writes a PDF without a permission dialog only inside its own container.
box="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/meeting_deck_check"
name="deck_check_$$.pptx"
pdf="$box/$(basename "$deck" .pptx).pdf"
mkdir -p "$box"
cp "$deck" "$box/$name"

osascript <<EOF
with timeout of 120 seconds
  tell application "Microsoft PowerPoint"
    open POSIX file "$box/$name"
    delay 3
    set p to presentation "$name"
    set n to count of slides of p
    save p in POSIX file "$pdf" as save as PDF
    close p saving no
    return "PowerPoint opened the deck: " & n & " slides"
  end tell
end timeout
EOF

rm -f "$box/$name"
echo "PDF from PowerPoint: $pdf"
