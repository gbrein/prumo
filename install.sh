#!/bin/sh
# Instala ou atualiza o Prumo no PowerPoint para Mac.
#   curl -fsSL https://raw.githubusercontent.com/gbrein/prumo/main/install.sh | sh
# Com os arquivos já baixados (Prumo.ppam, libprumo.dylib, Prumo.applescript) numa pasta:
#   sh install.sh /caminho/da/pasta
set -e
URL="https://github.com/gbrein/prumo/releases/latest/download"
DEST="$HOME/Library/Application Support/Microsoft/Prumo"   # a única pasta da casa que o sandbox do PowerPoint lê
SCRIPTS="$HOME/Library/Application Scripts/com.microsoft.Powerpoint"
mkdir -p "$DEST" "$SCRIPTS"

for f in Prumo.ppam libprumo.dylib Prumo.applescript; do
  if [ -n "$1" ]; then cp "$1/$f" "$DEST/$f.new"; else curl -fsSL -o "$DEST/$f.new" "$URL/$f"; fi
  xattr -d com.apple.quarantine "$DEST/$f.new" 2>/dev/null || true  # baixado pelo navegador: o macOS bloquearia o painel
  mv -f "$DEST/$f.new" "$DEST/$f"  # troca atômica: a cópia que o PowerPoint tem aberta continua íntegra
done
osacompile -o "$SCRIPTS/Prumo.scpt" "$DEST/Prumo.applescript"  # dá o nome do usuário para as iniciais dos comentários
echo "Files installed in: $DEST"

echo "Registering the add-in in PowerPoint (macOS may ask to let Terminal control PowerPoint: allow it)..."
if osascript - "$DEST/Prumo.ppam" <<'EOF'
on run {ppam}
  tell application "Microsoft PowerPoint"
    repeat with i from 1 to count add ins
      set a to add in i
      if name of a is "Prumo" and full name of a is not ppam then
        set loaded of a to false
        set auto load of a to false
      end if
    end repeat
    set p to register add in ppam
    set auto load of p to true
    set loaded of p to true
  end tell
end run
EOF
then
  echo "Done. If PowerPoint asks, click Enable Macros. The Prumo tab shows up on the ribbon."
  echo "Updating? Quit and reopen PowerPoint to load the new version of the panel."
else
  echo "Could not register it automatically. In PowerPoint: Tools > PowerPoint Add-ins > + and choose"
  echo "  $DEST/Prumo.ppam   (in the file dialog, press Cmd+Shift+G and paste the path)"
fi
