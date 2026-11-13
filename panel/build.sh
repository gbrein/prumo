#!/bin/sh
# Compila o painel (libprumo.dylib, carregada dentro do PowerPoint) e instala onde o sandbox do PowerPoint lê.
# A biblioteca já carregada só é trocada ao reiniciar o PowerPoint.
set -e
cd "$(dirname "$0")/.."
DEST="$HOME/Library/Application Support/Microsoft/Prumo"
mkdir -p dist "$DEST"
# universal (Apple Silicon + Intel), a partir do macOS 12: é este arquivo que vai no release
for a in arm64 x86_64; do
  clang -arch $a -mmacosx-version-min=12.0 -c -o dist/ctor-$a.o panel/ctor.c
  swiftc -O -target $a-apple-macos12 -emit-library -module-name PrumoPainel -o dist/libprumo-$a.dylib panel/PrumoPainel.swift dist/ctor-$a.o
done
lipo -create -output dist/libprumo.dylib dist/libprumo-arm64.dylib dist/libprumo-x86_64.dylib
codesign -s - --force dist/libprumo.dylib
cp dist/libprumo.dylib "$DEST/libprumo.dylib.new"
mv -f "$DEST/libprumo.dylib.new" "$DEST/libprumo.dylib"  # troca atômica: a cópia carregada continua íntegra
cp panel/Prumo.applescript dist/  # o install.sh local (sh install.sh dist) e o release usam esta cópia
osacompile -o ~/Library/Application\ Scripts/com.microsoft.Powerpoint/Prumo.scpt panel/Prumo.applescript
echo "instalado: $DEST/libprumo.dylib"
