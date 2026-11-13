// Gera os ícones próprios do ribbon (src/icons/*.png) a partir de SF Symbols, iguais aos do painel.
// Rodar: swift build/icons.swift   (só quando mudar a lista abaixo; os PNG ficam versionados)
import AppKit

let gray = NSColor(srgbRed: 0.70, green: 0.70, blue: 0.72, alpha: 1)  // legível no ribbon claro e no escuro
let icons: [(String, String, NSColor)] = [
    ("stkReview", "tag.fill", .systemBlue),
    ("stkDraft", "tag.fill", .systemGray),
    ("stkWip", "tag.fill", .systemYellow),
    ("stkDone", "tag.fill", .systemGreen),
    ("stkOther", "tag", gray),
    ("stkNext", "arrow.forward.circle", gray),
    ("stkDelete", "trash", gray),
    ("agenda", "list.number", gray),
    ("agendaDel", "text.badge.xmark", gray),
]

let dir = URL(fileURLWithPath: CommandLine.arguments[0]).deletingLastPathComponent()
    .appendingPathComponent("../src/icons").standardized
try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
let px = 32  // botão de 16 pt em tela retina
for (name, symbol, color) in icons {
    let cfg = NSImage.SymbolConfiguration(pointSize: 22, weight: .semibold)
        .applying(NSImage.SymbolConfiguration(paletteColors: [color]))
    guard let img = NSImage(systemSymbolName: symbol, accessibilityDescription: nil)?.withSymbolConfiguration(cfg) else {
        fatalError("símbolo não existe: \(symbol)")
    }
    let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: px, pixelsHigh: px, bitsPerSample: 8,
                               samplesPerPixel: 4, hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB,
                               bytesPerRow: 0, bitsPerPixel: 0)!
    NSGraphicsContext.saveGraphicsState()
    NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
    let s = img.size, k = min(CGFloat(px) / s.width, CGFloat(px) / s.height) * 0.9
    let w = s.width * k, h = s.height * k
    img.draw(in: NSRect(x: (CGFloat(px) - w) / 2, y: (CGFloat(px) - h) / 2, width: w, height: h))
    NSGraphicsContext.restoreGraphicsState()
    try rep.representation(using: .png, properties: [:])!.write(to: dir.appendingPathComponent("\(name).png"))
    print("\(name).png")
}
