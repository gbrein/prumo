// Prumo Painel: barra lateral nativa e offline, carregada DENTRO do PowerPoint (libprumo.dylib, via dlopen no VBA).
// Entra como mais uma coluna no divisor (NSSplitView) que fica abaixo do ribbon, ao lado dos painéis nativos
// do PowerPoint: a área do slide encolhe para dar espaço. Sem esse divisor, vira janela-filha por cima.
// Cada botão roda a macro PrumoP_<id> do suplemento Prumo.ppam (Apple event para o próprio PowerPoint).
import Cocoa

let PPT = "com.microsoft.Powerpoint"
let WIDTH: CGFloat = 232
let RIBBON_GAP: CGFloat = 145  // shortcut: altura fixa do ribbon expandido; medir a área do slide se incomodar
let STATUS_GAP: CGFloat = 28   // barra de status do PowerPoint

// Idioma: o da interface do PowerPoint (a biblioteca roda dentro dele). Português fica em português; outro, inglês.
let isEN = !(Bundle.main.preferredLocalizations.first ?? "pt").hasPrefix("pt")
func L(_ pt: String, _ en: String) -> String { isEN ? en : pt }

struct MenuEntry { let id, label, en: String }
func M(_ id: String, _ label: String, _ en: String) -> MenuEntry { MenuEntry(id: id, label: label, en: en) }

// tip/tipEN: dica em português e em inglês. title: texto no lugar do ícone (EN/PT). menu: opções na setinha ao lado.
// colors: menu com as cores do tema (Fill/Line/Font).
struct Item {
    let id, symbol, tip, tipEN: String
    var tint: NSColor? = nil
    var title: String? = nil
    var menu: [MenuEntry] = []
    var colors: String? = nil
}
struct Section { let title, en: String; let items: [Item]; var colorBar = false }

// Os ids são os Case de Prumo_Run (src/modRibbon.bas); tests/test_build.py confere.
let SECTIONS: [Section] = [
    Section(title: "Apresentação", en: "Presentation", items: [
        Item(id: "PageNumbers", symbol: "number.square", tip: "Numerar páginas (a capa fica sem)", tipEN: "Number pages (the cover gets none)"),
        Item(id: "PageNumbersDelete", symbol: "number.square.fill", tip: "Tirar números de página", tipEN: "Remove page numbers"),
        Item(id: "AgendaUpdate", symbol: "list.number", tip: "Agenda pelas seções: um slide antes de cada uma (apagar na seta)", tipEN: "Agenda from sections: one slide before each (delete in the arrow)", menu: [
            M("AgendaDelete", "Apagar slides de agenda", "Delete agenda slides"),
        ]),
        Item(id: "PasteOnSlides", symbol: "doc.on.clipboard", tip: "Colar nos slides selecionados nas miniaturas", tipEN: "Paste on the slides selected in the thumbnails"),
        Item(id: "SaveSelected", symbol: "square.and.arrow.down.on.square", tip: "Salvar só os slides selecionados (abre uma cópia)", tipEN: "Save only the selected slides (opens a copy)"),
        Item(id: "SlidesAsPictures", symbol: "photo.on.rectangle", tip: "Inserir os slides selecionados como imagens num slide novo", tipEN: "Insert the selected slides as pictures on a new slide"),
        Item(id: "CleanAnimations", symbol: "sparkles", tip: "Limpeza: remover animações (mais na seta)", tipEN: "Cleanup: remove animations (more in the arrow)", menu: [
            M("CleanNotes", "Remover anotações", "Remove notes"), M("CleanAltText", "Remover textos alternativos", "Remove alt texts"),
            M("CleanLayouts", "Remover layouts não usados", "Remove unused layouts"),
        ]),
        Item(id: "BlackWhite", symbol: "circle.lefthalf.filled", tip: "Modo preto e branco (liga/desliga)", tipEN: "Black and white mode (on/off)"),
    ]),
    Section(title: "Posição", en: "Position", items: [
        Item(id: "AlignLeft", symbol: "align.horizontal.left", tip: "Alinhar à esquerda da última", tipEN: "Align to the left of the last"),
        Item(id: "AlignCenter", symbol: "align.horizontal.center", tip: "Centralizar com a última (horizontal)", tipEN: "Center with the last (horizontal)"),
        Item(id: "AlignRight", symbol: "align.horizontal.right", tip: "Alinhar à direita da última", tipEN: "Align to the right of the last"),
        Item(id: "AlignTop", symbol: "align.vertical.top", tip: "Alinhar ao topo da última", tipEN: "Align to the top of the last"),
        Item(id: "AlignMiddle", symbol: "align.vertical.center", tip: "Centralizar com a última (vertical)", tipEN: "Center with the last (vertical)"),
        Item(id: "AlignBottom", symbol: "align.vertical.bottom", tip: "Alinhar à base da última", tipEN: "Align to the bottom of the last"),
        Item(id: "DockLeft", symbol: "arrow.left.to.line", tip: "Encostar à esquerda da última", tipEN: "Dock to the left of the last"),
        Item(id: "DockRight", symbol: "arrow.right.to.line", tip: "Encostar à direita da última", tipEN: "Dock to the right of the last"),
        Item(id: "DockUp", symbol: "arrow.up.to.line", tip: "Encostar acima da última", tipEN: "Dock above the last"),
        Item(id: "DockDown", symbol: "arrow.down.to.line", tip: "Encostar abaixo da última", tipEN: "Dock below the last"),
        Item(id: "StackH", symbol: "rectangle.split.3x1", tip: "Empilhar lado a lado, sem espaço", tipEN: "Stack side by side, no gap"),
        Item(id: "StackV", symbol: "rectangle.split.1x2", tip: "Empilhar uma embaixo da outra, sem espaço", tipEN: "Stack one below the other, no gap"),
        Item(id: "DistH", symbol: "distribute.horizontal.center", tip: "Distribuir na horizontal (espaçamento na seta)", tipEN: "Distribute horizontally (spacing in the arrow)", menu: [
            M("GapHPlus", "Abrir espaço na horizontal (+1 mm)", "Widen horizontal gap (+1 mm)"), M("GapHMinus", "Fechar espaço na horizontal (−1 mm)", "Tighten horizontal gap (−1 mm)"),
        ]),
        Item(id: "DistV", symbol: "distribute.vertical.center", tip: "Distribuir na vertical (espaçamento na seta)", tipEN: "Distribute vertically (spacing in the arrow)", menu: [
            M("GapVPlus", "Abrir espaço na vertical (+1 mm)", "Widen vertical gap (+1 mm)"), M("GapVMinus", "Fechar espaço na vertical (−1 mm)", "Tighten vertical gap (−1 mm)"),
        ]),
        Item(id: "Swap", symbol: "arrow.left.arrow.right", tip: "Trocar posição de 2 formas (variações na seta)", tipEN: "Swap the position of 2 shapes (variants in the arrow)", menu: [
            M("SwapCentered", "Trocar pelos centros", "Swap by centers"), M("SwapH", "Trocar só na horizontal", "Swap horizontally only"), M("SwapV", "Trocar só na vertical", "Swap vertically only"),
            M("SwapHC", "Trocar na horizontal, pelos centros", "Swap horizontally, by centers"), M("SwapVC", "Trocar na vertical, pelos centros", "Swap vertically, by centers"),
        ]),
        Item(id: "DupRight", symbol: "plus.rectangle.on.rectangle", tip: "Duplicar à direita (outras direções na seta)", tipEN: "Duplicate to the right (other directions in the arrow)", menu: [
            M("DupLeft", "Duplicar à esquerda", "Duplicate to the left"), M("DupUp", "Duplicar acima", "Duplicate above"), M("DupDown", "Duplicar abaixo", "Duplicate below"),
        ]),
        Item(id: "CornerTL", symbol: "square.grid.3x3.topleft.filled", tip: "Centralizar no canto superior esquerdo da última (outros cantos na seta)", tipEN: "Center on the top-left corner of the last (other corners in the arrow)", menu: [
            M("CornerTR", "Canto superior direito", "Top-right corner"), M("CornerBL", "Canto inferior esquerdo", "Bottom-left corner"), M("CornerBR", "Canto inferior direito", "Bottom-right corner"),
        ]),
        Item(id: "MemSave", symbol: "pin", tip: "Memorizar posição e tamanho da última", tipEN: "Memorize position and size of the last"),
        Item(id: "MemBoth", symbol: "pin.fill", tip: "Aplicar posição e tamanho memorizados (âncoras na seta)", tipEN: "Apply memorized position and size (anchors in the arrow)", menu: [
            M("MemPos", "Só a posição", "Position only"), M("MemSize", "Só o tamanho", "Size only"), M("MemWidth", "Só a largura", "Width only"), M("MemHeight", "Só a altura", "Height only"),
            M("MemTL", "Âncora: superior esquerda", "Anchor: top left"), M("MemTC", "Âncora: superior centro", "Anchor: top center"), M("MemTR", "Âncora: superior direita", "Anchor: top right"),
            M("MemML", "Âncora: meio esquerda", "Anchor: middle left"), M("MemMC", "Âncora: centro", "Anchor: center"), M("MemMR", "Âncora: meio direita", "Anchor: middle right"),
            M("MemBL", "Âncora: inferior esquerda", "Anchor: bottom left"), M("MemBC", "Âncora: inferior centro", "Anchor: bottom center"), M("MemBR", "Âncora: inferior direita", "Anchor: bottom right"),
        ]),
        Item(id: "AlignInTable", symbol: "tablecells", tip: "Encaixar nas células da tabela (selecione a tabela junto)", tipEN: "Fit into the table cells (select the table too)"),
        Item(id: "Matrix", symbol: "square.grid.3x2", tip: "Organizar em grade (colunas prontas na seta)", tipEN: "Arrange in a grid (preset columns in the arrow)", menu: [
            M("Matrix2", "2 colunas", "2 columns"), M("Matrix3", "3 colunas", "3 columns"), M("Matrix4", "4 colunas", "4 columns"), M("Matrix5", "5 colunas", "5 columns"), M("Matrix6", "6 colunas", "6 columns"),
        ]),
        Item(id: "Circle", symbol: "circle.dashed", tip: "Organizar em círculo (3+ formas)", tipEN: "Arrange in a circle (3+ shapes)"),
        Item(id: "Golden", symbol: "rectangle.ratio.16.to.9", tip: "Proporção áurea (altura = largura ÷ 1,618)", tipEN: "Golden ratio (height = width ÷ 1.618)"),
    ]),
    Section(title: "Tamanho", en: "Size", items: [
        Item(id: "SameWidth", symbol: "arrow.left.and.right.square", tip: "Mesma largura da última", tipEN: "Same width as the last"),
        Item(id: "SameHeight", symbol: "arrow.up.and.down.square", tip: "Mesma altura da última", tipEN: "Same height as the last"),
        Item(id: "SameSize", symbol: "square.on.square", tip: "Mesmo tamanho da última", tipEN: "Same size as the last"),
        Item(id: "WidthPlus", symbol: "plus.rectangle", tip: "Largura +1 mm (mais na seta)", tipEN: "Width +1 mm (more in the arrow)", menu: [
            M("WidthMinus", "Largura −1 mm", "Width −1 mm"), M("HeightPlus", "Altura +1 mm", "Height +1 mm"), M("HeightMinus", "Altura −1 mm", "Height −1 mm"),
        ]),
        Item(id: "StretchLeft", symbol: "arrow.left.to.line.compact", tip: "Esticar até a esquerda da última", tipEN: "Stretch to the left edge of the last"),
        Item(id: "StretchRight", symbol: "arrow.right.to.line.compact", tip: "Esticar até a direita da última", tipEN: "Stretch to the right edge of the last"),
        Item(id: "StretchUp", symbol: "arrow.up.to.line.compact", tip: "Esticar até o topo da última", tipEN: "Stretch to the top of the last"),
        Item(id: "StretchDown", symbol: "arrow.down.to.line.compact", tip: "Esticar até a base da última", tipEN: "Stretch to the bottom of the last"),
        Item(id: "FillLeft", symbol: "arrowshape.left", tip: "Preencher até a vizinha (ou a borda) à esquerda", tipEN: "Fill to the neighbor (or the edge) on the left"),
        Item(id: "FillRight", symbol: "arrowshape.right", tip: "Preencher até a vizinha (ou a borda) à direita", tipEN: "Fill to the neighbor (or the edge) on the right"),
        Item(id: "FillUp", symbol: "arrowshape.up", tip: "Preencher até a vizinha (ou a borda) acima", tipEN: "Fill to the neighbor (or the edge) above"),
        Item(id: "FillDown", symbol: "arrowshape.down", tip: "Preencher até a vizinha (ou a borda) abaixo", tipEN: "Fill to the neighbor (or the edge) below"),
        Item(id: "MagicResizer", symbol: "wand.and.stars", tip: "Redimensionar em % (tamanho, linha e fonte; valores prontos na seta)", tipEN: "Resize by % (size, line and font; presets in the arrow)", menu: [
            M("Resize50", "50%", "50%"), M("Resize75", "75%", "75%"), M("Resize90", "90%", "90%"),
            M("Resize110", "110%", "110%"), M("Resize125", "125%", "125%"), M("Resize150", "150%", "150%"), M("Resize200", "200%", "200%"),
        ]),
    ]),
    Section(title: "Forma", en: "Shape", items: [
        Item(id: "ProcessChain", symbol: "chevron.right.2", tip: "Cadeia de processo: encaixar chevrons/setas lado a lado", tipEN: "Process chain: fit chevrons/arrows side by side"),
        Item(id: "MatchAdjust", symbol: "skew", tip: "Igualar pontas de setas e cantos arredondados à última", tipEN: "Match arrow tips and rounded corners to the last"),
    ]),
    Section(title: "Cor", en: "Color", items: [
        Item(id: "LastFill", symbol: "paintbrush.fill", tip: "Preenchimento: repetir a última cor (cores do tema na seta)", tipEN: "Fill: repeat the last color (theme colors in the arrow)", colors: "Fill"),
        Item(id: "LastLine", symbol: "pencil.line", tip: "Linha: repetir a última cor (cores do tema na seta)", tipEN: "Line: repeat the last color (theme colors in the arrow)", colors: "Line"),
        Item(id: "LastFont", symbol: "textformat", tip: "Texto: repetir a última cor (cores do tema na seta)", tipEN: "Text: repeat the last color (theme colors in the arrow)", colors: "Font"),
        Item(id: "CopyFill", symbol: "eyedropper.full", tip: "Copiar o preenchimento da última", tipEN: "Copy the fill of the last"),
        Item(id: "CopyLine", symbol: "eyedropper.halffull", tip: "Copiar a cor da linha da última", tipEN: "Copy the line color of the last"),
        Item(id: "CopyFont", symbol: "eyedropper", tip: "Copiar a cor do texto da última", tipEN: "Copy the text color of the last"),
        Item(id: "ToRGB", symbol: "number", tip: "Converter cores do tema em RGB (o inverso na seta)", tipEN: "Convert theme colors to RGB (the reverse in the arrow)", menu: [
            M("ToTheme", "Converter RGB em cores do tema", "Convert RGB to theme colors"),
        ]),
    ], colorBar: true),
    Section(title: "Texto", en: "Text", items: [
        Item(id: "LangPT", symbol: "", tip: "Revisão em português (Brasil): seleção ou, sem seleção, a apresentação toda", tipEN: "Proofing in Portuguese (Brazil): selection or, with no selection, the whole presentation", title: "PT"),
        Item(id: "LangEN", symbol: "", tip: "Revisão em inglês (EUA): seleção ou, sem seleção, a apresentação toda", tipEN: "Proofing in English (US): selection or, with no selection, the whole presentation", title: "EN"),
        Item(id: "Spelling", symbol: "textformat.abc", tip: "Verificar ortografia", tipEN: "Check spelling"),
        Item(id: "MarginPlus", symbol: "increase.indent", tip: "Margens internas +1 mm (mais opções na seta)", tipEN: "Inner margins +1 mm (more options in the arrow)", menu: [
            M("MarginMinus", "Margens −1 mm", "Margins −1 mm"), M("MarginZero", "Margens zero", "Zero margins"),
            M("MarginHPlus", "Esquerda e direita +1 mm", "Left and right +1 mm"), M("MarginHMinus", "Esquerda e direita −1 mm", "Left and right −1 mm"),
            M("MarginVPlus", "Topo e base +1 mm", "Top and bottom +1 mm"), M("MarginVMinus", "Topo e base −1 mm", "Top and bottom −1 mm"),
            M("MarginLPlus", "Esquerda +1 mm", "Left +1 mm"), M("MarginLMinus", "Esquerda −1 mm", "Left −1 mm"),
            M("MarginRPlus", "Direita +1 mm", "Right +1 mm"), M("MarginRMinus", "Direita −1 mm", "Right −1 mm"),
            M("MarginTPlus", "Topo +1 mm", "Top +1 mm"), M("MarginTMinus", "Topo −1 mm", "Top −1 mm"),
            M("MarginBPlus", "Base +1 mm", "Bottom +1 mm"), M("MarginBMinus", "Base −1 mm", "Bottom −1 mm"),
        ]),
        Item(id: "FitToText", symbol: "arrow.up.and.down.text.horizontal", tip: "Ajustar a forma ao texto", tipEN: "Fit the shape to the text"),
        Item(id: "WrapText", symbol: "text.append", tip: "Quebrar texto na forma (liga/desliga)", tipEN: "Wrap text in the shape (on/off)"),
        Item(id: "SplitText", symbol: "rectangle.split.1x2", tip: "Dividir a caixa: uma por parágrafo", tipEN: "Split the box: one per paragraph"),
        Item(id: "MergeText", symbol: "rectangle.compress.vertical", tip: "Juntar as caixas de texto numa só", tipEN: "Merge the text boxes into one"),
        Item(id: "Footnote", symbol: "text.below.photo", tip: "Nota de rodapé no slide", tipEN: "Footnote on the slide"),
        Item(id: "ReplaceTexts", symbol: "character.textbox", tip: "Substituir ou apagar o texto das formas selecionadas", tipEN: "Replace or delete the text of the selected shapes"),
    ]),
    Section(title: "Navegação e visibilidade", en: "Navigation and visibility", items: [
        Item(id: "SlidePrev", symbol: "chevron.left", tip: "Slide anterior", tipEN: "Previous slide"),
        Item(id: "GoToSlide", symbol: "arrow.right.to.line.alt", tip: "Ir para o slide…", tipEN: "Go to slide…"),
        Item(id: "SlideNext", symbol: "chevron.right", tip: "Próximo slide", tipEN: "Next slide"),
        Item(id: "HideShapes", symbol: "eye.slash", tip: "Ocultar as formas selecionadas", tipEN: "Hide the selected shapes"),
        Item(id: "UnhideShapes", symbol: "eye", tip: "Reexibir as formas ocultas do slide", tipEN: "Show the hidden shapes on the slide"),
        Item(id: "MasterShapes", symbol: "rectangle.on.rectangle.angled", tip: "Mostrar/esconder as formas do slide mestre", tipEN: "Show/hide the slide master shapes"),
        Item(id: "SelectSimilar", symbol: "square.stack.3d.down.right", tip: "Selecionar formas parecidas com a última", tipEN: "Select shapes similar to the last"),
    ]),
    Section(title: "Revisão", en: "Review", items: [
        Item(id: "CommentInsert", symbol: "note.text.badge.plus", tip: "Comentário com suas iniciais e a data", tipEN: "Comment with your initials and the date", tint: .systemYellow),
        Item(id: "CommentNext", symbol: "arrow.forward.square", tip: "Próximo comentário", tipEN: "Next comment"),
        Item(id: "CommentDeleteAll", symbol: "xmark.bin", tip: "Apagar todos os comentários", tipEN: "Delete all comments"),
        Item(id: "StickerDraft", symbol: "tag.fill", tip: "Sticker: rascunho", tipEN: "Sticker: draft", tint: .systemGray),
        Item(id: "StickerWip", symbol: "tag.fill", tip: "Sticker: em andamento", tipEN: "Sticker: in progress", tint: .systemYellow),
        Item(id: "StickerReview", symbol: "tag.fill", tip: "Sticker: revisar", tipEN: "Sticker: review", tint: .systemBlue),
        Item(id: "StickerDone", symbol: "tag.fill", tip: "Sticker: pronto", tipEN: "Sticker: done", tint: .systemGreen),
        Item(id: "StickerNext", symbol: "arrow.forward.circle", tip: "Próximo sticker", tipEN: "Next sticker"),
        Item(id: "StickerDeleteAll", symbol: "trash", tip: "Apagar todos os stickers", tipEN: "Delete all stickers"),
    ]),
]

// Ações que perguntam algo antes de rodar. O painel pergunta com caixa nativa presa à janela do PowerPoint
// (as do VBA abrem fora da tela com dois monitores) e grava a resposta para a macro (Ask/Confirm em modRibbon).
enum Question { case confirm(String), input(String, String) }
let QUESTIONS: [String: Question] = [
    "CleanAnimations": .confirm(L("Remover TODAS as animações da apresentação?", "Remove ALL animations from the presentation?")),
    "CleanNotes": .confirm(L("Apagar TODAS as anotações do orador?", "Delete ALL speaker notes?")),
    "CleanAltText": .confirm(L("Apagar o texto alternativo de todas as formas?", "Delete the alt text of all shapes?")),
    "CleanLayouts": .confirm(L("Apagar os layouts que nenhum slide usa?", "Delete the layouts no slide uses?")),
    "StickerDeleteAll": .confirm(L("Apagar todos os stickers da apresentação?", "Delete all stickers in the presentation?")),
    "CommentDeleteAll": .confirm(L("Apagar todos os comentários da apresentação?", "Delete all comments in the presentation?")),
    "AgendaDelete": .confirm(L("Apagar os slides de agenda do Prumo?", "Delete the Prumo agenda slides?")),
    "MagicResizer": .input(L("Escala em % (80 reduz, 120 aumenta). Vale para tamanho, linha e fonte.",
                             "Scale in % (80 shrinks, 120 grows). Applies to size, line and font."), "100"),
    "Matrix": .input(L("Quantas colunas?", "How many columns?"), "3"),
    "GoToSlide": .input(L("Ir para o slide:", "Go to slide:"), ""),
    "ReplaceTexts": .input(L("Novo texto para as formas selecionadas (vazio apaga o texto):",
                             "New text for the selected shapes (empty deletes the text):"), ""),
]

// Troca com o VBA na pasta temporária do PowerPoint (a mesma para os dois), em códigos UTF-16 ("80,114,...").
let ARG_FILE = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("prumo_arg.txt")
let MSG_FILE = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("prumo_msg.txt")
let RUN_FILE = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("prumo_run.txt")
func encodeCodes(_ s: String) -> String { s.utf16.map { String($0) }.joined(separator: ",") }
func decodeCodes(_ s: String) -> String {
    let units = s.split(separator: ",").compactMap { Int($0.trimmingCharacters(in: .whitespacesAndNewlines)) }
        .map { UInt16(truncatingIfNeeded: $0) }  // AscW do VBA dá negativo acima de 32767
    return String(utf16CodeUnits: units, count: units.count)
}

final class Panel: NSPanel {
    override var canBecomeKey: Bool { false }  // a janela do slide continua key: a seleção não se perde
}

final class FlippedView: NSView { override var isFlipped: Bool { true } }

final class Prumo: NSObject {
    static let shared = Prumo()
    var panel: Panel!            // reserva: janela-filha por cima, se não achar o divisor
    var pane: NSView!            // o conteúdo do painel (vai para o divisor ou para a janela de reserva)
    var paneWidth: NSLayoutConstraint?
    let status = NSTextField(labelWithString: "")
    var colorBar: NSStackView?
    var itemsById: [String: Item] = [:]
    var visible = true

    func start() {
        buildPanel()
        let nc = NotificationCenter.default
        nc.addObserver(forName: NSWindow.didBecomeMainNotification, object: nil, queue: .main) { [weak self] _ in self?.attach() }
        nc.addObserver(forName: NSWindow.didResizeNotification, object: nil, queue: .main) { [weak self] n in
            guard let self, let w = n.object as? NSWindow, w === self.panel.parent else { return }
            self.layout(in: w)
        }
        // Qualquer caixa do VBA que ganhe foco fora da tela (dois monitores) volta para a janela do documento.
        nc.addObserver(forName: NSWindow.didBecomeKeyNotification, object: nil, queue: .main) { [weak self] _ in
            self?.rescueDialogs()
        }
        nc.addObserver(forName: NSWindow.willCloseNotification, object: nil, queue: .main) { [weak self] n in
            guard let self, let w = n.object as? NSWindow, self.pane.window === w else { return }
            self.pane.removeFromSuperview()  // a janela vai fechar: o painel espera a próxima
        }
        attach()
    }

    func show() {
        guard panel != nil else { return }
        visible = true
        attach()
    }

    @objc func hide() {
        visible = false
        let split = pane.superview as? NSSplitView
        pane.removeFromSuperview()
        split?.adjustSubviews()
        panel.parent?.removeChildWindow(panel)
        panel.orderOut(nil)
    }

    // O divisor de colunas abaixo do ribbon: o NSSplitView vertical mais externo que ocupa a largura da janela.
    func rootSplit(in w: NSWindow) -> NSSplitView? {
        guard let root = w.contentView else { return nil }
        var queue: [NSView] = [root]
        while !queue.isEmpty {
            let v = queue.removeFirst()
            if let s = v as? NSSplitView, s.isVertical, s.frame.width >= root.frame.width - 4 { return s }
            queue.append(contentsOf: v.subviews)
        }
        return nil
    }

    // Coloca o painel como última coluna do divisor, com largura fixa.
    func embed(in split: NSSplitView) {
        if pane.superview === split { return }
        pane.removeFromSuperview()
        panel.parent?.removeChildWindow(panel)
        panel.orderOut(nil)
        pane.translatesAutoresizingMaskIntoConstraints = false
        if split.arrangesAllSubviews { split.addArrangedSubview(pane) } else { split.addSubview(pane) }
        if paneWidth == nil { paneWidth = pane.widthAnchor.constraint(equalToConstant: WIDTH) }
        paneWidth?.isActive = true
        if let i = split.arrangedSubviews.firstIndex(of: pane) {
            split.setHoldingPriority(NSLayoutConstraint.Priority(rawValue: 490), forSubviewAt: i)
        }
        split.adjustSubviews()
        refreshColorBar()
    }

    // Janela de documento do PowerPoint (não o editor VBA, não diálogos, não a apresentação de slides).
    func isDocument(_ w: NSWindow) -> Bool {
        let cls = String(describing: type(of: w))
        return w !== panel && w.isVisible && !(w is NSPanel) && w.frame.width > 600 && w.frame.height > 400
            && !w.title.hasPrefix("Microsoft Visual Basic") && !cls.contains("Show") && !cls.contains("Presenter")
    }

    // A janela do documento onde o painel está (embutido ou como janela-filha), ou a principal.
    var docWindow: NSWindow? {
        if let w = pane.window, w !== panel { return w }
        if let p = panel.parent { return p }
        return [NSApp.mainWindow, NSApp.keyWindow].compactMap { $0 }.first(where: isDocument)
    }

    func attach() {
        guard visible else { return }
        let candidates = [NSApp.mainWindow, NSApp.keyWindow].compactMap { $0 } + NSApp.orderedWindows
        guard let w = candidates.first(where: isDocument) else { return }
        // Janela sem divisor (outra vista) não tira o painel de um documento que continua aberto.
        if rootSplit(in: w) == nil, let host = pane.superview as? NSSplitView, host.window?.isVisible == true { return }
        // Divisor controlado por NSSplitViewController não aceita coluna avulsa: nesse caso, fica a janela de reserva.
        if let split = rootSplit(in: w), !(split.delegate is NSSplitViewController) {
            embed(in: split)
            return
        }
        if pane.superview !== panel.contentView {
            pane.removeFromSuperview()
            paneWidth?.isActive = false
            pane.translatesAutoresizingMaskIntoConstraints = true
            pane.frame = panel.contentView?.bounds ?? .zero
            pane.autoresizingMask = [.width, .height]
            panel.contentView?.addSubview(pane)
        }
        if panel.parent !== w {
            panel.parent?.removeChildWindow(panel)
            w.addChildWindow(panel, ordered: .above)
            refreshColorBar()
        }
        layout(in: w)
    }

    func layout(in w: NSWindow) {
        let r = w.frame
        panel.setFrame(NSRect(x: r.maxX - WIDTH, y: r.minY + STATUS_GAP, width: WIDTH,
                              height: max(240, r.height - RIBBON_GAP - STATUS_GAP)), display: true)
    }

    // Apple event para o próprio PowerPoint.
    func script(_ source: String) -> NSAppleEventDescriptor? {
        var err: NSDictionary?
        let r = NSAppleScript(source: source)?.executeAndReturnError(&err)
        if let e = err {
            status.textColor = .systemRed
            status.stringValue = L("Falhou: ", "Failed: ") + ((e[NSAppleScript.errorMessage] as? String) ?? "\(e)")
            return nil
        }
        return r
    }

    // Clique num botão do painel.
    @objc func run(_ sender: Any) {
        guard let id = (sender as? NSUserInterfaceItemIdentification)?.identifier?.rawValue else { return }
        runId(id)
    }

    // Ação pedida pelo ribbon (o VBA gravou o id em prumo_run.txt).
    func runRequest() {
        guard let raw = try? String(contentsOf: RUN_FILE, encoding: .ascii) else { return }
        try? FileManager.default.removeItem(at: RUN_FILE)
        let id = decodeCodes(raw)
        if !id.isEmpty { runId(id) }
    }

    // Pergunta antes, se a ação pede; depois manda a macro.
    func runId(_ id: String) {
        status.stringValue = ""
        guard let q = QUESTIONS[id] else { send(id); return }
        guard let win = docWindow else { return }
        let alert = NSAlert()
        alert.messageText = "Prumo"
        alert.addButton(withTitle: "OK")
        alert.addButton(withTitle: L("Cancelar", "Cancel"))
        var field: NSTextField?
        switch q {
        case .confirm(let msg):
            alert.informativeText = msg
        case .input(let msg, let deflt):
            alert.informativeText = msg
            let tf = NSTextField(frame: NSRect(x: 0, y: 0, width: 260, height: 24))
            tf.stringValue = deflt
            alert.accessoryView = tf
            alert.window.initialFirstResponder = tf
            field = tf
        }
        alert.beginSheetModal(for: win) { [weak self] resp in
            guard let self, resp == .alertFirstButtonReturn else { return }
            let answer = field?.stringValue ?? "sim"
            if id == "ReplaceTexts", answer.isEmpty {
                self.confirmThen(L("Apagar o texto das formas selecionadas?", "Delete the text of the selected shapes?"), in: win) { self.send(id, arg: "") }
                return
            }
            if field != nil, answer.trimmingCharacters(in: .whitespaces).isEmpty { return }
            self.send(id, arg: answer)
        }
    }

    func confirmThen(_ msg: String, in win: NSWindow, _ action: @escaping () -> Void) {
        let a = NSAlert()
        a.messageText = "Prumo"
        a.informativeText = msg
        a.addButton(withTitle: L("Apagar", "Delete"))
        a.addButton(withTitle: L("Cancelar", "Cancel"))
        a.beginSheetModal(for: win) { if $0 == .alertFirstButtonReturn { action() } }
    }

    // Manda "run VB macro" para o próprio PowerPoint como evento na fila, sem esperar resposta
    // (esperando, a macro rodava dentro da nossa chamada). Depois lê o aviso que a macro deixar.
    func send(_ id: String, arg: String? = nil) {
        try? FileManager.default.removeItem(at: MSG_FILE)
        if let arg { try? encodeCodes(arg).write(to: ARG_FILE, atomically: true, encoding: .ascii) }
        let ev = NSAppleEventDescriptor(eventClass: fourCC("sPPT"), eventID: fourCC("rVbM"),
                                        targetDescriptor: NSAppleEventDescriptor(processIdentifier: getpid()),
                                        returnID: AEReturnID(kAutoGenerateReturnID), transactionID: AETransactionID(kAnyTransactionID))
        ev.setParam(NSAppleEventDescriptor(string: "PrumoP_\(id)"), forKeyword: fourCC("Str#"))
        do {
            try ev.sendEvent(options: [.noReply], timeout: TimeInterval(kAEDefaultTimeout))
        } catch {
            status.stringValue = L("Falhou: ", "Failed: ") + error.localizedDescription
            return
        }
        var ticks = 0
        let poll = Timer(timeInterval: 0.2, repeats: true) { [weak self] t in
            ticks += 1
            self?.rescueDialogs()
            if let raw = try? String(contentsOf: MSG_FILE, encoding: .ascii) {
                try? FileManager.default.removeItem(at: MSG_FILE)
                self?.showMessage(decodeCodes(raw))
                t.invalidate()
            } else if ticks > 75 { t.invalidate() }
        }
        RunLoop.main.add(poll, forMode: .common)
    }

    // Aviso da macro no rodapé: "!" no início = erro (vermelho). Com o painel escondido, vira caixa nativa.
    func showMessage(_ m: String) {
        let isError = m.hasPrefix("!")
        if pane.window == nil || !visible, let win = docWindow {
            let a = NSAlert()
            a.messageText = "Prumo"
            a.informativeText = isError ? String(m.dropFirst()) : m
            a.alertStyle = isError ? .warning : .informational
            a.beginSheetModal(for: win)
            return
        }
        status.textColor = isError ? .systemRed : .secondaryLabelColor
        status.stringValue = isError ? String(m.dropFirst()) : m
    }

    func rescueDialogs() {
        guard let doc = docWindow else { return }
        for w in NSApp.windows where w.isVisible && w !== panel && w !== doc && w.frame.width < 900 && w.frame.height < 700 {
            if !NSScreen.screens.contains(where: { $0.visibleFrame.contains(w.frame) }) {
                let f = doc.frame
                w.setFrameOrigin(NSPoint(x: f.midX - w.frame.width / 2, y: f.midY - w.frame.height / 2))
            }
        }
    }

    // Cores do tema da apresentação ativa (Escuro 1, Claro 1, Escuro 2, Claro 2, Ênfase 1-6).
    func themeColors() -> [NSColor] {
        guard let d = script("tell application id \"\(PPT)\" to get RGB of every theme color of theme color scheme of theme of slide master of active presentation"),
              d.numberOfItems > 0 else { status.stringValue = ""; return [] }
        var n: [Int] = []
        for i in 1...d.numberOfItems {
            guard let item = d.atIndex(i) else { continue }
            if item.numberOfItems > 0 {
                for j in 1...item.numberOfItems { n.append(Int(item.atIndex(j)?.int32Value ?? 0)) }
            } else {
                n.append(Int(item.int32Value))
            }
        }
        var colors: [NSColor] = []
        var i = 0
        while i + 2 < n.count, colors.count < 10 {
            let r = CGFloat(n[i]) / 255, g = CGFloat(n[i + 1]) / 255, b = CGFloat(n[i + 2]) / 255
            colors.append(NSColor(srgbRed: r, green: g, blue: b, alpha: 1))
            i += 3
        }
        return colors
    }

    func swatch(_ c: NSColor, _ side: CGFloat = 14) -> NSImage {
        NSImage(size: NSSize(width: side, height: side), flipped: false) { r in
            c.setFill(); r.fill()
            NSColor.gray.setStroke(); NSBezierPath(rect: r.insetBy(dx: 0.5, dy: 0.5)).stroke()
            return true
        }
    }

    // Barra de cores: um clique pinta o preenchimento com a cor do tema.
    @objc func refreshColorBar() {
        guard let bar = colorBar else { return }
        bar.arrangedSubviews.forEach { $0.removeFromSuperview() }
        for (i, c) in themeColors().enumerated() {
            let b = NSButton(image: swatch(c, 16), target: self, action: #selector(run(_:)))
            b.isBordered = false
            b.identifier = NSUserInterfaceItemIdentifier("FillT\(i + 1)")
            b.toolTip = L("Preencher com a cor \(i + 1) do tema", "Fill with theme color \(i + 1)")
            bar.addArrangedSubview(b)
        }
        let again = NSButton(image: NSImage(systemSymbolName: "arrow.clockwise", accessibilityDescription: L("Atualizar cores", "Refresh colors")) ?? NSImage(),
                             target: self, action: #selector(refreshColorBar))
        again.isBordered = false
        again.toolTip = L("Atualizar as cores (depois de trocar de apresentação)", "Refresh the colors (after switching presentations)")
        bar.addArrangedSubview(again)
    }

    @objc func showMenu(_ sender: NSButton) {
        guard let id = sender.identifier?.rawValue, let it = itemsById[id] else { return }
        let menu = NSMenu()
        func add(_ id: String, _ label: String, _ image: NSImage? = nil) {
            let mi = NSMenuItem(title: label, action: #selector(run(_:)), keyEquivalent: "")
            mi.target = self
            mi.identifier = NSUserInterfaceItemIdentifier(id)
            mi.image = image
            menu.addItem(mi)
        }
        if let kind = it.colors {
            let names = isEN
                ? ["Dark 1", "Light 1", "Dark 2", "Light 2", "Accent 1", "Accent 2", "Accent 3", "Accent 4", "Accent 5", "Accent 6"]
                : ["Escuro 1", "Claro 1", "Escuro 2", "Claro 2", "Ênfase 1", "Ênfase 2", "Ênfase 3", "Ênfase 4", "Ênfase 5", "Ênfase 6"]
            let colors = themeColors()
            for (i, name) in names.enumerated() { add("\(kind)T\(i + 1)", name, i < colors.count ? swatch(colors[i]) : nil) }
            if kind != "Font" {
                menu.addItem(.separator())
                add("\(kind)None", kind == "Fill" ? L("Sem preenchimento", "No fill") : L("Sem linha", "No line"))
            }
        } else {
            it.menu.forEach { add($0.id, L($0.label, $0.en)) }
        }
        menu.popUp(positioning: nil, at: NSPoint(x: 0, y: sender.bounds.height + 2), in: sender)
    }

    func button(_ it: Item) -> NSButton {
        let b = NSButton()
        let tip = L(it.tip, it.tipEN)
        b.target = self
        b.action = #selector(run(_:))
        if let t = it.title {
            b.title = t
            b.font = .boldSystemFont(ofSize: 11)
        } else if let img = NSImage(systemSymbolName: it.symbol, accessibilityDescription: tip)?
            .withSymbolConfiguration(.init(pointSize: 15, weight: .regular)) {
            b.image = img
        } else {
            b.title = String(it.id.prefix(3))  // símbolo ausente nesta versão do macOS
            b.font = .systemFont(ofSize: 9)
        }
        b.identifier = NSUserInterfaceItemIdentifier(it.id)
        b.toolTip = tip
        b.isBordered = false
        b.contentTintColor = it.tint
        b.setAccessibilityLabel(tip)
        b.translatesAutoresizingMaskIntoConstraints = false
        b.widthAnchor.constraint(equalToConstant: 28).isActive = true
        b.heightAnchor.constraint(equalToConstant: 26).isActive = true
        return b
    }

    // Botão simples, ou botão + setinha de menu (como os botões divididos do Office).
    func control(_ it: Item) -> NSView {
        let b = button(it)
        guard !it.menu.isEmpty || it.colors != nil else { return b }
        itemsById[it.id] = it
        let arrow = NSButton()
        arrow.image = NSImage(systemSymbolName: "chevron.down", accessibilityDescription: L("Mais opções", "More options"))?
            .withSymbolConfiguration(.init(pointSize: 8, weight: .semibold))
        arrow.target = self
        arrow.action = #selector(showMenu(_:))
        arrow.identifier = NSUserInterfaceItemIdentifier(it.id)
        arrow.toolTip = it.colors != nil ? L("Cores do tema", "Theme colors") : L("Mais opções", "More options")
        arrow.isBordered = false
        arrow.translatesAutoresizingMaskIntoConstraints = false
        arrow.widthAnchor.constraint(equalToConstant: 12).isActive = true
        arrow.heightAnchor.constraint(equalToConstant: 26).isActive = true
        let pair = NSStackView(views: [b, arrow])
        pair.spacing = 0
        return pair
    }

    func buildPanel() {
        panel = Panel(contentRect: NSRect(x: 0, y: 0, width: WIDTH, height: 600),
                      styleMask: [.borderless, .nonactivatingPanel], backing: .buffered, defer: false)
        panel.isReleasedWhenClosed = false
        panel.hidesOnDeactivate = false
        panel.becomesKeyOnlyIfNeeded = true
        panel.collectionBehavior = [.fullScreenAuxiliary]

        let stack = NSStackView()
        stack.orientation = .vertical
        stack.alignment = .leading
        stack.spacing = 4
        stack.edgeInsets = NSEdgeInsets(top: 8, left: 10, bottom: 8, right: 10)
        stack.translatesAutoresizingMaskIntoConstraints = false

        // Cabeçalho: nome e botão de fechar (o botão Painel do ribbon traz de volta).
        let name = NSTextField(labelWithString: "Prumo")
        name.font = .boldSystemFont(ofSize: 15)
        let close = NSButton(image: NSImage(systemSymbolName: "xmark.circle.fill", accessibilityDescription: L("Fechar", "Close")) ?? NSImage(),
                             target: self, action: #selector(hide))
        close.isBordered = false
        close.toolTip = L("Fechar o painel (o botão Painel do ribbon traz de volta)", "Close the panel (the Panel button on the ribbon brings it back)")
        let spacer = NSView()
        spacer.setContentHuggingPriority(.defaultLow, for: .horizontal)
        let header = NSStackView(views: [name, spacer, close])
        header.widthAnchor.constraint(equalToConstant: WIDTH - 20).isActive = true
        stack.addArrangedSubview(header)

        for s in SECTIONS {
            let sep = NSBox(); sep.boxType = .separator
            stack.addArrangedSubview(sep)
            sep.widthAnchor.constraint(equalToConstant: WIDTH - 20).isActive = true
            let t = NSTextField(labelWithString: L(s.title, s.en))
            t.font = .boldSystemFont(ofSize: 11)
            stack.addArrangedSubview(t)
            // Linhas por largura: botão simples 28 pt, com menu 40 pt.
            var row: [NSView] = [], used: CGFloat = 0
            for it in s.items {
                let w: CGFloat = (it.menu.isEmpty && it.colors == nil) ? 28 : 40
                if used + w > WIDTH - 20, !row.isEmpty {
                    let r = NSStackView(views: row); r.spacing = 2; stack.addArrangedSubview(r)
                    row = []; used = 0
                }
                row.append(control(it)); used += w + 2
            }
            if !row.isEmpty { let r = NSStackView(views: row); r.spacing = 2; stack.addArrangedSubview(r) }
            if s.colorBar {
                let bar = NSStackView()
                bar.spacing = 3
                colorBar = bar
                stack.addArrangedSubview(bar)
            }
        }
        status.font = .systemFont(ofSize: 10)
        status.textColor = .systemRed
        status.lineBreakMode = .byWordWrapping
        status.preferredMaxLayoutWidth = WIDTH - 20
        stack.addArrangedSubview(status)

        let doc = FlippedView()
        doc.translatesAutoresizingMaskIntoConstraints = false
        doc.addSubview(stack)
        let scroll = NSScrollView()
        scroll.hasVerticalScroller = true
        scroll.drawsBackground = false
        scroll.documentView = doc
        NSLayoutConstraint.activate([
            stack.topAnchor.constraint(equalTo: doc.topAnchor),
            stack.leadingAnchor.constraint(equalTo: doc.leadingAnchor),
            stack.trailingAnchor.constraint(equalTo: doc.trailingAnchor),
            stack.bottomAnchor.constraint(equalTo: doc.bottomAnchor),
            doc.widthAnchor.constraint(equalTo: scroll.contentView.widthAnchor),
        ])
        // Molduras com autoresizing (não constraints): o tamanho da janela é quem manda, não o conteúdo.
        let bg = NSVisualEffectView(frame: NSRect(x: 0, y: 0, width: WIDTH, height: 600))
        bg.identifier = NSUserInterfaceItemIdentifier("PrumoPainel")
        bg.material = .sidebar
        bg.state = .active
        scroll.frame = NSRect(x: 1, y: 0, width: WIDTH - 1, height: 600)
        scroll.autoresizingMask = [.width, .height]
        bg.addSubview(scroll)
        let edge = NSBox(frame: NSRect(x: 0, y: 0, width: 1, height: 600))
        edge.boxType = .separator
        edge.autoresizingMask = [.height]
        bg.addSubview(edge)
        pane = bg
        panel.contentView = NSView(frame: bg.frame)
    }
}

func fourCC(_ s: String) -> FourCharCode { s.utf8.reduce(0) { ($0 << 8) | FourCharCode($1) } }

// Chamado pelo construtor em ctor.c quando o VBA faz dlopen da biblioteca.
@_cdecl("prumo_start")
public func prumoStart() {
    DispatchQueue.main.async { Prumo.shared.start() }
}

// Chamado (na thread principal) quando o VBA manda notify_post("com.gbrein.prumo.run").
@_cdecl("prumo_run_request")
public func prumoRunRequest() {
    Prumo.shared.runRequest()
}

// Chamado (na thread principal) quando o VBA manda notify_post("com.gbrein.prumo.show").
@_cdecl("prumo_show")
public func prumoShow() {
    Prumo.shared.show()
}
