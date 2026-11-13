Attribute VB_Name = "modRibbon"
Option Explicit

' Painel lateral: biblioteca nativa carregada dentro do PowerPoint (como janela-filha, fixa na janela).
' O sandbox do PowerPoint so le de ~/Library/Application Support/Microsoft/, por isso ela fica la.
Private Declare PtrSafe Function dlopen Lib "/usr/lib/libSystem.B.dylib" (ByVal path As String, ByVal mode As Long) As LongPtr
Private Declare PtrSafe Function notify_post Lib "/usr/lib/libSystem.B.dylib" (ByVal name As String) As Long

' Verdadeiro enquanto roda uma acao vinda do painel. No Mac com dois monitores, MsgBox/InputBox abrem fora da tela
' e travam o PowerPoint; por isso, vindo do painel, perguntas e confirmacoes chegam prontas (o painel pergunta
' com caixa nativa e grava a resposta) e avisos voltam por arquivo para o rodape do painel.
Private gFromPanel As Boolean
' Verdadeiro quando o painel carregou: o ribbon manda as acoes por ele (caixas nativas e avisos no rodape).
Private gPanelOn As Boolean

' Callback unico do ribbon: o id do botao escolhe a acao. Com o painel carregado, a acao vai por ele.
Public Sub Prumo_Action(control As IRibbonControl)
    If gPanelOn And control.ID <> "Panel" Then
        WriteCodes PanelFile("prumo_run.txt"), control.ID
        notify_post "com.gbrein.prumo.run"
    Else
        Prumo_Do control.ID
    End If
End Sub

' Entrada comum do ribbon e do painel lateral: erro vira aviso, nunca a tela do VBA.
Public Sub Prumo_Do(id As String)
    On Error GoTo Falha
    Prumo_Run id
    Exit Sub
Falha:
    Notify "Prumo: " & Err.Description, True
End Sub

' Entrada dos PrumoP_ (modPanel), chamados pelo painel.
Public Sub Prumo_Panel(id As String)
    gFromPanel = True
    Prumo_Do id
    gFromPanel = False
End Sub

Public Function FromPanel() As Boolean
    FromPanel = gFromPanel
End Function

' Aviso: no painel vai para o rodape (prefixo "!" = erro); no ribbon, MsgBox.
Public Sub Notify(msg As String, Optional isError As Boolean = False)
    If gFromPanel Then
        WriteCodes PanelFile("prumo_msg.txt"), IIf(isError, "!", "") & msg
    Else
        MsgBox msg, IIf(isError, vbExclamation, vbInformation), "Prumo"
    End If
End Sub

' Pergunta: do painel, a resposta ja veio gravada; do ribbon, InputBox.
Public Function Ask(msg As String, deflt As String) As String
    If gFromPanel Then
        Ask = PanelArg()
    Else
        Ask = InputBox(msg, "Prumo", deflt)
    End If
End Function

' Confirmacao: do painel, o painel ja confirmou (grava "sim"); do ribbon, MsgBox Sim/Nao.
Public Function Confirm(msg As String) As Boolean
    If gFromPanel Then
        Confirm = (PanelArg() = "sim")
    Else
        Confirm = (MsgBox(msg, vbYesNo + vbQuestion, "Prumo") = vbYes)
    End If
End Function

' Arquivos de troca com o painel: na pasta temporaria do proprio PowerPoint (a mesma para o VBA e o painel).
Private Function PanelFile(name As String) As String
    PanelFile = Environ("TMPDIR") & name
End Function

Private Function PanelArg() As String
    Dim path As String
    path = PanelFile("prumo_arg.txt")
    PanelArg = ReadCodes(path)
    On Error Resume Next
    Kill path
End Function

' Texto como lista de codigos UTF-16 ("80,114,..."): sem depender da codificacao de arquivo do VBA no Mac.
Private Sub WriteCodes(path As String, s As String)
    Dim f As Integer, i As Long, buf As String
    For i = 1 To Len(s)
        buf = buf & AscW(Mid$(s, i, 1)) & ","
    Next
    f = FreeFile
    Open path For Output As #f
    Print #f, buf;
    Close #f
End Sub

Private Function ReadCodes(path As String) As String
    Dim f As Integer, raw As String, p As Variant
    On Error GoTo Nada
    f = FreeFile
    Open path For Input As #f
    If Not EOF(f) Then Line Input #f, raw
    Close #f
    For Each p In Split(raw, ",")
        If Len(Trim$(p)) > 0 Then ReadCodes = ReadCodes & ChrW(CLng(p))
    Next
    Exit Function
Nada:
    ReadCodes = ""
End Function

' Ao carregar o suplemento, abre o painel lateral (se o app Prumo Painel estiver instalado).
Public Sub Prumo_OnLoad(ribbon As IRibbonUI)
    On Error Resume Next
    OpenPanel
End Sub

' Carrega o painel (na 1a vez) e pede para ele aparecer (o X do painel esconde).
' Com o arquivo "disable" na pasta do painel, nao carrega (chave de seguranca, ver panel/ctor.c).
Private Sub OpenPanel()
    Dim dirPrumo As String
    dirPrumo = "/Users/" & Environ("USER") & "/Library/Application Support/Microsoft/Prumo/"
    If Len(Dir(dirPrumo & "disable")) > 0 Then Exit Sub
    If dlopen(dirPrumo & "libprumo.dylib", 2) = 0 Then
        Err.Raise vbObjectError + 14, , "painel nao instalado: rode panel/build.sh."
    End If
    gPanelOn = True
    notify_post "com.gbrein.prumo.show"
End Sub

' Tambem chamavel por nome (atalhos, AppleScript).
Public Sub Prumo_Run(id As String)
    Select Case id
        Case "Panel": OpenPanel
        ' Alinhar a 1a forma selecionada
        Case "AlignLeft": AlignToRef Sel(2), "L"
        Case "AlignCenter": AlignToRef Sel(2), "C"
        Case "AlignRight": AlignToRef Sel(2), "R"
        Case "AlignTop": AlignToRef Sel(2), "T"
        Case "AlignMiddle": AlignToRef Sel(2), "M"
        Case "AlignBottom": AlignToRef Sel(2), "B"
        ' Tamanho
        Case "SameWidth": MatchSize Sel(2), True, False
        Case "SameHeight": MatchSize Sel(2), False, True
        Case "SameSize": MatchSize Sel(2), True, True
        Case "WidthPlus": Nudge Sel(1), STEP_PT, 0
        Case "WidthMinus": Nudge Sel(1), -STEP_PT, 0
        Case "HeightPlus": Nudge Sel(1), 0, STEP_PT
        Case "HeightMinus": Nudge Sel(1), 0, -STEP_PT
        ' Espaco
        Case "StackH": Stack Sel(2), True
        Case "StackV": Stack Sel(2), False
        Case "DistH": Sel(3).Distribute msoDistributeHorizontally, msoFalse
        Case "DistV": Sel(3).Distribute msoDistributeVertically, msoFalse
        Case "GapHPlus": ChangeGap Sel(2), True, STEP_PT
        Case "GapHMinus": ChangeGap Sel(2), True, -STEP_PT
        Case "GapVPlus": ChangeGap Sel(2), False, STEP_PT
        Case "GapVMinus": ChangeGap Sel(2), False, -STEP_PT
        ' Posicao
        Case "DockLeft": Dock Sel(2), "L"
        Case "DockRight": Dock Sel(2), "R"
        Case "DockUp": Dock Sel(2), "U"
        Case "DockDown": Dock Sel(2), "D"
        Case "StretchLeft": Stretch Sel(2), "L"
        Case "StretchRight": Stretch Sel(2), "R"
        Case "StretchUp": Stretch Sel(2), "U"
        Case "StretchDown": Stretch Sel(2), "D"
        Case "AlignInTable": AlignInTable Sel(2)
        Case "MagicResizer": MagicResizer
        Case "Swap": Swap Sel(2), False
        Case "SwapCentered": Swap Sel(2), True
        Case "DupRight": DuplicateTo(Sel(1), "R").Select
        Case "DupDown": DuplicateTo(Sel(1), "D").Select
        Case "DupLeft": DuplicateTo(Sel(1), "L").Select
        Case "DupUp": DuplicateTo(Sel(1), "U").Select
        Case "SwapH": SwapAxis Sel(2), True, False
        Case "SwapV": SwapAxis Sel(2), False, False
        Case "SwapHC": SwapAxis Sel(2), True, True
        Case "SwapVC": SwapAxis Sel(2), False, True
        Case "FillLeft": FillGap Sel(1), "L"
        Case "FillRight": FillGap Sel(1), "R"
        Case "FillUp": FillGap Sel(1), "U"
        Case "FillDown": FillGap Sel(1), "D"
        Case "CornerTL": CenterOnCorner Sel(2), "TL"
        Case "CornerTR": CenterOnCorner Sel(2), "TR"
        Case "CornerBL": CenterOnCorner Sel(2), "BL"
        Case "CornerBR": CenterOnCorner Sel(2), "BR"
        Case "Golden": GoldenRatio Sel(1)
        Case "Matrix": Matrix
        Case "Matrix2", "Matrix3", "Matrix4", "Matrix5", "Matrix6": AlignInMatrix Sel(2), CLng(Mid$(id, 7))
        Case "Resize50", "Resize75", "Resize90", "Resize110", "Resize125", "Resize150", "Resize200": MagicResize Sel(1), CSng(Mid$(id, 7))
        Case "Circle": ArrangeCircle Sel(3)
        Case "ProcessChain": ProcessChain Sel(2)
        Case "MatchAdjust": MatchAdjust Sel(2)
        ' Memoria de posicao e tamanho
        Case "MemSave": MemSave Sel(1)
        Case "MemPos": MemApply Sel(1), True, False
        Case "MemSize": MemApply Sel(1), False, True
        Case "MemBoth": MemApply Sel(1), True, True
        Case "MemTL", "MemTC", "MemTR", "MemML", "MemMC", "MemMR", "MemBL", "MemBC", "MemBR": MemAnchor Sel(1), Mid$(id, 4)
        Case "MemWidth": MemDim Sel(1), True
        Case "MemHeight": MemDim Sel(1), False
        ' Texto e cor
        Case "MarginPlus": Margins Sel(1), STEP_PT
        Case "MarginMinus": Margins Sel(1), -STEP_PT
        Case "MarginZero": Margins Sel(1), -10000  ' piso 0 zera as margens
        Case "MarginHPlus": Margins Sel(1), STEP_PT, "LR"
        Case "MarginHMinus": Margins Sel(1), -STEP_PT, "LR"
        Case "MarginVPlus": Margins Sel(1), STEP_PT, "TB"
        Case "MarginVMinus": Margins Sel(1), -STEP_PT, "TB"
        Case "MarginLPlus": Margins Sel(1), STEP_PT, "L"
        Case "MarginLMinus": Margins Sel(1), -STEP_PT, "L"
        Case "MarginRPlus": Margins Sel(1), STEP_PT, "R"
        Case "MarginRMinus": Margins Sel(1), -STEP_PT, "R"
        Case "MarginTPlus": Margins Sel(1), STEP_PT, "T"
        Case "MarginTMinus": Margins Sel(1), -STEP_PT, "T"
        Case "MarginBPlus": Margins Sel(1), STEP_PT, "B"
        Case "MarginBMinus": Margins Sel(1), -STEP_PT, "B"
        Case "FitToText": FitToText Sel(1)
        Case "WrapText": ToggleWrap Sel(1)
        Case "LangPT": SetLanguage msoLanguageIDBrazilianPortuguese, "portugues (Brasil)"
        Case "LangEN": SetLanguage msoLanguageIDEnglishUS, "ingles (EUA)"
        Case "Spelling": RunMso "Spelling"
        ' Cores do tema (1-10), nenhuma e repetir a ultima
        Case "FillT1", "FillT2", "FillT3", "FillT4", "FillT5", "FillT6", "FillT7", "FillT8", "FillT9", "FillT10": ApplyColor Sel(1), 0, CLng(Mid$(id, 6)), 0
        Case "LineT1", "LineT2", "LineT3", "LineT4", "LineT5", "LineT6", "LineT7", "LineT8", "LineT9", "LineT10": ApplyColor Sel(1), 1, CLng(Mid$(id, 6)), 0
        Case "FontT1", "FontT2", "FontT3", "FontT4", "FontT5", "FontT6", "FontT7", "FontT8", "FontT9", "FontT10": ApplyColor Sel(1), 2, CLng(Mid$(id, 6)), 0
        Case "FillNone": ApplyColor Sel(1), 0, -1, 0
        Case "LineNone": ApplyColor Sel(1), 1, -1, 0
        Case "LastFill": ApplyLastColor Sel(1), 0
        Case "LastLine": ApplyLastColor Sel(1), 1
        Case "LastFont": ApplyLastColor Sel(1), 2
        Case "ToRGB": ConvertToRGB Sel(1)
        Case "ToTheme": ConvertToTheme Sel(1)
        ' Texto
        Case "SplitText": SplitTextBox Sel(1)
        Case "MergeText": MergeTextBoxes Sel(2)
        Case "Footnote": SelectEnd InsertFootnote(CurSlide)
        Case "ReplaceTexts": ReplaceTextsAsk
        ' Navegacao e visibilidade
        Case "GoToSlide": GoToSlideAsk
        Case "SlidePrev": GoToSlideN CurSlide.SlideIndex - 1
        Case "SlideNext": GoToSlideN CurSlide.SlideIndex + 1
        Case "HideShapes": HideShapes Sel(1)
        Case "UnhideShapes": Notify UnhideShapes(CurSlide) & " forma(s) reexibida(s)."
        Case "MasterShapes": ToggleMasterShapes CurSlide
        ' Avancadas
        Case "SelectSimilar": CurSlide.Shapes.Range(SimilarNames(Sel(1)(1))).Select
        Case "SlidesAsPictures": InsertSlidesAsPictures SelSlides()
        Case "BlackWhite": ToggleBlackWhite
        Case "SaveSelected": SaveSelectedSlides SelSlides()
        Case "CopyFill": CopyColors Sel(2), "fill"
        Case "CopyLine": CopyColors Sel(2), "line"
        Case "CopyFont": CopyColors Sel(2), "font"
        ' Revisao
        Case "StickerDraft": InsertSticker CurSlide, "RASCUNHO", RGB(141, 141, 141), RGB(255, 255, 255)
        Case "StickerWip": InsertSticker CurSlide, "EM ANDAMENTO", RGB(241, 194, 27), RGB(22, 22, 22)
        Case "StickerReview": InsertSticker CurSlide, "REVISAR", RGB(15, 98, 254), RGB(255, 255, 255)
        Case "StickerDone": InsertSticker CurSlide, "PRONTO", RGB(36, 161, 72), RGB(255, 255, 255)
        Case "StickerNext": FindNextTagged TAG_STICKER, "sticker"
        Case "StickerDeleteAll": DeleteAllTagged ActivePresentation, TAG_STICKER, "sticker"
        ' Comentarios (caixa amarela com iniciais e data)
        Case "CommentInsert": InsertCommentHere
        Case "CommentNext": FindNextTagged TAG_COMMENT, "comentario"
        Case "CommentDeleteAll": DeleteAllTagged ActivePresentation, TAG_COMMENT, "comentario"
        ' Slides
        Case "PageNumbers": InsertPageNumbers ActivePresentation
        Case "PageNumbersDelete": DeletePageNumbers ActivePresentation
        Case "AgendaUpdate": Notify "Agenda: " & AgendaUpdate(ActivePresentation) & " item(ns), um slide antes de cada secao."
        Case "AgendaDelete": If Confirm("Apagar os slides de agenda do Prumo?") Then Notify AgendaDelete(ActivePresentation) & " slide(s) de agenda apagado(s)."
        Case "PasteOnSlides": PasteOnSelectedSlides
        Case "CleanAnimations": CleanAnimations ActivePresentation
        Case "CleanNotes": CleanNotes ActivePresentation
        Case "CleanAltText": CleanAltTexts ActivePresentation
        Case "CleanLayouts": CleanUnusedLayouts ActivePresentation
        Case Else: Err.Raise vbObjectError + 2, , "acao desconhecida: " & id
    End Select
End Sub

' Formas selecionadas; a ULTIMA selecionada e a referencia (Sel a coloca em r(1)).
' Idioma de revisao na selecao; sem selecao, na apresentacao inteira.
Private Sub SetLanguage(ByVal lang As Long, nome As String)
    Dim r As ShapeRange, n As Long
    Set r = SelOrNothing()
    If r Is Nothing Then
        n = ApplyLanguageAll(ActivePresentation, lang)
        Notify "Idioma " & nome & " aplicado em " & n & " forma(s) da apresentacao."
    Else
        ApplyLanguage r, lang
    End If
End Sub

'' Recursos que podem faltar no VBA do Mac ficam em Object (ligacao tardia): se faltarem, viram aviso na hora, nao erro de compilacao.
Private Sub RunMso(idMso As String)
    Dim app As Object
    Set app = Application
    On Error GoTo Falta
    app.CommandBars.ExecuteMso idMso
    Exit Sub
Falta:
    Err.Raise vbObjectError + 15, , "este comando nao esta disponivel pelo VBA no Mac: use o menu Revisao > Ortografia."
End Sub

Private Sub ToggleBlackWhite()
    Dim w As Object
    Set w = ActiveWindow
    On Error GoTo Falta
    w.BlackAndWhite = IIf(w.BlackAndWhite = msoTrue, msoFalse, msoTrue)
    Exit Sub
Falta:
    Err.Raise vbObjectError + 15, , "o modo preto e branco nao esta disponivel pelo VBA no Mac."
End Sub

Private Sub Matrix()
    Dim r As ShapeRange, n As Long
    Set r = Sel(2)
    n = AskInt("Quantas colunas?", Int(Sqr(r.Count - 1)) + 1)
    If n > 0 Then AlignInMatrix r, n
End Sub

Private Sub GoToSlideAsk()
    Dim n As Long
    n = AskInt("Ir para o slide:", CurSlide.SlideIndex)
    If n > 0 Then GoToSlideN n
End Sub

Private Sub ReplaceTextsAsk()
    Dim r As ShapeRange, v As String
    Set r = Sel(1)
    v = Ask("Novo texto para as " & r.Count & " forma(s) selecionada(s). Deixe vazio para apagar o texto.", "")
    If v = "" And Not FromPanel() Then  ' do painel, o apagar ja foi confirmado la
        If Not Confirm("Apagar o texto das " & r.Count & " forma(s)?") Then Exit Sub
    End If
    ReplaceTexts r, v
End Sub

' Cursor no fim do texto da forma (ou so a forma selecionada, se o Mac recusar).
Private Sub SelectEnd(shp As Shape)
    On Error Resume Next
    shp.TextFrame.TextRange.Characters(Len(shp.TextFrame.TextRange.Text) + 1, 0).Select
    If Err.Number <> 0 Then shp.Select
End Sub

Public Function SelSlides() As SlideRange
    On Error Resume Next
    Set SelSlides = ActiveWindow.Selection.SlideRange
    On Error GoTo 0
    If SelSlides Is Nothing Then Err.Raise vbObjectError + 6, , "selecione os slides no painel de miniaturas."
End Function

Private Sub MagicResizer()
    Dim r As ShapeRange, pct As Single
    Set r = Sel(1)
    pct = AskPercent()
    If pct > 0 Then MagicResize r, pct
End Sub

Public Function SelOrNothing() As ShapeRange
    On Error Resume Next
    If ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText Then
        Set SelOrNothing = ActiveWindow.Selection.ShapeRange
    End If
End Function

Public Function Sel(minCount As Long) As ShapeRange
    Dim n As Long
    On Error Resume Next
    n = ActiveWindow.Selection.ShapeRange.Count
    On Error GoTo 0
    If n < minCount Then Err.Raise vbObjectError + 1, , "selecione pelo menos " & minCount & " forma(s)."
    Set Sel = LastFirst(ActiveWindow.Selection.ShapeRange)
End Function

' A referencia e a ULTIMA forma clicada: devolve a selecao com ela em 1o (as demais na ordem do clique),
' assim as acoes seguem usando r(1). Formas dentro de grupo ou nomes ambiguos: devolve como veio.
Public Function LastFirst(r As ShapeRange) As ShapeRange
    Dim idx() As Variant, i As Long, n As Long, o As ShapeRange
    Set LastFirst = r
    n = r.Count
    If n < 2 Then Exit Function
    On Error GoTo Fim
    ReDim idx(0 To n - 1)
    idx(0) = r(n).ZOrderPosition
    For i = 1 To n - 1: idx(i) = r(i).ZOrderPosition: Next
    Set o = r(1).Parent.Shapes.Range(idx)
    If o.Count <> n Then Exit Function
    If o(1).Id <> r(n).Id Then Exit Function
    For i = 2 To n
        If o(i).Id <> r(i - 1).Id Then Exit Function
    Next
    Set LastFirst = o
Fim:
End Function

' Slide em foco; sem slide no modo Normal (VBE ou miniaturas ativos), usa o selecionado nas miniaturas.
Public Function CurSlide() As Slide
    On Error Resume Next
    Set CurSlide = ActiveWindow.View.Slide
    If CurSlide Is Nothing Then Set CurSlide = ActiveWindow.Selection.SlideRange(1)
    On Error GoTo 0
    If CurSlide Is Nothing Then Err.Raise vbObjectError + 7, , "nenhum slide em foco: clique num slide no modo Normal."
End Function
