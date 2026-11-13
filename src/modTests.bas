Attribute VB_Name = "modTests"
Option Explicit

' Autoteste: cria uma apresentacao temporaria, exercita as acoes e fecha sem salvar.
' Rodar em Ferramentas > Macro > Prumo_SelfTest.
Public Sub Prumo_SelfTest()
    Dim pres As Presentation, sld As Slide, a As Shape, b As Shape, r As ShapeRange, d As ShapeRange
    Dim fails As String, n As Long, c As Shape, t As Shape, e As Shape, rh As Single, f As Shape, g As Shape, h As Shape

    On Error GoTo Erro
    Set pres = Presentations.Add(msoTrue)
    Set sld = pres.Slides.Add(1, ppLayoutBlank)
    Set a = sld.Shapes.AddShape(msoShapeRectangle, 100, 100, 100, 50): a.Name = "A"
    Set b = sld.Shapes.AddShape(msoShapeRectangle, 300, 220, 60, 80): b.Name = "B"
    Set r = sld.Shapes.Range(Array("A", "B"))
    Check fails, "ordem da referencia", r(1).Name = "A"
    Set d = LastFirst(r)
    Check fails, "referencia = ultima clicada", d(1).Name = "B" And d(2).Name = "A"

    AlignToRef r, "L": Check fails, "AlignLeft", Near(b.Left, 100)
    AlignToRef r, "R": Check fails, "AlignRight", Near(b.Left + b.Width, 200)
    AlignToRef r, "C": Check fails, "AlignCenter", Near(b.Left + b.Width / 2, 150)
    AlignToRef r, "T": Check fails, "AlignTop", Near(b.Top, 100)
    AlignToRef r, "B": Check fails, "AlignBottom", Near(b.Top + b.Height, 150)
    AlignToRef r, "M": Check fails, "AlignMiddle", Near(b.Top + b.Height / 2, 125)
    Check fails, "referencia parada", Near(a.Left, 100) And Near(a.Top, 100)

    b.LockAspectRatio = msoTrue
    MatchSize r, True, False: Check fails, "SameWidth", Near(b.Width, 100) And Near(b.Height, 80)
    Check fails, "trava de proporcao restaurada", b.LockAspectRatio = msoTrue
    b.LockAspectRatio = msoFalse  ' daqui em diante o teste muda medidas a mao
    MatchSize r, True, True: Check fails, "SameSize", Near(b.Width, 100) And Near(b.Height, 50)
    Nudge r, STEP_PT, 0: Check fails, "WidthPlus", Near(a.Width, 100 + STEP_PT)
    Nudge r, -STEP_PT, 0

    b.Left = 400: b.Top = 100
    Stack r, True: Check fails, "StackH", Near(b.Left, 200)
    ChangeGap r, True, 10: Check fails, "GapHPlus", Near(b.Left, 210)
    b.Left = 100: b.Top = 300
    Stack r, False: Check fails, "StackV", Near(b.Top, 150)

    Dock r, "R": Check fails, "DockRight", Near(b.Left, 200)
    Dock r, "D": Check fails, "DockDown", Near(b.Top, 150)

    b.Left = 150: b.Top = 300: b.Width = 20
    Stretch r, "R": Check fails, "StretchRight", Near(b.Left + b.Width, 200) And Near(b.Left, 150)
    Stretch r, "U": Check fails, "StretchUp", Near(b.Top, 100) And Near(b.Top + b.Height, 350)

    b.Left = 400: b.Top = 300
    Swap r, False: Check fails, "Swap", Near(a.Left, 400) And Near(b.Left, 100) And Near(b.Top, 100)

    n = sld.Shapes.Count
    Set d = DuplicateTo(sld.Shapes.Range(Array("A")), "R")
    Check fails, "DupRight", sld.Shapes.Count = n + 1 And Near(d(1).Left, a.Left + a.Width + GAP_PT)
    d.Delete

    MemSave sld.Shapes.Range(Array("A"))
    MemApply sld.Shapes.Range(Array("B")), True, True
    Check fails, "Memoria", Near(b.Left, a.Left) And Near(b.Width, a.Width)

    a.TextFrame.TextRange.Text = "x"
    Margins sld.Shapes.Range(Array("A")), -10000: Check fails, "MarginZero", Near(a.TextFrame.MarginLeft, 0)
    Margins sld.Shapes.Range(Array("A")), STEP_PT: Check fails, "MarginPlus", Near(a.TextFrame.MarginLeft, STEP_PT)

    a.Fill.ForeColor.RGB = RGB(10, 20, 30)
    CopyColors r, "fill": Check fails, "CopyFill", b.Fill.ForeColor.RGB = RGB(10, 20, 30)

    InsertSticker sld, "RASCUNHO", 0, 0
    InsertSticker sld, "PRONTO", 0, 0
    Check fails, "Sticker substitui", DeleteTagged(sld, TAG_STICKER) = 1

    Set d = sld.Shapes.Range(Array(InsertComment(sld, "GB").Name))
    Check fails, "Comentario", Left$(d(1).TextFrame.TextRange.Text, 3) = "GB " And d(1).Tags(TAG_COMMENT) = "GB"
    Check fails, "Iniciais", Len(UserInitials()) >= 1
    DeleteTagged sld, TAG_COMMENT

    Margins sld.Shapes.Range(Array("A")), STEP_PT, "L"
    Check fails, "Margem por lado", Near(a.TextFrame.MarginLeft, 2 * STEP_PT) And Near(a.TextFrame.MarginRight, STEP_PT)

    Set c = sld.Shapes.AddShape(msoShapeRectangle, 50, 300, 200, 200): c.Name = "C"
    c.TextFrame.TextRange.Text = "x"
    FitToText sld.Shapes.Range(Array("C")): Check fails, "FitToText", c.Height < 100
    c.TextFrame.WordWrap = msoTrue
    ToggleWrap sld.Shapes.Range(Array("C")): Check fails, "WrapText", c.TextFrame.WordWrap = msoFalse
    ApplyLanguage sld.Shapes.Range(Array("C")), msoLanguageIDEnglishUS
    Check fails, "Idioma", c.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS

    ApplyColor sld.Shapes.Range(Array("C")), 0, 5, 0: Check fails, "Cor do tema", c.Fill.ForeColor.ObjectThemeColor = 5
    c.Fill.ForeColor.RGB = RGB(255, 0, 0)
    ApplyLastColor sld.Shapes.Range(Array("C")), 0: Check fails, "Ultima cor", c.Fill.ForeColor.ObjectThemeColor = 5
    ApplyColor sld.Shapes.Range(Array("C")), 1, -1, 0: Check fails, "Sem linha", c.Line.Visible = msoFalse

    Set t = sld.Shapes.AddTable(2, 2, 400, 300, 200, 100): t.Name = "T"
    rh = t.Table.Rows(1).Height
    Set e = sld.Shapes.AddShape(msoShapeOval, 520, 300 + rh + 5, 10, 10): e.Name = "E"
    AlignInTable sld.Shapes.Range(Array("T", "E"))
    Check fails, "AlignInTable", Near(e.Left, 545) And Near(e.Top + 5, 300 + rh + t.Table.Rows(2).Height / 2)

    e.Width = 100: e.Height = 50: e.Line.Weight = 2
    e.TextFrame.TextRange.Text = "y": e.TextFrame.TextRange.Font.Size = 20
    MagicResize sld.Shapes.Range(Array("E")), 50
    Check fails, "MagicResize", Near(e.Width, 50) And Near(e.Height, 25) And Near(e.Line.Weight, 1) And Near(e.TextFrame.TextRange.Font.Size, 10)

    ' --- leva 3: posicao, forma, cor, texto e visibilidade ---
    sld.Shapes.Range(Array("C", "T", "E")).Delete
    a.Left = 100: a.Top = 100: a.Width = 100: a.Height = 50
    b.Left = 300: b.Top = 100: b.Width = 60: b.Height = 50
    SwapAxis r, True, False: Check fails, "SwapH", Near(a.Left, 300) And Near(b.Left, 100) And Near(a.Top, 100)
    SwapAxis r, True, False
    Set d = DuplicateTo(sld.Shapes.Range(Array("B")), "L")
    Check fails, "DupLeft", Near(d(1).Left, 300 - 60 - GAP_PT)
    d.Delete
    b.Left = 300: b.Top = 100
    FillGap sld.Shapes.Range(Array("B")), "L": Check fails, "FillLeft", Near(b.Left, 200) And Near(b.Left + b.Width, 360)
    b.Left = 300: b.Width = 60
    CenterOnCorner r, "BR": Check fails, "CornerBR", Near(b.Left + b.Width / 2, 200) And Near(b.Top + b.Height / 2, 150)
    MemSave sld.Shapes.Range(Array("A"))
    MemAnchor sld.Shapes.Range(Array("B")), "BR": Check fails, "MemBR", Near(b.Left + b.Width, 200) And Near(b.Top + b.Height, 150)
    MemDim sld.Shapes.Range(Array("B")), True: Check fails, "MemWidth", Near(b.Width, 100) And Near(b.Height, 50)
    GoldenRatio sld.Shapes.Range(Array("B")): Check fails, "Golden", Near(b.Height, 100 / 1.618034)

    Set f = sld.Shapes.AddShape(msoShapeRectangle, 10, 400, 20, 20): f.Name = "F"
    Set g = sld.Shapes.AddShape(msoShapeRectangle, 50, 405, 20, 20): g.Name = "G"
    Set h = sld.Shapes.AddShape(msoShapeRectangle, 90, 398, 20, 20): h.Name = "H"
    AlignInMatrix sld.Shapes.Range(Array("H", "F", "G")), 2
    Check fails, "Matrix", Near(f.Left, 10) And Near(g.Left, 30 + GAP_PT) And Near(h.Top, 398 + 20 + GAP_PT) And Near(f.Top, 398)
    ArrangeCircle sld.Shapes.Range(Array("F", "G", "H"))
    Check fails, "Circle", f.Top < g.Top And Near(g.Top, h.Top) And g.Left > h.Left

    f.AutoShapeType = msoShapeChevron: g.AutoShapeType = msoShapeChevron
    f.Width = 100: f.Height = 40: g.Width = 80: g.Height = 30
    ProcessChain sld.Shapes.Range(Array("F", "G"))
    Check fails, "ProcessChain", Near(g.Top, f.Top) And Near(g.Height, 40) And g.Left < f.Left + f.Width
    g.Height = 20: f.Adjustments(1) = 0.5
    MatchAdjust sld.Shapes.Range(Array("F", "G")): Check fails, "MatchAdjust", Near(g.Adjustments(1) * 20, 0.5 * 40)

    h.Fill.ForeColor.ObjectThemeColor = 5
    ConvertToRGB sld.Shapes.Range(Array("H")): Check fails, "ToRGB", h.Fill.ForeColor.ObjectThemeColor <= 0
    ConvertToTheme sld.Shapes.Range(Array("H")): Check fails, "ToTheme", h.Fill.ForeColor.ObjectThemeColor = 5

    h.TextFrame.TextRange.Text = "um" & vbCr & "dois"
    SplitTextBox sld.Shapes.Range(Array("H"))
    Check fails, "SplitText", h.TextFrame.TextRange.Text = "um" And sld.Shapes(sld.Shapes.Count).TextFrame.TextRange.Text = "dois"
    sld.Shapes(sld.Shapes.Count).Name = "H2"
    MergeTextBoxes sld.Shapes.Range(Array("H", "H2"))
    Check fails, "MergeText", h.TextFrame.TextRange.Paragraphs.Count = 2
    ReplaceTexts sld.Shapes.Range(Array("H")), "z": Check fails, "ReplaceTexts", h.TextFrame.TextRange.Text = "z"

    InsertFootnote sld: InsertFootnote sld
    Check fails, "Footnote", DeleteTagged(sld, TAG_FOOTNOTE) = 1
    HideShapes sld.Shapes.Range(Array("F", "G"))
    Check fails, "Hide/Unhide", f.Visible = msoFalse And UnhideShapes(sld) = 2 And g.Visible = msoTrue
    Check fails, "SelectSimilar", UBound(SimilarNames(g)) = 1

    pres.Slides.Add 2, ppLayoutBlank
    InsertPageNumbers pres: InsertPageNumbers pres
    Check fails, "Numeracao sem capa e sem duplicar", FindTagged(pres.Slides(1), TAG_PAGE) Is Nothing _
        And FindTagged(pres.Slides(2), TAG_PAGE).TextFrame.TextRange.Text = "2" And DeleteTaggedAll(pres, TAG_PAGE) = 1

    ' Agenda: capa + 2 secoes (Contexto: slides 2-3, Proposta: 4-5)
    Do While pres.Slides.Count < 5: pres.Slides.Add pres.Slides.Count + 1, ppLayoutBlank: Loop
    pres.SectionProperties.AddBeforeSlide 1, "Capa"
    pres.SectionProperties.AddBeforeSlide 2, "Contexto"
    pres.SectionProperties.AddBeforeSlide 4, "Proposta"
    n = AgendaUpdate(pres): n = AgendaUpdate(pres)
    Check fails, "Agenda antes de cada secao, sem duplicar", n = 2 And pres.Slides.Count = 7 _
        And pres.Slides(2).Tags(TAG_AGENDA) = "1" And pres.Slides(5).Tags(TAG_AGENDA) = "2" _
        And pres.Slides(5).Shapes("Prumo Agenda").TextFrame.TextRange.Paragraphs(2).Font.Bold = msoTrue
    Check fails, "Agenda apagar", AgendaDelete(pres) = 2 And pres.Slides.Count = 5

Fim:
    On Error Resume Next
    pres.Saved = msoTrue
    pres.Close
    On Error GoTo 0
    If fails = "" Then
        MsgBox "Prumo: todos os testes passaram.", vbInformation, "Prumo"
    Else
        MsgBox "Prumo: falhas em " & fails, vbExclamation, "Prumo"
    End If
    Exit Sub
Erro:
    fails = fails & "[erro: " & Err.Description & "] "
    Resume Fim
End Sub

Private Sub Check(fails As String, name As String, ByVal ok As Boolean)
    If Not ok Then fails = fails & "[" & name & "] "
End Sub

Private Function Near(ByVal x As Single, ByVal y As Single) As Boolean
    Near = Abs(x - y) < 0.5
End Function

' Chamado pelo build: "run VB macro" do AppleScript nao passa parametros no Mac.
' O SaveAs do Mac nao tem o formato .ppam: salva .pptm e o build.py converte.
' HOME aponta para o container do sandbox, por isso o caminho usa USER.
Public Sub Prumo_BuildAddIn()
    ActivePresentation.SaveAs "/Users/" & Environ("USER") & "/git/vault/repos/prumo/dist/Prumo.pptm", ppSaveAsOpenXMLPresentationMacroEnabled
End Sub
