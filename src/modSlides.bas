Attribute VB_Name = "modSlides"
Option Explicit
Option Private Module  ' internos: fora da lista de Macros, so o ribbon chama

Public Const TAG_STICKER As String = "PRUMO_STICKER"
Public Const TAG_PAGE As String = "PRUMO_PAGE"
Public Const TAG_COMMENT As String = "PRUMO_COMMENT"
Public Const TAG_FOOTNOTE As String = "PRUMO_FOOTNOTE"
Public Const TAG_AGENDA As String = "PRUMO_AGENDA"  ' no slide: numero do item em destaque

' ---------- Stickers de status (um por slide; inserir de novo substitui) ----------

Public Function InsertSticker(sld As Slide, label As String, ByVal back As Long, ByVal fore As Long) As Shape
    Dim shp As Shape, w As Single
    w = 110
    DeleteTagged sld, TAG_STICKER
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, sld.Parent.PageSetup.SlideWidth - w - 10, 10, w, 22)
    shp.Name = "Prumo Sticker"
    shp.Tags.Add TAG_STICKER, label
    shp.Fill.Solid: shp.Fill.ForeColor.RGB = back
    shp.Line.Visible = msoFalse
    With shp.TextFrame.TextRange
        .Text = label
        .Font.Size = 10: .Font.Bold = msoTrue: .Font.Color.RGB = fore
        .ParagraphFormat.Alignment = ppAlignCenter
    End With
    Set InsertSticker = shp
End Function

' Proximo sticker ou comentario (pela tag), a partir do slide atual e voltando ao inicio.
Public Sub FindNextTagged(tag As String, nome As String)
    Dim pres As Presentation, n As Long, start As Long, k As Long, sld As Slide, shp As Shape
    Set pres = ActivePresentation
    n = pres.Slides.Count
    start = CurSlide.SlideIndex
    For k = 1 To n  ' k = n volta ao slide atual
        Set sld = pres.Slides(((start - 1 + k) Mod n) + 1)
        Set shp = FindTagged(sld, tag)
        If Not shp Is Nothing Then
            ActiveWindow.View.GotoSlide sld.SlideIndex
            shp.Select
            Exit Sub
        End If
    Next
    Notify "Nenhum " & nome & " na apresentacao."
End Sub

Public Sub DeleteAllTagged(pres As Presentation, tag As String, nome As String)
    If Not Confirm("Apagar todos os " & nome & "s da apresentacao?") Then Exit Sub
    Notify DeleteTaggedAll(pres, tag) & " " & nome & "(s) apagado(s)."
End Sub

' ---------- Comentarios: caixa amarela com iniciais e data (varios por slide) ----------

Public Function InsertComment(sld As Slide, ByVal author As String) As Shape
    Dim shp As Shape, sw As Single, sh As Single
    sw = sld.Parent.PageSetup.SlideWidth: sh = sld.Parent.PageSetup.SlideHeight
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, (sw - 300) / 2, (sh - 130) / 2, 300, 130)
    shp.Name = "Prumo Comentario"
    shp.Tags.Add TAG_COMMENT, author
    shp.Fill.Solid: shp.Fill.ForeColor.RGB = RGB(255, 255, 153)
    shp.Line.ForeColor.RGB = RGB(89, 89, 89): shp.Line.Weight = 0.75
    With shp.TextFrame
        .VerticalAnchor = msoAnchorTop
        .WordWrap = msoTrue
        .AutoSize = ppAutoSizeNone
        With .TextRange
            .Text = author & " " & Format(Date, "dd/mm/yy") & ": "
            .Font.Size = 14: .Font.Bold = msoFalse: .Font.Color.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
    End With
    Set InsertComment = shp
End Function

' Insere no slide atual e deixa o cursor no fim do texto, pronto para digitar.
Public Sub InsertCommentHere()
    Dim shp As Shape
    Set shp = InsertComment(CurSlide, UserInitials())
    On Error Resume Next
    shp.TextFrame.TextRange.Characters(Len(shp.TextFrame.TextRange.Text) + 1, 0).Select
    If Err.Number <> 0 Then shp.Select  ' se o Mac recusar o cursor no texto, fica a forma selecionada
End Sub

' Iniciais do nome completo do usuario do Mac (via Prumo.scpt): primeira e ultima. Sem o script, usa o login.
Public Function UserInitials() As String
    Dim nome As String, p As Variant, r As String
    On Error Resume Next
    nome = Trim$(AppleScriptTask("Prumo.scpt", "nomeUsuario", ""))
    On Error GoTo 0
    If nome = "" Then UserInitials = UCase$(Left$(Environ("USER"), 2)): Exit Function
    For Each p In Split(nome, " ")
        If Len(p) > 0 Then r = r & UCase$(Left$(p, 1))
    Next
    UserInitials = Left$(r, 1) & Mid$(r, Len(r), IIf(Len(r) > 1, 1, 0))
End Function

' ---------- Nota de rodape (uma por slide; se ja existe, so seleciona) ----------

Public Function InsertFootnote(sld As Slide) As Shape
    Dim shp As Shape, sw As Single, sh As Single
    Set shp = FindTagged(sld, TAG_FOOTNOTE)
    If shp Is Nothing Then
        sw = sld.Parent.PageSetup.SlideWidth: sh = sld.Parent.PageSetup.SlideHeight
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 36, sh - 46, sw - 130, 20)
        shp.Name = "Prumo Nota"
        shp.Tags.Add TAG_FOOTNOTE, "1"
        With shp.TextFrame.TextRange
            .Text = "Fonte: "
            .Font.Size = 9
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
    End If
    Set InsertFootnote = shp
End Function

' ---------- Navegacao e visibilidade ----------

Public Sub GoToSlideN(ByVal n As Long)
    If n < 1 Or n > ActivePresentation.Slides.Count Then Err.Raise vbObjectError + 13, , "a apresentacao tem " & ActivePresentation.Slides.Count & " slides."
    ActiveWindow.View.GotoSlide n
End Sub

' Oculta as formas (some da tela e da apresentacao; "Reexibir" traz de volta).
Public Sub HideShapes(r As ShapeRange)
    Dim s As Shape
    For Each s In r
        s.Visible = msoFalse
    Next
End Sub

' Reexibe todas as formas ocultas do slide. Devolve quantas.
Public Function UnhideShapes(sld As Slide) As Long
    Dim s As Shape
    For Each s In sld.Shapes
        If s.Visible = msoFalse Then s.Visible = msoTrue: UnhideShapes = UnhideShapes + 1
    Next
End Function

Public Sub ToggleMasterShapes(sld As Slide)
    sld.DisplayMasterShapes = IIf(sld.DisplayMasterShapes = msoTrue, msoFalse, msoTrue)
End Sub

' ---------- Ferramentas avancadas ----------

' Nomes das formas do slide parecidas com a 1a: mesmo tipo, mesma forma e mesma cor de preenchimento.
Public Function SimilarNames(ref As Shape) As Variant
    Dim s As Shape, names() As String, n As Long
    ReDim names(0 To 0)
    For Each s In ref.Parent.Shapes
        If IsSimilar(ref, s) Then
            ReDim Preserve names(0 To n)
            names(n) = s.Name: n = n + 1
        End If
    Next
    SimilarNames = names
End Function

Private Function IsSimilar(a As Shape, b As Shape) As Boolean
    On Error Resume Next
    If a.Type <> b.Type Or b.Visible = msoFalse Then Exit Function
    If a.Type = msoAutoShape Then If a.AutoShapeType <> b.AutoShapeType Then Exit Function
    If a.Fill.Visible <> b.Fill.Visible Then Exit Function
    If a.Fill.Visible Then If a.Fill.ForeColor.RGB <> b.Fill.ForeColor.RGB Then Exit Function
    IsSimilar = True
End Function

' Cola os slides selecionados como imagens num slide novo logo depois deles, em grade.
Public Sub InsertSlidesAsPictures(sr As SlideRange)
    Dim pres As Presentation, dst As Slide, i As Long, last As Long, path As String, pic As Shape, cols As Long, cw As Single, src As Object
    Set pres = ActivePresentation
    For i = 1 To sr.Count
        If sr(i).SlideIndex > last Then last = sr(i).SlideIndex
    Next
    Set dst = pres.Slides.Add(last + 1, ppLayoutBlank)
    cols = Int(Sqr(sr.Count - 1)) + 1
    cw = (pres.PageSetup.SlideWidth - 72 - (cols - 1) * GAP_PT) / cols
    For i = 1 To sr.Count
        path = Environ("TMPDIR") & "prumo_slide_" & i & ".png"
        Set src = sr(i)  ' Export por ligacao tardia (pode faltar no Mac)
        src.Export path, "PNG"
        Set pic = dst.Shapes.AddPicture(path, msoFalse, msoTrue, 36, 36)
        pic.LockAspectRatio = msoTrue
        pic.Width = cw
        pic.Line.Visible = msoTrue: pic.Line.ForeColor.RGB = RGB(191, 191, 191): pic.Line.Weight = 0.5
        Kill path
    Next
    AlignInMatrix dst.Shapes.Range, cols
    ActiveWindow.View.GotoSlide dst.SlideIndex
End Sub

' Abre uma copia sem titulo so com os slides selecionados, para salvar onde quiser.
Public Sub SaveSelectedSlides(sr As SlideRange)
    Dim pres As Presentation, cp As Presentation, ids As String, i As Long, path As String
    Set pres = ActivePresentation
    For i = 1 To sr.Count
        ids = ids & "[" & sr(i).SlideID & "]"
    Next
    path = Environ("TMPDIR") & "prumo_selecao.pptx"
    pres.SaveCopyAs path
    Set cp = Presentations.Open(path, msoFalse, msoTrue, msoTrue)  ' somente leitura nao; sem titulo; com janela
    For i = cp.Slides.Count To 1 Step -1
        If InStr(ids, "[" & cp.Slides(i).SlideID & "]") = 0 Then cp.Slides(i).Delete
    Next
    Notify sr.Count & " slide(s) numa apresentacao nova. Salve com Arquivo > Salvar como."
End Sub

' ---------- Numeracao de paginas (caixa propria; a capa fica sem numero) ----------

Public Sub InsertPageNumbers(pres As Presentation)
    Dim sld As Slide, shp As Shape, sw As Single, sh As Single
    sw = pres.PageSetup.SlideWidth: sh = pres.PageSetup.SlideHeight
    For Each sld In pres.Slides
        DeleteTagged sld, TAG_PAGE
        If sld.SlideIndex > 1 Then
            Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, sw - 58, sh - 32, 40, 20)
            shp.Name = "Prumo Pagina"
            shp.Tags.Add TAG_PAGE, "1"
            With shp.TextFrame.TextRange
                .Text = CStr(sld.SlideIndex)
                .Font.Size = 9
                .ParagraphFormat.Alignment = ppAlignRight
            End With
        End If
    Next
End Sub

Public Sub DeletePageNumbers(pres As Presentation)
    DeleteTaggedAll pres, TAG_PAGE
End Sub

' ---------- Agenda: um slide antes de cada secao, com o item da secao em destaque ----------
' Itens = nomes das secoes, menos a 1a (capa e o que vem antes da agenda). Atualizar apaga e refaz.
' Secoes ficam em Object (ligacao tardia): se faltarem no VBA do Mac, vira aviso, nao erro de compilacao.

Public Function AgendaUpdate(pres As Presentation) As Long
    Dim sp As Object, i As Long, n As Long, names() As String, first() As Long, k As Long, sld As Slide
    AgendaDelete pres
    Set sp = pres.SectionProperties
    ReDim names(1 To sp.Count + 1): ReDim first(1 To sp.Count + 1)
    For i = 2 To sp.Count
        If sp.SlidesCount(i) > 0 Then n = n + 1: names(n) = sp.Name(i): first(n) = sp.FirstSlide(i)
    Next
    If n = 0 Then Err.Raise vbObjectError + 16, , "crie secoes nas miniaturas (clique direito > Adicionar secao): cada secao depois da 1a vira um item da agenda."
    For k = n To 1 Step -1  ' de tras para frente: os indices das secoes anteriores nao mudam
        Set sld = pres.Slides.Add(first(k), ppLayoutTitleOnly)
        sld.Tags.Add TAG_AGENDA, CStr(k)
        FillAgenda sld, names, n, k
    Next
    AgendaUpdate = n
End Function

Private Sub FillAgenda(sld As Slide, names() As String, ByVal n As Long, ByVal cur As Long)
    Dim sw As Single, sh As Single, shp As Shape, i As Long, txt As String
    sw = sld.Parent.PageSetup.SlideWidth: sh = sld.Parent.PageSetup.SlideHeight
    If sld.Shapes.HasTitle Then sld.Shapes.Title.TextFrame.TextRange.Text = "Agenda"
    For i = 1 To n
        txt = txt & IIf(i > 1, vbCr, "") & i & "   " & names(i)
    Next
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, sw * 0.08, sh * 0.26, sw * 0.84, sh * 0.6)
    shp.Name = "Prumo Agenda"
    With shp.TextFrame.TextRange
        .Text = txt
        .Font.Size = 20: .Font.Bold = msoFalse: .Font.Color.RGB = RGB(150, 150, 150)
        .ParagraphFormat.SpaceAfter = 10
        With .Paragraphs(cur).Font
            .Bold = msoTrue
            On Error Resume Next
            .Color.RGB = RGB(0, 0, 0)
            .Color.ObjectThemeColor = msoThemeColorAccent1
            On Error GoTo 0
        End With
    End With
End Sub

Public Function AgendaDelete(pres As Presentation) As Long
    Dim i As Long
    For i = pres.Slides.Count To 1 Step -1
        If pres.Slides(i).Tags(TAG_AGENDA) <> "" Then pres.Slides(i).Delete: AgendaDelete = AgendaDelete + 1
    Next
End Function

' ---------- Colar o conteudo da area de transferencia em cada slide selecionado ----------

Public Sub PasteOnSelectedSlides()
    Dim sr As SlideRange, i As Long
    On Error Resume Next
    Set sr = ActiveWindow.Selection.SlideRange
    On Error GoTo 0
    If sr Is Nothing Then Err.Raise vbObjectError + 6, , "copie as formas e selecione os slides no painel de miniaturas."
    For i = 1 To sr.Count
        sr(i).Shapes.Paste
    Next
End Sub

' ---------- Limpeza (sempre com confirmacao; Cmd+Z desfaz) ----------

Public Sub CleanAnimations(pres As Presentation)
    Dim sld As Slide, i As Long, n As Long
    If Not Confirm("Remover TODAS as animacoes da apresentacao?") Then Exit Sub
    For Each sld In pres.Slides
        With sld.TimeLine.MainSequence
            For i = .Count To 1 Step -1
                .Item(i).Delete: n = n + 1
            Next
        End With
    Next
    Notify n & " animacao(oes) removida(s)."
End Sub

Public Sub CleanNotes(pres As Presentation)
    Dim sld As Slide, shp As Shape
    If Not Confirm("Apagar TODAS as anotacoes do orador?") Then Exit Sub
    For Each sld In pres.Slides
        For Each shp In sld.NotesPage.Shapes.Placeholders
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then shp.TextFrame.TextRange.Text = ""
        Next
    Next
End Sub

Public Sub CleanAltTexts(pres As Presentation)
    Dim sld As Slide, shp As Shape
    If Not Confirm("Apagar o texto alternativo de todas as formas?") Then Exit Sub
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            shp.AlternativeText = ""
        Next
    Next
End Sub

' Apaga layouts que nenhum slide usa. Sem Scripting.Dictionary: nao existe no Mac.
Public Sub CleanUnusedLayouts(pres As Presentation)
    Dim sld As Slide, d As Design, used As String, i As Long, n As Long
    If Not Confirm("Apagar os layouts que nenhum slide usa?") Then Exit Sub
    For Each sld In pres.Slides
        used = used & "[" & sld.Design.Index & "|" & sld.CustomLayout.Index & "]"
    Next
    For Each d In pres.Designs
        For i = d.SlideMaster.CustomLayouts.Count To 1 Step -1  ' de tras para frente: indices abaixo nao mudam
            If InStr(used, "[" & d.Index & "|" & i & "]") = 0 Then
                On Error Resume Next
                d.SlideMaster.CustomLayouts(i).Delete
                If Err.Number = 0 Then n = n + 1
                On Error GoTo 0
            End If
        Next
    Next
    Notify n & " layout(s) apagado(s)."
End Sub

' ---------- Auxiliares ----------

Public Function FindTagged(sld As Slide, tag As String) As Shape
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Tags(tag) <> "" Then Set FindTagged = shp: Exit Function
    Next
End Function

Public Function DeleteTagged(sld As Slide, tag As String) As Long
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Tags(tag) <> "" Then sld.Shapes(i).Delete: DeleteTagged = DeleteTagged + 1
    Next
End Function

Public Function DeleteTaggedAll(pres As Presentation, tag As String) As Long
    Dim sld As Slide
    For Each sld In pres.Slides
        DeleteTaggedAll = DeleteTaggedAll + DeleteTagged(sld, tag)
    Next
End Function

