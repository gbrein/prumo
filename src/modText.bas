Attribute VB_Name = "modText"
Option Explicit
Option Private Module  ' internos: fora da lista de Macros, so o ribbon chama

Private mHasLast(0 To 2) As Boolean, mLastTheme(0 To 2) As Long, mLastRGB(0 To 2) As Long

' Soma delta nas margens internas indicadas (L, R, T, B; padrao: as quatro); nunca abaixo de zero.
Public Sub Margins(r As ShapeRange, ByVal delta As Single, Optional sides As String = "LRTB")
    Dim s As Shape
    For Each s In r
        If s.HasTextFrame Then
            With s.TextFrame
                If InStr(sides, "L") > 0 Then .MarginLeft = Floor0(.MarginLeft + delta)
                If InStr(sides, "R") > 0 Then .MarginRight = Floor0(.MarginRight + delta)
                If InStr(sides, "T") > 0 Then .MarginTop = Floor0(.MarginTop + delta)
                If InStr(sides, "B") > 0 Then .MarginBottom = Floor0(.MarginBottom + delta)
            End With
        End If
    Next
End Sub

' Ajusta a forma ao texto uma vez (nao deixa o ajuste automatico ligado).
Public Sub FitToText(r As ShapeRange)
    Dim s As Shape
    For Each s In r
        If s.HasTextFrame Then
            s.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            s.TextFrame2.AutoSize = msoAutoSizeNone
        End If
    Next
End Sub

' Liga ou desliga a quebra de texto, conforme o estado da 1a forma.
Public Sub ToggleWrap(r As ShapeRange)
    Dim s As Shape, v As MsoTriState
    v = msoTrue
    If r(1).HasTextFrame Then If r(1).TextFrame.WordWrap = msoTrue Then v = msoFalse
    For Each s In r
        If s.HasTextFrame Then s.TextFrame.WordWrap = v
    Next
End Sub

' Idioma de revisao nas formas (com grupos e tabelas). Devolve quantas formas com texto mudaram.
Public Function ApplyLanguage(r As ShapeRange, ByVal lang As Long) As Long
    Dim s As Shape
    For Each s In r
        ApplyLanguage = ApplyLanguage + LangShape(s, lang)
    Next
End Function

' Sem selecao: a apresentacao inteira, inclusive as anotacoes.
Public Function ApplyLanguageAll(pres As Presentation, ByVal lang As Long) As Long
    Dim sld As Slide, s As Shape
    For Each sld In pres.Slides
        For Each s In sld.Shapes
            ApplyLanguageAll = ApplyLanguageAll + LangShape(s, lang)
        Next
        For Each s In sld.NotesPage.Shapes
            ApplyLanguageAll = ApplyLanguageAll + LangShape(s, lang)
        Next
    Next
End Function

Private Function LangShape(s As Shape, ByVal lang As Long) As Long
    Dim i As Long, j As Long
    If s.Type = msoGroup Then
        For i = 1 To s.GroupItems.Count
            LangShape = LangShape + LangShape(s.GroupItems(i), lang)
        Next
    ElseIf s.HasTable Then
        For i = 1 To s.Table.Rows.Count
            For j = 1 To s.Table.Columns.Count
                s.Table.Cell(i, j).Shape.TextFrame2.TextRange.LanguageID = lang  ' TextRange.LanguageID nao existe no Mac
            Next
        Next
        LangShape = 1
    ElseIf s.HasTextFrame Then
        If s.TextFrame.HasText Then
            s.TextFrame2.TextRange.LanguageID = lang
            LangShape = 1
        End If
    End If
End Function

' kind: 0 preenchimento, 1 linha, 2 fonte. theme: 1-10 cor do tema (msoThemeColorIndex), 0 usa rgbv, -1 nenhuma.
Public Sub ApplyColor(r As ShapeRange, ByVal kind As Long, ByVal theme As Long, ByVal rgbv As Long)
    Dim s As Shape
    For Each s In r
        Select Case kind
            Case 0
                If theme = -1 Then
                    s.Fill.Visible = msoFalse
                Else
                    s.Fill.Visible = msoTrue: s.Fill.Solid
                    SetColor s.Fill.ForeColor, theme, rgbv
                End If
            Case 1
                If theme = -1 Then
                    s.Line.Visible = msoFalse
                Else
                    s.Line.Visible = msoTrue
                    SetColor s.Line.ForeColor, theme, rgbv
                End If
            Case 2
                If s.HasTextFrame And theme <> -1 Then SetColor s.TextFrame.TextRange.Font.Color, theme, rgbv
        End Select
    Next
    mHasLast(kind) = True: mLastTheme(kind) = theme: mLastRGB(kind) = rgbv
End Sub

' Repete a ultima cor aplicada pelo Prumo (por tipo).
Public Sub ApplyLastColor(r As ShapeRange, ByVal kind As Long)
    If Not mHasLast(kind) Then Err.Raise vbObjectError + 9, , "nenhuma cor usada ainda: escolha uma no menu ao lado."
    ApplyColor r, kind, mLastTheme(kind), mLastRGB(kind)
End Sub

' Troca cores do tema pelo RGB que elas mostram hoje (preenchimento, linha e texto; grupos item a item).
Public Sub ConvertToRGB(r As ShapeRange)
    Dim s As Shape
    For Each s In r
        ColorsToRGB s
    Next
End Sub

Private Sub ColorsToRGB(s As Shape)
    Dim i As Long
    If s.Type = msoGroup Then
        For i = 1 To s.GroupItems.Count
            ColorsToRGB s.GroupItems(i)
        Next
        Exit Sub
    End If
    If s.Fill.Visible = msoTrue Then s.Fill.ForeColor.RGB = s.Fill.ForeColor.RGB
    If s.Line.Visible = msoTrue Then s.Line.ForeColor.RGB = s.Line.ForeColor.RGB
    If s.HasTextFrame Then
        For i = 1 To s.TextFrame.TextRange.Runs.Count
            With s.TextFrame.TextRange.Runs(i).Font.Color
                .RGB = .RGB
            End With
        Next
    End If
End Sub

' Troca RGB pela cor do tema de mesmo valor, quando existe (preenchimento, linha e texto; grupos item a item).
Public Sub ConvertToTheme(r As ShapeRange)
    Dim s As Shape, i As Long, tc(1 To 10) As Long
    For i = 1 To 10
        tc(i) = r(1).Parent.ThemeColorScheme.Colors(i).RGB
    Next
    For Each s In r
        ColorsToTheme s, tc
    Next
End Sub

Private Sub ColorsToTheme(s As Shape, tc() As Long)
    Dim i As Long
    If s.Type = msoGroup Then
        For i = 1 To s.GroupItems.Count
            ColorsToTheme s.GroupItems(i), tc
        Next
        Exit Sub
    End If
    If s.Fill.Visible = msoTrue Then ToTheme s.Fill.ForeColor, tc
    If s.Line.Visible = msoTrue Then ToTheme s.Line.ForeColor, tc
    If s.HasTextFrame Then
        For i = 1 To s.TextFrame.TextRange.Runs.Count
            ToTheme s.TextFrame.TextRange.Runs(i).Font.Color, tc
        Next
    End If
End Sub

Private Sub ToTheme(cf As ColorFormat, tc() As Long)
    Dim i As Long, v As Long
    If cf.ObjectThemeColor > 0 Then Exit Sub
    v = cf.RGB
    For i = 1 To 10
        If tc(i) = v Then cf.ObjectThemeColor = i: Exit Sub
    Next
End Sub

' Divide uma caixa de texto em uma caixa por paragrafo, empilhadas no mesmo lugar.
Public Sub SplitTextBox(r As ShapeRange)
    Dim s As Shape, c As Shape, n As Long, k As Long, i As Long, h As Single, x As Single, y As Single, tr As TextRange
    Set s = r(1)
    If Not s.HasTextFrame Then Err.Raise vbObjectError + 12, , "selecione uma caixa de texto."
    n = s.TextFrame.TextRange.Paragraphs.Count
    If n < 2 Then Err.Raise vbObjectError + 12, , "a caixa tem um paragrafo so."
    x = s.Left: y = s.Top: h = s.Height / n
    For k = n To 1 Step -1  ' copias antes; a original fica com o 1o paragrafo no fim
        If k > 1 Then Set c = s.Duplicate(1) Else Set c = s
        Set tr = c.TextFrame.TextRange
        For i = n To k + 1 Step -1
            tr.Paragraphs(i).Delete
        Next
        For i = 1 To k - 1
            tr.Paragraphs(1).Delete
        Next
        If Right$(tr.Text, 1) = vbCr Then tr.Characters(Len(tr.Text), 1).Delete
        SetBox c, x, y + (k - 1) * h, c.Width, h
    Next
End Sub

' Junta as caixas de texto (de cima para baixo) na mais alta, um paragrafo por caixa; apaga as demais.
Public Sub MergeTextBoxes(r As ShapeRange)
    Dim idx() As Long, i As Long, dst As Shape, src As Shape
    If r.Count < 2 Then Err.Raise vbObjectError + 2, , "selecione pelo menos 2 caixas de texto."
    idx = SortedIdx(r, False)
    Set dst = r(idx(1))
    For i = 2 To r.Count
        Set src = r(idx(i))
        If src.HasTextFrame Then dst.TextFrame.TextRange.InsertAfter vbCr & src.TextFrame.TextRange.Text
    Next
    For i = r.Count To 2 Step -1
        If r(idx(i)).HasTextFrame Then r(idx(i)).Delete
    Next
End Sub

' Substitui (ou apaga) o texto das formas selecionadas, inclusive celulas de tabela.
Public Sub ReplaceTexts(r As ShapeRange, ByVal novo As String)
    Dim s As Shape, i As Long, j As Long
    For Each s In r
        If s.HasTable Then
            For i = 1 To s.Table.Rows.Count
                For j = 1 To s.Table.Columns.Count
                    s.Table.Cell(i, j).Shape.TextFrame.TextRange.Text = novo
                Next
            Next
        ElseIf s.HasTextFrame Then
            s.TextFrame.TextRange.Text = novo
        End If
    Next
End Sub

Private Sub SetColor(cf As ColorFormat, ByVal theme As Long, ByVal rgbv As Long)
    If theme > 0 Then
        cf.ObjectThemeColor = theme
    Else
        cf.RGB = rgbv
    End If
End Sub

' Copia preenchimento, linha ou cor da fonte da 1a forma para as demais.
Public Sub CopyColors(r As ShapeRange, what As String)
    Dim ref As Shape, i As Long
    Set ref = r(1)
    For i = 2 To r.Count
        With r(i)
            Select Case what
                Case "fill"
                    .Fill.Visible = ref.Fill.Visible
                    If ref.Fill.Visible Then .Fill.Solid: CopyColor ref.Fill.ForeColor, .Fill.ForeColor: .Fill.Transparency = ref.Fill.Transparency
                Case "line"
                    .Line.Visible = ref.Line.Visible
                    If ref.Line.Visible Then CopyColor ref.Line.ForeColor, .Line.ForeColor: .Line.Weight = ref.Line.Weight
                Case "font"
                    If ref.HasTextFrame And .HasTextFrame Then CopyColor ref.TextFrame.TextRange.Characters(1, 1).Font.Color, .TextFrame.TextRange.Font.Color
            End Select
        End With
    Next
End Sub

' Mantem a cor do tema quando a origem usa tema; senao copia o RGB.
Private Sub CopyColor(src As ColorFormat, dst As ColorFormat)
    On Error GoTo SoRGB
    If src.ObjectThemeColor > 0 Then
        dst.ObjectThemeColor = src.ObjectThemeColor
        dst.Brightness = src.Brightness
        Exit Sub
    End If
SoRGB:
    dst.RGB = src.RGB
End Sub
