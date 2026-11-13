Attribute VB_Name = "modLayout"
Option Explicit
Option Private Module  ' internos: fora da lista de Macros, so o ribbon chama

' Unidades em pontos. 1 mm = 2,835 pt.
Public Const STEP_PT As Single = 2.835
Public Const GAP_PT As Single = 5.67

' shortcut: memoria vive so enquanto o add-in esta carregado; persistir em Tags se precisar sobreviver ao fechar o PowerPoint.
Private memLeft As Single, memTop As Single, memWidth As Single, memHeight As Single, memSet As Boolean

Public Sub AlignToRef(r As ShapeRange, edge As String)
    Dim ref As Shape, i As Long
    Set ref = r(1)
    For i = 2 To r.Count
        With r(i)
            Select Case edge
                Case "L": .Left = ref.Left
                Case "R": .Left = ref.Left + ref.Width - .Width
                Case "C": .Left = ref.Left + (ref.Width - .Width) / 2
                Case "T": .Top = ref.Top
                Case "B": .Top = ref.Top + ref.Height - .Height
                Case "M": .Top = ref.Top + (ref.Height - .Height) / 2
            End Select
        End With
    Next
End Sub

Public Sub MatchSize(r As ShapeRange, matchW As Boolean, matchH As Boolean)
    Dim ref As Shape, s As Shape, i As Long
    Set ref = r(1)
    For i = 2 To r.Count
        Set s = r(i)
        SetBox s, s.Left, s.Top, IIf(matchW, ref.Width, s.Width), IIf(matchH, ref.Height, s.Height)
    Next
End Sub

Public Sub Nudge(r As ShapeRange, ByVal dw As Single, ByVal dh As Single)
    Dim s As Shape
    For Each s In r
        SetBox s, s.Left, s.Top, Floor0(s.Width + dw), Floor0(s.Height + dh)
    Next
End Sub

' Encosta uma forma na outra, sem espaco, na ordem da esquerda para a direita (ou de cima para baixo).
Public Sub Stack(r As ShapeRange, horizontal As Boolean)
    Dim idx() As Long, k As Long, prev As Shape, cur As Shape
    idx = SortedIdx(r, horizontal)
    For k = 2 To r.Count
        Set prev = r(idx(k - 1)): Set cur = r(idx(k))
        If horizontal Then cur.Left = prev.Left + prev.Width Else cur.Top = prev.Top + prev.Height
    Next
End Sub

' Abre ou fecha o espaco entre as formas; a primeira (mais a esquerda ou no topo) fica parada.
Public Sub ChangeGap(r As ShapeRange, horizontal As Boolean, ByVal delta As Single)
    Dim idx() As Long, k As Long
    idx = SortedIdx(r, horizontal)
    For k = 2 To r.Count
        With r(idx(k))
            If horizontal Then .Left = .Left + delta * (k - 1) Else .Top = .Top + delta * (k - 1)
        End With
    Next
End Sub

' Cola as demais formas no lado escolhido da referencia.
Public Sub Dock(r As ShapeRange, side As String)
    Dim ref As Shape, i As Long
    Set ref = r(1)
    For i = 2 To r.Count
        With r(i)
            Select Case side
                Case "L": .Left = ref.Left - .Width
                Case "R": .Left = ref.Left + ref.Width
                Case "U": .Top = ref.Top - .Height
                Case "D": .Top = ref.Top + ref.Height
            End Select
        End With
    Next
End Sub

' Estica a borda escolhida das demais formas ate a mesma borda da referencia.
Public Sub Stretch(r As ShapeRange, side As String)
    Dim ref As Shape, s As Shape, i As Long
    Set ref = r(1)
    For i = 2 To r.Count
        Set s = r(i)
        Select Case side
            Case "L": SetBox s, ref.Left, s.Top, s.Left + s.Width - ref.Left, s.Height
            Case "R": SetBox s, s.Left, s.Top, ref.Left + ref.Width - s.Left, s.Height
            Case "U": SetBox s, s.Left, ref.Top, s.Width, s.Top + s.Height - ref.Top
            Case "D": SetBox s, s.Left, s.Top, s.Width, ref.Top + ref.Height - s.Top
        End Select
    Next
End Sub

Public Sub Swap(r As ShapeRange, centered As Boolean)
    Dim a As Shape, b As Shape, ax As Single, ay As Single, bx As Single, by As Single
    If r.Count <> 2 Then Err.Raise vbObjectError + 3, , "selecione exatamente 2 formas."
    Set a = r(1): Set b = r(2)
    If centered Then
        ax = a.Left + a.Width / 2: ay = a.Top + a.Height / 2
        bx = b.Left + b.Width / 2: by = b.Top + b.Height / 2
        a.Left = bx - a.Width / 2: a.Top = by - a.Height / 2
        b.Left = ax - b.Width / 2: b.Top = ay - b.Height / 2
    Else
        ax = a.Left: ay = a.Top
        a.Left = b.Left: a.Top = b.Top
        b.Left = ax: b.Top = ay
    End If
End Sub

' Troca so a posicao horizontal (ou vertical) de 2 formas, pela borda ou pelo centro.
Public Sub SwapAxis(r As ShapeRange, horizontal As Boolean, centered As Boolean)
    Dim a As Shape, b As Shape, pa As Single, pb As Single, ka As Single, kb As Single
    If r.Count <> 2 Then Err.Raise vbObjectError + 3, , "selecione exatamente 2 formas."
    Set a = r(1): Set b = r(2)
    If horizontal Then
        If centered Then ka = a.Width / 2: kb = b.Width / 2
        pa = a.Left + ka: pb = b.Left + kb
        a.Left = pb - ka: b.Left = pa - kb
    Else
        If centered Then ka = a.Height / 2: kb = b.Height / 2
        pa = a.Top + ka: pb = b.Top + kb
        a.Top = pb - ka: b.Top = pa - kb
    End If
End Sub

' Duplica a selecao para o lado side (R, L, D, U), com GAP_PT de folga. Devolve as copias para encadear cliques.
Public Function DuplicateTo(r As ShapeRange, side As String) As ShapeRange
    Dim d As ShapeRange, i As Long, lo As Single, hi As Single, a As Single, b As Single, horizontal As Boolean, dirSign As Single  ' Sgn e Dir sao funcoes do VBA: nao servem de nome
    horizontal = (side = "R" Or side = "L")
    dirSign = IIf(side = "L" Or side = "U", -1, 1)
    lo = 1E+30: hi = -1E+30
    For i = 1 To r.Count
        If horizontal Then a = r(i).Left: b = a + r(i).Width Else a = r(i).Top: b = a + r(i).Height
        If a < lo Then lo = a
        If b > hi Then hi = b
    Next
    Set d = r.Duplicate
    For i = 1 To d.Count
        d(i).Left = r(i).Left + IIf(horizontal, dirSign * (hi - lo + GAP_PT), 0)
        d(i).Top = r(i).Top + IIf(horizontal, 0, dirSign * (hi - lo + GAP_PT))
    Next
    Set DuplicateTo = d
End Function

Public Sub MemSave(r As ShapeRange)
    memLeft = r(1).Left: memTop = r(1).Top: memWidth = r(1).Width: memHeight = r(1).Height
    memSet = True
End Sub

Public Sub MemApply(r As ShapeRange, applyPos As Boolean, applySize As Boolean)
    Dim s As Shape
    If Not memSet Then Err.Raise vbObjectError + 4, , "nada memorizado: use 'Memorizar' antes."
    For Each s In r
        SetBox s, IIf(applyPos, memLeft, s.Left), IIf(applyPos, memTop, s.Top), _
                  IIf(applySize, memWidth, s.Width), IIf(applySize, memHeight, s.Height)
    Next
End Sub

' Aplica a posicao memorizada por uma ancora: 1a letra T/M/B (vertical), 2a L/C/R (horizontal).
' O ponto da ancora da forma vai para o mesmo ponto da forma memorizada.
Public Sub MemAnchor(r As ShapeRange, anchor As String)
    Dim s As Shape, fx As Single, fy As Single
    If Not memSet Then Err.Raise vbObjectError + 4, , "nada memorizado: use 'Memorizar' antes."
    fy = (InStr("TMB", Left$(anchor, 1)) - 1) / 2
    fx = (InStr("LCR", Right$(anchor, 1)) - 1) / 2
    For Each s In r
        s.Left = memLeft + fx * memWidth - fx * s.Width
        s.Top = memTop + fy * memHeight - fy * s.Height
    Next
End Sub

' So a largura (ou so a altura) memorizada.
Public Sub MemDim(r As ShapeRange, width As Boolean)
    Dim s As Shape
    If Not memSet Then Err.Raise vbObjectError + 4, , "nada memorizado: use 'Memorizar' antes."
    For Each s In r
        SetBox s, s.Left, s.Top, IIf(width, memWidth, s.Width), IIf(width, s.Height, memHeight)
    Next
End Sub

' Estica cada forma ate a forma vizinha mais proxima no lado pedido (L, R, U, D) ou, sem vizinha, ate a borda do slide.
Public Sub FillGap(r As ShapeRange, side As String)
    Dim s As Shape, o As Shape, lim As Single, sw As Single, sh As Single
    sw = ActivePresentation.PageSetup.SlideWidth: sh = ActivePresentation.PageSetup.SlideHeight
    For Each s In r
        Select Case side
            Case "L", "U": lim = 0
            Case "R": lim = sw
            Case "D": lim = sh
        End Select
        For Each o In s.Parent.Shapes
            If o.Name <> s.Name And o.Visible Then
                Select Case side
                    Case "L": If Overlaps(o.Top, o.Height, s.Top, s.Height) And o.Left + o.Width <= s.Left + 0.5 And o.Left + o.Width > lim Then lim = o.Left + o.Width
                    Case "R": If Overlaps(o.Top, o.Height, s.Top, s.Height) And o.Left >= s.Left + s.Width - 0.5 And o.Left < lim Then lim = o.Left
                    Case "U": If Overlaps(o.Left, o.Width, s.Left, s.Width) And o.Top + o.Height <= s.Top + 0.5 And o.Top + o.Height > lim Then lim = o.Top + o.Height
                    Case "D": If Overlaps(o.Left, o.Width, s.Left, s.Width) And o.Top >= s.Top + s.Height - 0.5 And o.Top < lim Then lim = o.Top
                End Select
            End If
        Next
        Select Case side
            Case "L": SetBox s, lim, s.Top, s.Left + s.Width - lim, s.Height
            Case "R": SetBox s, s.Left, s.Top, lim - s.Left, s.Height
            Case "U": SetBox s, s.Left, lim, s.Width, s.Top + s.Height - lim
            Case "D": SetBox s, s.Left, s.Top, s.Width, lim - s.Top
        End Select
    Next
End Sub

Private Function Overlaps(ByVal a As Single, ByVal al As Single, ByVal b As Single, ByVal bl As Single) As Boolean
    Overlaps = (a < b + bl - 0.5) And (b < a + al - 0.5)
End Function

' Centraliza as demais formas num canto da 1a (TL, TR, BL, BR).
Public Sub CenterOnCorner(r As ShapeRange, corner As String)
    Dim i As Long, x As Single, y As Single
    x = r(1).Left + IIf(Right$(corner, 1) = "R", r(1).Width, 0)
    y = r(1).Top + IIf(Left$(corner, 1) = "B", r(1).Height, 0)
    For i = 2 To r.Count
        r(i).Left = x - r(i).Width / 2
        r(i).Top = y - r(i).Height / 2
    Next
End Sub

' Altura = largura / phi (retangulo aureo), mantendo o canto superior esquerdo.
Public Sub GoldenRatio(r As ShapeRange)
    Dim s As Shape
    For Each s In r
        SetBox s, s.Left, s.Top, s.Width, s.Width / 1.618034
    Next
End Sub

' Organiza em grade de cols colunas, na ordem de leitura, com celula do tamanho da maior forma e GAP_PT entre celulas.
Public Sub AlignInMatrix(r As ShapeRange, ByVal cols As Long)
    Dim idx() As Long, i As Long, k As Long, cw As Single, ch As Single, x0 As Single, y0 As Single, s As Shape
    If cols < 1 Then cols = 1
    x0 = 1E+30: y0 = 1E+30
    For Each s In r
        If s.Width > cw Then cw = s.Width
        If s.Height > ch Then ch = s.Height
        If s.Left < x0 Then x0 = s.Left
        If s.Top < y0 Then y0 = s.Top
    Next
    idx = ReadingOrder(r)
    For k = 0 To r.Count - 1
        Set s = r(idx(k + 1))
        s.Left = x0 + (k Mod cols) * (cw + GAP_PT) + (cw - s.Width) / 2
        s.Top = y0 + (k \ cols) * (ch + GAP_PT) + (ch - s.Height) / 2
    Next
End Sub

' Ordem de leitura: de cima para baixo por linhas (nova linha quando o topo passa de meia altura media),
' e da esquerda para a direita dentro da linha.
Private Function ReadingOrder(r As ShapeRange) As Long()
    Dim idx() As Long, row() As Long, i As Long, j As Long, t As Long, tol As Single, rowTop As Single, n As Long, rowNo As Long
    n = r.Count
    ReDim idx(1 To n): ReDim row(1 To n)
    For i = 1 To n
        idx(i) = i
        tol = tol + r(i).Height / n
    Next
    tol = tol / 2
    If tol < 1 Then tol = 1
    For i = 1 To n - 1  ' por topo
        For j = 1 To n - i
            If r(idx(j)).Top > r(idx(j + 1)).Top Then t = idx(j): idx(j) = idx(j + 1): idx(j + 1) = t
        Next
    Next
    rowTop = r(idx(1)).Top
    For i = 1 To n
        If r(idx(i)).Top > rowTop + tol Then rowTop = r(idx(i)).Top: rowNo = rowNo + 1
        row(idx(i)) = rowNo
    Next
    For i = 1 To n - 1  ' por linha e, dentro dela, pela esquerda
        For j = 1 To n - i
            If row(idx(j)) > row(idx(j + 1)) Or (row(idx(j)) = row(idx(j + 1)) And r(idx(j)).Left > r(idx(j + 1)).Left) Then
                t = idx(j): idx(j) = idx(j + 1): idx(j + 1) = t
            End If
        Next
    Next
    ReadingOrder = idx
End Function

' Distribui as formas num circulo, na ordem da selecao, a partir do topo e no sentido horario.
Public Sub ArrangeCircle(r As ShapeRange)
    Dim s As Shape, k As Long, x0 As Single, y0 As Single, x1 As Single, y1 As Single, big As Single, rad As Single, a As Double
    If r.Count < 3 Then Err.Raise vbObjectError + 2, , "selecione pelo menos 3 formas."
    x0 = 1E+30: y0 = 1E+30: x1 = -1E+30: y1 = -1E+30
    For Each s In r
        If s.Left < x0 Then x0 = s.Left
        If s.Top < y0 Then y0 = s.Top
        If s.Left + s.Width > x1 Then x1 = s.Left + s.Width
        If s.Top + s.Height > y1 Then y1 = s.Top + s.Height
        If s.Width > big Then big = s.Width
        If s.Height > big Then big = s.Height
    Next
    rad = ((x1 - x0) + (y1 - y0)) / 4 - big / 2
    If rad < big Then rad = big
    For k = 1 To r.Count
        a = -1.5707963 + 6.2831853 * (k - 1) / r.Count
        r(k).Left = (x0 + x1) / 2 + rad * Cos(a) - r(k).Width / 2
        r(k).Top = (y0 + y1) / 2 + rad * Sin(a) - r(k).Height / 2
    Next
End Sub

' Cadeia de processo: da esquerda para a direita, mesma altura e topo da 1a, cada forma encaixada na ponta da anterior.
Public Sub ProcessChain(r As ShapeRange)
    Dim idx() As Long, i As Long, a As Shape, b As Shape
    idx = SortedIdx(r, True)
    Set a = r(idx(1))
    For i = 2 To r.Count
        Set b = r(idx(i))
        SetBox b, a.Left + a.Width - PointDepth(a) + STEP_PT, r(idx(1)).Top, b.Width, r(idx(1)).Height
        Set a = b
    Next
End Sub

' Profundidade da ponta de seta/chevron/pentagono (0 nas demais formas).
Private Function PointDepth(s As Shape) As Single
    On Error Resume Next
    If s.Type = msoAutoShape Then
        Select Case s.AutoShapeType
            Case msoShapeChevron, msoShapePentagon: PointDepth = s.Adjustments(1) * IIf(s.Width < s.Height, s.Width, s.Height)
        End Select
    End If
End Function

' Iguala pontas de setas e cantos arredondados a 1a: o 1o ajuste em tamanho absoluto, os demais copiados.
Public Sub MatchAdjust(r As ShapeRange)
    Dim ref As Shape, s As Shape, i As Long, k As Long, mr As Single, ms As Single
    Set ref = r(1)
    If ref.Adjustments.Count = 0 Then Err.Raise vbObjectError + 11, , "a 1a forma nao tem pontas nem cantos ajustaveis."
    mr = IIf(ref.Width < ref.Height, ref.Width, ref.Height)
    For i = 2 To r.Count
        Set s = r(i)
        If s.Type = msoAutoShape And s.AutoShapeType = ref.AutoShapeType Then
            ms = IIf(s.Width < s.Height, s.Width, s.Height)
            s.Adjustments(1) = ref.Adjustments(1) * mr / ms
            For k = 2 To ref.Adjustments.Count
                s.Adjustments(k) = ref.Adjustments(k)
            Next
        End If
    Next
End Sub

' Pergunta um inteiro positivo; 0 = cancelado.
Public Function AskInt(msg As String, ByVal deflt As Long) As Long
    Dim v As String
    v = Ask(msg, CStr(deflt))
    If Trim$(v) = "" Then Exit Function
    If Not IsNumeric(v) Then Err.Raise vbObjectError + 10, , "digite so o numero."
    If CLng(v) < 1 Then Err.Raise vbObjectError + 10, , "use um numero a partir de 1."
    AskInt = CLng(v)
End Function

' Posiciona e dimensiona ignorando a trava de proporcao (restaurada no fim).
Public Sub SetBox(s As Shape, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim wasLocked As MsoTriState  ' "Lock" e palavra reservada do VBA
    If w < 0 Or h < 0 Then Err.Raise vbObjectError + 5, , "a forma '" & s.Name & "' esta do outro lado da referencia."
    wasLocked = s.LockAspectRatio
    s.LockAspectRatio = msoFalse
    s.Width = w: s.Height = h
    s.Left = l: s.Top = t
    s.LockAspectRatio = wasLocked
End Sub

' Centraliza cada objeto na celula da tabela sobre a qual esta o centro dele.
' shortcut: ignora celulas mescladas (usa a grade de linhas e colunas); tratar se aparecer tabela mesclada.
Public Sub AlignInTable(r As ShapeRange)
    Dim tbl As Shape, s As Shape, n As Long, cl As Single, ct As Single, cw As Single, ch As Single
    For Each s In r
        If s.HasTable Then Set tbl = s: Exit For
    Next
    If tbl Is Nothing Then Err.Raise vbObjectError + 8, , "selecione a tabela junto com os objetos."
    For Each s In r
        If s.Name <> tbl.Name Then
            If CellAt(tbl, s.Left + s.Width / 2, s.Top + s.Height / 2, cl, ct, cw, ch) Then
                s.Left = cl + (cw - s.Width) / 2
                s.Top = ct + (ch - s.Height) / 2
                n = n + 1
            End If
        End If
    Next
    If n = 0 Then Err.Raise vbObjectError + 8, , "nenhum objeto esta sobre uma celula da tabela."
End Sub

Private Function CellAt(tbl As Shape, ByVal x As Single, ByVal y As Single, cl As Single, ct As Single, cw As Single, ch As Single) As Boolean
    Dim i As Long
    If x < tbl.Left Or y < tbl.Top Then Exit Function
    cl = tbl.Left
    For i = 1 To tbl.Table.Columns.Count
        cw = tbl.Table.Columns(i).Width
        If x < cl + cw Then Exit For
        cl = cl + cw
    Next
    If i > tbl.Table.Columns.Count Then Exit Function
    ct = tbl.Top
    For i = 1 To tbl.Table.Rows.Count
        ch = tbl.Table.Rows(i).Height
        If y < ct + ch Then Exit For
        ct = ct + ch
    Next
    CellAt = (i <= tbl.Table.Rows.Count)
End Function

' Escala os objetos pelo fator em %: tamanho (a partir do canto superior esquerdo), espessura da linha e fonte.
Public Sub MagicResize(r As ShapeRange, ByVal pct As Single)
    Dim s As Shape, f As Single, wasLocked As MsoTriState, i As Long
    f = pct / 100
    For Each s In r
        wasLocked = s.LockAspectRatio
        s.LockAspectRatio = msoFalse
        s.Width = s.Width * f: s.Height = s.Height * f
        s.LockAspectRatio = wasLocked
        On Error Resume Next  ' imagens e tabelas nao tem linha ou texto do jeito comum
        If s.Line.Visible Then s.Line.Weight = s.Line.Weight * f
        If s.HasTextFrame Then
            For i = 1 To s.TextFrame.TextRange.Runs.Count
                With s.TextFrame.TextRange.Runs(i).Font
                    .Size = .Size * f
                End With
            Next
        End If
        On Error GoTo 0
    Next
End Sub

' Pergunta o fator do redimensionador; 0 = cancelado.
Public Function AskPercent() As Single
    Dim v As String
    v = Ask("Escala em % (80 reduz, 120 aumenta). Vale para tamanho, linha e fonte.", "100")
    If Trim$(v) = "" Then Exit Function
    If Not IsNumeric(v) Then Err.Raise vbObjectError + 10, , "digite so o numero, por exemplo 80."
    If CSng(v) <= 0 Or CSng(v) > 1000 Then Err.Raise vbObjectError + 10, , "use um valor entre 1 e 1000."
    AskPercent = CSng(v)
End Function

Public Function Floor0(ByVal x As Single) As Single
    If x < 0 Then Floor0 = 0 Else Floor0 = x
End Function

' Indices de r ordenados pela esquerda (ou pelo topo).
Public Function SortedIdx(r As ShapeRange, horizontal As Boolean) As Long()
    Dim idx() As Long, i As Long, j As Long, t As Long
    ReDim idx(1 To r.Count)
    For i = 1 To r.Count: idx(i) = i: Next
    For i = 1 To r.Count - 1
        For j = i + 1 To r.Count
            If PosOf(r(idx(j)), horizontal) < PosOf(r(idx(i)), horizontal) Then t = idx(i): idx(i) = idx(j): idx(j) = t
        Next
    Next
    SortedIdx = idx
End Function

Private Function PosOf(s As Shape, horizontal As Boolean) As Single
    If horizontal Then PosOf = s.Left Else PosOf = s.Top
End Function
