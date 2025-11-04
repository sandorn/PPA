Sub FormatAllTables()
    On Error Resume Next
    Application.Visible = False
    
    Dim sld As Slide
    Dim shp As Shape

    Set sld = Application.ActiveWindow.View.Slide

    For Each shp In sld.Shapes
        If shp.HasTable Then
            FormatSingleTable shp.Table, 2
        End If
    Next shp
    
    Application.Visible = True
    Application.ActiveWindow.Activate
End Sub

Private Sub FormatSingleTable(tbl As Table, ByVal decimalPlaces As Long)
    Const thin As Single = 1#
    Const thick As Single = 2#
    Const fontSize As Single = 9#
    Const bigFontSize As Single = 10#

    Dim rows As Long: rows = tbl.Rows.Count
    Dim cols As Long: cols = tbl.Columns.Count
    Dim r As Long, c As Long
    Dim cell As Cell
    Dim txtRng As TextRange
    
    Dim txtColor As MsoRGBType: txtColor = msoThemeColorText1
    Dim bdColor1 As MsoRGBType: bdColor1 = msoThemeColorAccent1
    Dim bdColor2 As MsoRGBType: bdColor2 = msoThemeColorAccent2

    ' ===== 首行处理 =====
    For c = 1 To cols
        Set cell = tbl.Cell(1, c)
        cell.Shape.Fill.Visible = msoFalse
        Set txtRng = cell.Shape.TextFrame.TextRange

        With txtRng.Font
            .Name = "+mn-lt"
            .NameFarEast = "+mn-ea"
            .Size = bigFontSize
            .Bold = msoTrue
            .Color.ObjectThemeColor = txtColor
        End With
        txtRng.ParagraphFormat.Alignment = ppAlignCenter

        With cell.Borders(ppBorderTop)
            .Weight = thick
            .ForeColor.ObjectThemeColor = bdColor1
        End With
        With cell.Borders(ppBorderBottom)
            .Weight = thick
            .ForeColor.ObjectThemeColor = bdColor1
        End With
    Next c

    ' ===== 其他行处理 =====
    For r = 2 To rows
        For c = 1 To cols
            Set cell = tbl.Cell(r, c)
            cell.Shape.Fill.Visible = msoFalse
            Set txtRng = cell.Shape.TextFrame.TextRange

            With txtRng.Font
                .Name = "+mn-lt"
                .NameFarEast = "+mn-ea"
                .Size = fontSize
                .Bold = msoFalse
                .Color.ObjectThemeColor = txtColor
            End With

            SmartNumberFormat txtRng, decimalPlaces

            With cell.Borders(ppBorderBottom)
                If r = rows Then
                    .Weight = thick
                    .ForeColor.ObjectThemeColor = bdColor1
                Else
                    .Weight = thin
                    .ForeColor.ObjectThemeColor = bdColor2
                End If
            End With
        Next c
    Next r
End Sub

Private Sub SmartNumberFormat(rng As TextRange, ByVal decimalPlaces As Long)
    Dim original As String
    Dim isPercentage As Boolean
    Dim numStr As String
    Dim numValue As Double
    Dim formatted As String
    Dim negativeColor As Long

    negativeColor = RGB(255, 0, 0)

    original = Trim(rng.Text)
    If Len(original) = 0 Then Exit Sub

    isPercentage = (Right(original, 1) = "%")
    If isPercentage Then
        numStr = Trim(Left(original, Len(original) - 1))
    Else
        numStr = original
    End If

    If Not IsNumeric(numStr) Then Exit Sub

    numStr = Replace(numStr, ",", ".")
    numValue = CDbl(numStr)

    Dim formatStr As String
    If decimalPlaces > 0 Then
        formatStr = "0." & String(decimalPlaces, "0")
    Else
        formatStr = "0"
    End If

    formatted = Format(numValue, formatStr)

    If isPercentage Then
        formatted = formatted & "%"
    End If

    If original <> formatted Then
        rng.Text = formatted
    End If

    If numValue < 0 Then
        rng.Font.Color.RGB = negativeColor
    End If
End Sub
