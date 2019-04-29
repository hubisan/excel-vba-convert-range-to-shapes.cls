VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConvertRangeToShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This content is released under the (http://opensource.org/licenses/MIT) MIT License.
' Copyright (c) Daniel Hubmann (hubisan@gmail.com)

' *´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸

' Class Module: ConvertRangeToShapes
' Version: 1.0.0

' Convert an Excel range (texts, borders, fills) to shapes.

' DESCRIPTION
' ---------------------------------------------------------------------------------------

' Convert texts, borders and fills of an Excel range to shapes.
' The shapes version is as identical to the Excel range as possible.

' This can be used to easily copy an Excel range as it is to PowerPoint.

' Remark for texts:
'   Make sure to explicitly set the text alignment to be able to copy
'   the alignment for the shapes. Else it can lead to not identically aligned texts.
'   If a text spans multiple cells the cells need to be merged to make the shape cover
'   the same range. Else the text might get trunctuated.

' Examples:


' Motivation:

' I was not satisfied with any of the pasting options:
'   - Pasting and keeping formats often leads to misalignments and changes of the size.
'   - If pasted as image or embedded the quality is bad.

' Todo:
'   * Someday: Use a freehand form for all borders that have corners. This improves the
'     quality of the corners. Such borders are rarely used.

' *´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸

' =======================================================================================
' GLOBAL VARIABLES
' =======================================================================================

' Ranges for cells that have to be ignore drawing border lines because already a rectangle.
Private rngIgnoreTop As Range, rngIgnoreLeft As Range

' Name for the shapes. Needed for deleting.
Private strShapesName As String

' =======================================================================================
' MAIN
' =======================================================================================

' **
' Convert an Excel range (texts, borders, fills) to shapes.
'
' Convert an Excel range to shapes. Deletes existing shapes with the same name beforehand
' and selects the shapes after to copy them.
'
' Args:
'   rng (Range): Range to convert into shapes.
'   strNameForShapes (String): Name to use for the grouped shapes.
'   [ignoreFillColor] (Long): If specified the fill (interior) color will be ignored.
'     Useful to ignore white as fill color.
'
' Examples:
'   convertAll ThisWorkbook.Worksheets("-").Range("B8:C12"), "test", RGB(255, 255, 255)
Public Sub convertAll(ByVal rng As Range, ByVal strNameForShapes As String, _
Optional ByVal lngIgnoreFillColor As Long = -1)
    Dim rng1 As Range
    
    ' Set global name for shapes.
    setShapesName strNameForShapes
    
    ' Delete the grouped shapes if it already exists.
    deleteShapes rng.Parent, strShapesName
    
    ' Convert the range to shapes.
    convertTexts rng
    convertBorders rng
    convertFills rng, lngIgnoreFillColor
    
    ' Select the grouped shapes.
    selectShapes rng.Parent
End Sub

' =======================================================================================
' FUNCTIONS
' =======================================================================================

' **
' Set the global name for the shapes.
'
' Args:
'   strNameForShapes (String): Name to use for the grouped shapes.
Public Sub setShapesName(ByVal strNameForShapes As String)
    strShapesName = strNameForShapes
End Sub

' **
' Delete shapes with specified name in a workskheet.
'
' Args:
'   wks (Worksheet): Worksheet to delete the shapes in.
'   [strNameOfShapes] (String): Name of shapes to delete.
'     Use this name if set else use the global variable created with convert.
'     Wildcards can be used: * ? #
'
' Examples:
'   deleteShapes activesheet, "name*"
Public Sub deleteShapes(ByVal wks As Worksheet, Optional ByVal strNameOfShapes As String)
    Dim shp As Shape
    
    ' If strNameOfShapes is not set use global variable.
    If strNameOfShapes = vbNullString Then strNameOfShapes = strShapesName
    
    ' Delete if name is valid.
    If strNameOfShapes = vbNullString Then
        Err.Raise vbObjectError + 513, "ConvertRangeToShapes.selectShapes", _
        "Please set a name for the shapes."
    Else
        For Each shp In wks.Shapes
            If shp.Name Like strNameOfShapes Then shp.Delete
        Next
    End If
End Sub

'**
' Select shapes with specified name in a workskheet. Activates the worksheet.
'
' Args:
'   wks (Worksheet): Worksheet to select the shapes in.
'   [strNameOfShapes] (String): Name of shapes to select.
'     Use this name if set else use the global variable created with convert.
'     Wildcards can be used: * ? #
'
' Examples:
'   selectShapes activesheet, "name*"
Public Sub selectShapes(ByVal wks As Worksheet, Optional ByVal strNameOfShapes As String)
    Dim shp As Shape
    
    ' If strNameOfShapes is not set use global variable.
    If strNameOfShapes = vbNullString Then strNameOfShapes = strShapesName
    
    ' Select if name is valid.
    If strNameOfShapes = vbNullString Then
        Err.Raise vbObjectError + 513, "ConvertRangeToShapes.selectShapes", _
        "Please set a name for the shapes."
    Else
        wks.Activate
        wks.Range("A1").Activate
        ActiveCell.Select
        For Each shp In wks.Shapes
            If shp.Name Like strNameOfShapes Then
                shp.Select False
                'False = extend instead of replace the current selection.
            End If
        Next
    End If
End Sub

' **
' Convert borders to shapes.
'
' Note:
' convertRectangleBorders have to be ran before to make convertLineBorders ignore cells that
' already have a rectangle border.
'
' Args:
'   rngToConvert (Range): Range to convert borders in.
'   [strNameForShapes] (String): Name for shapes.
'
' Examples:
'   convertBorders Range("A2:B3"), "name"
Public Sub convertBorders(ByVal rngToConvert As Range, Optional ByVal strNameForShapes As String)
    Dim rngMax As Range, rng As Range

    ' Resize needed to draw borders at edges.
    Set rngMax = rngToConvert.Resize(rngToConvert.Rows.Count + 1, rngToConvert.Columns.Count + 1)
    
    ' If strNameForShapes is not set use global variable.
    If strNameForShapes = vbNullString Then strNameForShapes = strShapesName
    
    ' Only convert if name is valid.
    If strNameForShapes = vbNullString Then
        Err.Raise vbObjectError + 513, "ConvertRangeToShapes.convertRectangleBorders", _
        "Please set a name for the shapes."
    Else
        For Each rng In rngMax
            convertRectangleBorder rng, rngMax, strNameForShapes
            convertLineBorder rng, rngMax, strNameForShapes
        Next
    End If
End Sub

' **
' Convert rectangle border to shape.
'
' Checks if there is a identical border surrounding one or more cells.
' If there is a border around it use a rectangle to draw the border instead of lines.
' This makes the corners look better than when using single lines.
'
' Note:
'   Not using .offset() because if doing so the linestyle for the 2nd and last cell
'   in a merge area is not correct (it always return xlLineStyleNone).
'   Strangely when using .cells() it is correct.
'
' Args:
'   rngStart (Range): Current cell to convert borders in.
'   rngMax (Range): Range to convert borders in.
'   [strNameForShapes] (String): Name for shapes.
Private Sub convertRectangleBorder(ByVal rgnStart As Range, ByVal rngMax As Range, _
ByVal strNameForShapes As String)
    Dim rng As Range, rngMaxHori As Range
    Dim wks As Worksheet, bdr As Border
    Dim rng1 As Range, rng1Start As Range, rng2 As Range, rng3 As Range
    Dim lngMaxVertiRow As Long
    Dim shp As Shape
    
    If rgnStart.Row < rngMax(rngMax.Cells.Count).Row And _
    rgnStart.Column < rngMax(rngMax.Cells.Count).Column Then
        Set rng = rgnStart
        Set rngMaxHori = rgnStart
        Set wks = rgnStart.Parent
        Set bdr = rgnStart.Borders(xlEdgeTop)
        
        ' Rectangle is only possible if left border is the same as top and if there is a border at all.
        If identicalBorders(bdr, rgnStart.Borders(xlEdgeLeft)) And Not bdr.LineStyle = xlLineStyleNone Then
                    
            ' Move to the right until max range with identical border.
            Do While identicalBorders(bdr, wks.Cells(rng.Row, rng.Column + 1).Borders(xlEdgeTop)) _
            And rng.Column + 1 < rngMax(rngMax.Cells.Count).Column
                Set rng = wks.Cells(rng.Row, rng.Column + 1)
            Loop
            Set rngMaxHori = rng
            
            ' Go through each rng from start to max.
            For Each rng In wks.Range(rgnStart, rngMaxHori)
                
                ' If the range has an identical right border follow it down.
                If identicalBorders(bdr, rng.Borders(xlEdgeRight)) Then
                    Set rng1 = rng
                    Set rng1Start = rng
                        
                    ' Follows the right border down.
                    Do While identicalBorders(bdr, wks.Cells(rng1.Row + 1, rng1.Column).Borders(xlEdgeRight)) _
                    And rng1.Row + 1 < rngMax(rngMax.Cells.Count).Row
                        Set rng1 = wks.Cells(rng1.Row + 1, rng1.Column)
                    Loop
                    
                    ' Move down the range with identical right border.
                    For Each rng2 In wks.Range(rng1Start, rng1)
                        
                        'if the range has an identical bottom border follow it left
                        If identicalBorders(bdr, rng2.Borders(xlEdgeBottom)) Then
                            Set rng3 = rng2
                            
                            'follow the border to the left
                            Do While identicalBorders(bdr, wks.Cells(rng3.Row, rng3.Column - 1).Borders(xlEdgeBottom)) _
                            And rng3.Column > rgnStart.Column
                                Set rng3 = wks.Cells(rng3.Row, rng3.Column - 1)
                            Loop
                            
                            'if at the same column as start (needed for rectangle) and indentical border on the left
                            'follow the border up
                            If rng3.Column = rgnStart.Column And identicalBorders(bdr, rng3.Borders(xlEdgeLeft)) Then
                            
                                'follows the border up
                                Do While identicalBorders(bdr, wks.Cells(rng3.Row - 1, rng3.Column).Borders(xlEdgeLeft)) _
                                And rng3.Row > rgnStart.Row
                                    Set rng3 = wks.Cells(rng3.Row - 1, rng3.Column)
                                Loop
                                
                                'if back at start and the row of the bottom right corner is bigger than any rectangle before
                                'then it's a new rectangle that can be drawn
                                If Not Application.Intersect(rng3, rgnStart) Is Nothing Then
                                    If rng2.Row > lngMaxVertiRow Then
                                        'success
                                        lngMaxVertiRow = rng2.Row
                                        With wks.Range(rgnStart, rng2)
                                            Set shp = wks.Shapes.AddShape(msoShapeRectangle, .Left, .Top, .Width, .Height)
                                            shp.Fill.Visible = msoFalse
                                            shp.ZOrder msoSendToBack
                                            shp.Name = strNameForShapes
                                            makeShapeLineMatchCellBorder shp, bdr
                                            If rngIgnoreTop Is Nothing Then
                                                Set rngIgnoreTop = Application.Union(.Rows(1), .Rows(.Rows.Count + 1))
                                            Else
                                                Set rngIgnoreTop = Application.Union(rngIgnoreTop, .Rows(1), .Rows(.Rows.Count + 1))
                                            End If
                                            If rngIgnoreLeft Is Nothing Then
                                                Set rngIgnoreLeft = Application.Union(.Columns(1), .Columns(.Columns.Count + 1))
                                            Else
                                                Set rngIgnoreLeft = Application.Union(rngIgnoreLeft, .Columns(1), .Columns(.Columns.Count + 1))
                                            End If
                                        End With
                                    End If
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
    End If
End Sub

' **
' Convert border to shapes unless a rectangle.
'
' Follows the line as long as it can.
'
' Note:
'   Not using .offset() because if doing so the linestyle for the 2nd and last cell
'   in a merge area is not correct (it always return xlLineStyleNone).
'   Strangely when using .cells() it is correct.
'
' Args:
'   rngStart (Range): Current cell to convert borders in.
'   rngMax (Range): Range to convert borders in.
'   [strNameForShapes] (String): Name for shapes.
Private Sub convertLineBorder(ByVal rngStart As Range, ByVal rngMax As Range, _
ByVal strNameForShapes As String)
    Dim shp As Shape, bolNotInIgnore As Boolean, rng As Range
    Dim bdr As Border
    
    ' Draw vertical lines as long as possible if more than 1 cell
    If rngStart.Row < rngMax(rngMax.Cells.Count).Row Then
        Set rng = rngStart
        Set bdr = rngStart.Borders(xlEdgeLeft)
        If rngIgnoreLeft Is Nothing Then
            bolNotInIgnore = True
        Else
            bolNotInIgnore = Application.Intersect(rng, rngIgnoreLeft) Is Nothing
        End If
        If Not bdr.LineStyle = xlLineStyleNone And bolNotInIgnore Then
            ' Move to the right until max range with identical border.
            Do While identicalBorders(bdr, rng.Parent.Cells(rng.Row + 1, rng.Column).Borders(xlEdgeLeft)) _
            And bolNotInIgnore And rng.Row + 1 < rngMax(rngMax.Cells.Count).Row
                Set rng = rng.Parent.Cells(rng.Row + 1, rng.Column)
                If Not rngIgnoreLeft Is Nothing Then bolNotInIgnore = Application.Intersect(rng, rngIgnoreLeft) Is Nothing
            Loop
            With rng.Parent.Range(rngStart, rng)
                If rngIgnoreLeft Is Nothing Then
                    Set rngIgnoreLeft = .Cells()
                Else
                    Set rngIgnoreLeft = Application.Union(rngIgnoreLeft, .Cells())
                End If
                Set shp = .Parent.Shapes.AddConnector(msoConnectorStraight, .Left, .Top, .Left, .Top + .Height)
                    makeShapeLineMatchCellBorder shp, .Borders(xlEdgeLeft)
                    shp.ZOrder msoSendToBack
                    shp.Name = strNameForShapes
            End With
        End If
    End If
    
    ' Draw horizontal lines as long as possible if more than one cell.
    If rngStart.Column < rngMax(rngMax.Cells.Count).Column Then
        Set rng = rngStart
        Set bdr = rngStart.Borders(xlEdgeTop)
        If rngIgnoreTop Is Nothing Then
            bolNotInIgnore = True
        Else
            bolNotInIgnore = Application.Intersect(rng, rngIgnoreTop) Is Nothing
        End If
        If Not bdr.LineStyle = xlLineStyleNone And bolNotInIgnore Then
            ' Move to the right until max range with identical border.
            Do While identicalBorders(bdr, rng.Parent.Cells(rng.Row, rng.Column + 1).Borders(xlEdgeTop)) _
            And bolNotInIgnore And rng.Column + 1 < rngMax(rngMax.Cells.Count).Column
                Set rng = rng.Parent.Cells(rng.Row, rng.Column + 1)
                If Not rngIgnoreTop Is Nothing Then bolNotInIgnore = Application.Intersect(rng, rngIgnoreTop) Is Nothing
            Loop
            With rng.Parent.Range(rngStart, rng)
                If rngIgnoreTop Is Nothing Then
                    Set rngIgnoreTop = .Cells()
                Else
                    Set rngIgnoreTop = Application.Union(rngIgnoreTop, .Cells())
                End If
                Set shp = .Parent.Shapes.AddConnector(msoConnectorStraight, .Left, .Top, .Left + .Width, .Top)
                    makeShapeLineMatchCellBorder shp, .Borders(xlEdgeTop)
                    shp.ZOrder msoSendToBack
                    shp.Name = strNameForShapes
            End With
        End If
    End If
End Sub


' **
' Make the line of the shape match the border of the cell (color, weight, style).
'
' Note:
'   Not all lines styles are supported as they don't exists for shapes:
'     xlDouble, xlSlantDashDot
'
' Args:
'   shp (Shape): The shape to modify the line of.
'   bdr (Border): The border to
Private Sub makeShapeLineMatchCellBorder(ByVal shp As Shape, ByVal bdr As Border)
    With shp.Line
        .ForeColor.RGB = bdr.Color
        Select Case bdr.LineStyle
                Case xlContinuous
                    .DashStyle = msoLineSolid
                Case xlDot
                    .DashStyle = msoLineSysDash
                Case xlDash
                    .DashStyle = msoLineDash
                Case xlDashDot
                    .DashStyle = msoLineDashDot
                Case xlDashDotDot
                    .DashStyle = msoLineDashDotDot
                Case Else
                    .DashStyle = msoLineSolid
        End Select
        Select Case bdr.Weight
                Case xlThin
                .Weight = 0.75
            Case xlHairline, xlThin
                .Weight = 0.75
                .DashStyle = msoLineSysDot
            Case xlMedium
                .Weight = 1.5
            Case xlThick
                .Weight = 2.25
        End Select
        
    End With
End Sub

' **
' Check if two borders are identical (style, weight and color).
'
' Args:
'   firstBorder (Border): First border to compare with second.
'   secondBorder (Border): Second border to compare with first.
'
' Returns:
'   True if the borders are identical else false.
Private Function identicalBorders(ByVal firstBorder As Border, ByVal secondBorder As Border) As Boolean
    identicalBorders = _
        firstBorder.LineStyle = secondBorder.LineStyle _
        And firstBorder.Weight = secondBorder.Weight _
        And firstBorder.Color = secondBorder.Color
End Function


' **
' Convert fills to shapes unless the color is ignored.
'
' Args:
'   rngToConvert (Range): Range to convert fills in.
'   [ignoreFillColor] (Long): If specified the fill (interior) color will be ignored.
'     Useful to ignore white as fill color.
'   [strNameForShapes] (String): Name for shapes.
' Examples:
'   convertFills Range("A2:B3"), "name"
Public Sub convertFills(ByVal rngToConvert As Range, _
Optional lngIgnoreFillColor As Long = -1, Optional ByVal strNameForShapes As String)
    Dim shp As Shape, lng_color As Long, backgroundColor As Boolean
    Dim rng As Range, rngMax As Range
    
    ' Resize needed to draw borders at edges.
    Set rngMax = rngToConvert.Resize(rngToConvert.Rows.Count + 1, rngToConvert.Columns.Count + 1)
    
    ' If strNameForShapes is not set use global variable.
    If strNameForShapes = vbNullString Then strNameForShapes = strShapesName
    
    ' Only convert if name is valid.
    If strNameForShapes = vbNullString Then
        Err.Raise vbObjectError + 513, "ConvertRangeToShapes.convertRectangleBorders", _
        "Please set a name for the shapes."
    Else
        For Each rng In rngMax
            If rng.Row < rngMax(rngMax.Cells.Count).Row And _
            rng.Column < rngMax(rngMax.Cells.Count).Column Then
                backgroundColor = False
                With rng
                    If .DisplayFormat.Interior.Color <> lngIgnoreFillColor And .DisplayFormat.Interior.ColorIndex <> xlNone Then
                        lng_color = .DisplayFormat.Interior.Color
                        backgroundColor = True
                    ElseIf .Interior.Color <> lngIgnoreFillColor And .Interior.ColorIndex <> xlNone Then
                        lng_color = .Interior.Color
                        backgroundColor = True
                    End If
                    'create shape if color is not ignored and it has a fill
                    If backgroundColor Then
                        Set shp = rng.Parent.Shapes.AddShape(msoShapeRectangle, .Left, .Top, .Width, .Height)
                        shp.Name = strNameForShapes
                        shp.Fill.ForeColor.RGB = lng_color
                        shp.Line.Visible = msoFalse
                        shp.ZOrder msoSendToBack
                    End If
                End With
            End If
        Next
    End If
End Sub

' **
' Convert texts to shapes.
'
' Args:
'   rngToConvert (Range): Range to convert fills in.
'   [strNameForShapes] (String): Name for shapes.
' Examples:
'   convertTexts Range("A2:B3"), "name"
Public Sub convertTexts(ByVal rngToConvert As Range, Optional ByVal strNameForShapes As String)
    Dim shp As Shape, rngForText As Range
    Dim rng As Range, rngMax As Range
    
    ' Resize needed to draw borders at edges.
    Set rngMax = rngToConvert.Resize(rngToConvert.Rows.Count + 1, rngToConvert.Columns.Count + 1)
    
    ' If strNameForShapes is not set use global variable.
    If strNameForShapes = vbNullString Then strNameForShapes = strShapesName
    
    ' Only convert if name is valid.
    If strNameForShapes = vbNullString Then
        Err.Raise vbObjectError + 513, "ConvertRangeToShapes.convertRectangleBorders", _
        "Please set a name for the shapes."
    Else
        For Each rng In rngMax
            If rng.Row < rngMax(rngMax.Cells.Count).Row And _
            rng.Column < rngMax(rngMax.Cells.Count).Column Then
                If Not IsError(rng.Value) Then
                    If rng.Value <> vbNullString Then
                        
                        'CREATE THE TEXTBOX
                        '---------------------------------------------------------------------------------------------------
                        Set rngForText = rng
                        If rngForText.MergeCells Then Set rngForText = rngForText.MergeArea
                        Set shp = rng.Parent.Shapes.AddTextbox(msoTextOrientationHorizontal, rng.Left, rng.Top, _
                        rngForText.Width, rngForText.Height)
                        
                        With shp
                            .Name = strNameForShapes
                            'copy text from range to textbox
                            .TextFrame2.TextRange.Text = rng.Text
                            
                            'remove border and background
                            .Fill.Visible = msoFalse
                            .Line.Visible = msoFalse
                            
                            'apply the formats
                            With .TextFrame2
                                
                                .MarginBottom = 1
                                .MarginLeft = 2
                                .MarginRight = 1
                                .MarginTop = 1
                                .AutoSize = msoAutoSizeNone
                                
                                Select Case rng.WrapText
                                Case True
                                    .WordWrap = msoTrue
                                Case Else
                                    .WordWrap = msoFalse
                                End Select
                                
                                Select Case rng.HorizontalAlignment
                                Case xlRight
                                    .TextRange.ParagraphFormat.Alignment = msoAlignRight
                                    If rng.IndentLevel > 0 Then
                                        .TextRange.ParagraphFormat.RightIndent = rng.IndentLevel * 7
                                    End If
                                Case xlCenter
                                    .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                                Case Else
                                    .TextRange.ParagraphFormat.Alignment = msoAlignLeft
                                    If rng.IndentLevel > 0 Then
                                        .TextRange.ParagraphFormat.LeftIndent = rng.IndentLevel * 7
                                    End If
                                End Select
                                
                                Select Case rng.VerticalAlignment
                                Case xlBottom
                                    .VerticalAnchor = msoAnchorBottom
                                Case xlCenter
                                    .VerticalAnchor = msoAnchorMiddle
                                Case Else
                                    .VerticalAnchor = msoAnchorTop
                                End Select
                                
                                
                                Select Case rng.Orientation
                                Case -4171
                                    .Orientation = msoTextOrientationUpward
                                Case Else
                                    .Orientation = msoTextOrientationHorizontal
                                End Select
                                With .TextRange.Font
                                    .Name = rng.Font.Name
                                    .Bold = rng.Font.Bold
                                    .Italic = rng.Font.Italic
                                    .Size = rng.Font.Size
                                    .Fill.ForeColor.RGB = rng.Font.Color
                                End With
                            End With
                        End With
                    End If
                End If
            End If
        Next
    End If
End Sub

