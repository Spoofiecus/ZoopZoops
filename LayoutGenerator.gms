Attribute VB_Name = "LayoutGenerator"

Public Sub CreateCutlineGrid()
    ' ==========================================================================
    ' === Layout Generator for CorelDRAW                                     ===
    ' === Creates a grid of cutlines based on a selected shape's dimensions. ===
    ' ==========================================================================

    ' --- CONFIGURATION ---
    ' The name of the spot color for the cutlines.
    ' Make sure this color exists in your document's palette as a spot color.
    Const CUTLINE_COLOR_NAME As String = "CutContour"
    ' The media width in millimeters.
    Const MEDIA_WIDTH As Double = 650
    ' --- END CONFIGURATION ---

    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    If doc Is Nothing Then
        MsgBox "No active document. Please open or create a document first.", vbExclamation, "Layout Generator"
        Exit Sub
    End If

    If doc.Selection.Shapes.Count <> 1 Then
        MsgBox "Please select exactly one shape to use as the basis for the layout.", vbExclamation, "Layout Generator"
        Exit Sub
    End If

    Dim label As Shape
    Set label = doc.Selection.Shapes(1)
    Dim labelWidth As Double, labelHeight As Double
    labelWidth = label.SizeWidth
    labelHeight = label.SizeHeight

    Dim batchHeightStr As String
    batchHeightStr = InputBox("Enter the batch height in mm:", "Layout Generator", "300")
    If Not IsNumeric(batchHeightStr) Or Val(batchHeightStr) <= 0 Then
        MsgBox "Invalid height entered. Please enter a positive number.", vbCritical, "Layout Generator"
        Exit Sub
    End If
    Dim batchHeight As Double
    batchHeight = Val(batchHeightStr)

    doc.Unit = cdrMillimeter

    Dim cols As Long, rows As Long
    cols = Int(MEDIA_WIDTH / labelWidth)
    rows = Int(batchHeight / labelHeight)

    If cols = 0 Or rows = 0 Then
        MsgBox "The selected shape is too large for the specified media and batch size.", vbExclamation, "Layout Generator"
        Exit Sub
    End If

    Dim gridWidth As Double, gridHeight As Double
    gridWidth = cols * labelWidth
    gridHeight = rows * labelHeight

    doc.BeginCommandGroup "Create Cutline Grid"

    Dim cutlineLayer As Layer
    Set cutlineLayer = doc.ActivePage.CreateLayer("Cutlines")

    Dim cutlines As New ShapeRange
    Dim i As Long

    ' Create vertical lines
    For i = 0 To cols
        Dim x As Double
        x = i * labelWidth
        Dim line As Shape
        Set line = cutlineLayer.CreateLineSegment(x, 0, x, gridHeight)
        If i Mod 2 = 1 Then
            line.Curve.SubPaths(1).Reverse
        End If
        cutlines.Add line
    Next i

    ' Create horizontal lines
    For i = 0 To rows
        Dim y As Double
        y = i * labelHeight
        Dim line As Shape
        Set line = cutlineLayer.CreateLineSegment(0, y, gridWidth, y)
        If i Mod 2 = 1 Then
            line.Curve.SubPaths(1).Reverse
        End If
        cutlines.Add line
    Next i

    Dim cutlineGroup As Shape
    Set cutlineGroup = cutlines.Group

    Dim cutlineColor As New Color
    cutlineColor.Spot.Name = CUTLINE_COLOR_NAME

    With cutlineGroup.Outline
        .SetProperties Width:=0.076, Color:=cutlineColor ' 0.076mm is hairline
    End With

    doc.EndCommandGroup

    MsgBox "Layout generation complete! " & cols & " columns and " & rows & " rows created.", vbInformation, "Layout Generator"

    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Layout Generator"
    If Not (doc Is Nothing) Then doc.EndCommandGroup
End Sub
