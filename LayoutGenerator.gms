Public Sub CreateCutlineGrid()
    ' ==========================================================================
    ' === Layout Generator for CorelDRAW                                     ===
    ' === Creates a grid of cutlines and populates it with the selected design. ===
    ' ==========================================================================

    ' --- CONFIGURATION ---
    Const MEDIA_WIDTH As Double = 650
    Const BATCH_HEIGHT As Double = 300 ' Default batch height in mm
    Const OVERLAP As Double = 1 ' How much the cutlines should extend past the grid in mm
    ' --- FINE-TUNE ALIGNMENT ---
    ' If the designs are not perfectly centered, you can adjust these offsets.
    ' Use small numbers, e.g., 0.5 or -1. All units are in mm.
    Const X_OFFSET As Double = 0
    Const Y_OFFSET As Double = 0
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

    doc.Unit = cdrMillimeter

    Dim cols As Long, rows As Long
    cols = Int(MEDIA_WIDTH / labelWidth)
    rows = Int(BATCH_HEIGHT / labelHeight)

    If cols = 0 Or rows = 0 Then
        MsgBox "The selected shape is too large for the specified media and batch size.", vbExclamation, "Layout Generator"
        Exit Sub
    End If

    Dim gridWidth As Double, gridHeight As Double
    gridWidth = cols * labelWidth
    gridHeight = rows * labelHeight

    doc.BeginCommandGroup "Create Populated Layout"

    Dim cutlineLayer As Layer
    Set cutlineLayer = doc.ActivePage.CreateLayer("Cutlines")

    Dim cutlines As New ShapeRange
    Dim i As Long, j As Long
    Dim x As Double, y As Double
    Dim line As Shape

    ' Create vertical lines with overlap
    For i = 0 To cols
        x = i * labelWidth
        If i Mod 2 = 1 Then
            Set line = cutlineLayer.CreateLineSegment(x, gridHeight + OVERLAP, x, 0 - OVERLAP)
        Else
            Set line = cutlineLayer.CreateLineSegment(x, 0 - OVERLAP, x, gridHeight + OVERLAP)
        End If
        cutlines.Add line
    Next i

    ' Create horizontal lines with overlap
    For i = 0 To rows
        y = i * labelHeight
        If i Mod 2 = 1 Then
            Set line = cutlineLayer.CreateLineSegment(gridWidth + OVERLAP, y, 0 - OVERLAP, y)
        Else
            Set line = cutlineLayer.CreateLineSegment(0 - OVERLAP, y, gridWidth + OVERLAP, y)
        End If
        cutlines.Add line
    Next i

    Dim cutlineGroup As Shape
    Set cutlineGroup = cutlines.Group

    ' Set cutline color to Cyan
    Dim cyan As New Color
    cyan.CMYKAssign 100, 0, 0, 0
    cutlineGroup.Outline.Width = 0.076
    cutlineGroup.Outline.Color = cyan

    ' --- POPULATE GRID WITH DESIGNS ---
    Dim designs As New ShapeRange
    Dim newLabel As Shape

    For i = 0 To rows - 1
        For j = 0 To cols - 1
            Set newLabel = label.Duplicate
            x = (j * labelWidth) + (labelWidth / 2) + X_OFFSET
            y = (i * labelHeight) + (labelHeight / 2) + Y_OFFSET
            newLabel.SetPosition x, y
            designs.Add newLabel
        Next j
    Next i

    Dim designGroup As Shape
    Set designGroup = designs.Group

    ' --- END POPULATION ---

    doc.EndCommandGroup

    MsgBox "Layout generation complete! The designs and cutlines have been created as two separate groups.", vbInformation, "Layout Generator"

    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Layout Generator"
    If Not (doc Is Nothing) Then doc.EndCommandGroup
End Sub
