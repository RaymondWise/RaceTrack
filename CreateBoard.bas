Attribute VB_Name = "CreateBoard"
Option Explicit
Sub CreateGrid()
Attribute CreateGrid.VB_ProcData.VB_Invoke_Func = " \n14"
    'Store background color in a variable so that adjusting only takes one edit
    Const BACKGROUND_COLOR As Long = vbBlack
    'In the properties of my worksheet, I gave the WS object an inherent name (like Sheet8), but called it GameBoardSheet
    With GameBoardSheet
        .Name = "GameBoard"
        Columns("B:Y").ColumnWidth = 2.14
        Columns("A").ColumnWidth = 50
        Columns("Z").ColumnWidth = 50
        Rows(1).RowHeight = 100
        Rows(24).RowHeight = 100
        Range("A1:Z1").Merge
        Range("A1").Interior.Color = BACKGROUND_COLOR
        Range("A24:Z24").Merge
        Range("A24").Interior.Color = BACKGROUND_COLOR
        Range("A2:A23").Merge
        Range("A2").Interior.Color = BACKGROUND_COLOR
        Range("Z2:Z23").Merge
        Range("z2").Interior.Color = BACKGROUND_COLOR
        Range("B2").Select
    End With
End Sub

Sub FillOuterGrid()
Dim i As Integer
Dim rngCell As Range
    For Each rngCell In Range("B2:Y2")
      i = Application.WorksheetFunction.RandBetween(0, 2)
      rngCell.Offset(i, 0).Interior.ColorIndex = 15
    Next
    For Each rngCell In Range("b23:Y23")
        i = Application.WorksheetFunction.RandBetween(-2, 0)
        rngCell.Offset(i, 0).Interior.ColorIndex = 15
    Next
    For Each rngCell In Range("B5:B20")
        i = Application.WorksheetFunction.RandBetween(0, 2)
        rngCell.Offset(0, i).Interior.ColorIndex = 15
    Next
    For Each rngCell In Range("Y5:Y20")
        i = Application.WorksheetFunction.RandBetween(-2, 0)
        rngCell.Offset(0, i).Interior.ColorIndex = 15
    Next
    
    For Each rngCell In Range("B4:Y4")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(-1).Interior.ColorIndex = 15
            rngCell.Offset(-2).Interior.ColorIndex = 15
        End If
    Next
    For Each rngCell In Range("B3:Y3")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(-1).Interior.ColorIndex = 15
        End If
    Next
    For Each rngCell In Range("B21:Y21")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(1).Interior.ColorIndex = 15
            rngCell.Offset(2).Interior.ColorIndex = 15
        End If
    Next
    For Each rngCell In Range("B22:Y22")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(1).Interior.ColorIndex = 15
        End If
    Next
    
    For Each rngCell In Range("D2:D23")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(, -1).Interior.ColorIndex = 15
            rngCell.Offset(, -2).Interior.ColorIndex = 15
        End If
    Next
    For Each rngCell In Range("C2:C23")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(, -1).Interior.ColorIndex = 15
        End If
    Next
    For Each rngCell In Range("W2:W23")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(, 1).Interior.ColorIndex = 15
            rngCell.Offset(, 2).Interior.ColorIndex = 15
        End If
    Next
    For Each rngCell In Range("X2:X23")
        If rngCell.Interior.ColorIndex = 15 Then
            rngCell.Offset(, 1).Interior.ColorIndex = 15
        End If
    Next
End Sub

Sub FillInnerCircle()
Dim rngCell As Range
Dim i As Integer
Range("J11:P14").Interior.ColorIndex = 15
For Each rngCell In Range("J9:P9")
    i = Application.WorksheetFunction.RandBetween(0, 1)
    rngCell.Offset(i).Interior.ColorIndex = 15
Next
For Each rngCell In Range("J16:P16")
    i = Application.WorksheetFunction.RandBetween(-1, 0)
    rngCell.Offset(i).Interior.ColorIndex = 15
Next

For Each rngCell In Range("H11:H14")
    i = Application.WorksheetFunction.RandBetween(0, 1)
    rngCell.Offset(, i).Interior.ColorIndex = 15
Next
For Each rngCell In Range("R11:R14")
    i = Application.WorksheetFunction.RandBetween(-1, 0)
    rngCell.Offset(, i).Interior.ColorIndex = 15
Next

'fill
For Each rngCell In Range("J9:P9")
    If rngCell.Interior.ColorIndex = 15 Then
        rngCell.Offset(1).Interior.ColorIndex = 15
    End If
Next
For Each rngCell In Range("J16:P16")
    If rngCell.Interior.ColorIndex = 15 Then
        rngCell.Offset(-1).Interior.ColorIndex = 15
    End If
Next
For Each rngCell In Range("H11:H14")
    If rngCell.Interior.ColorIndex = 15 Then
        rngCell.Offset(, 1).Interior.ColorIndex = 15
    End If
Next
For Each rngCell In Range("R11:R14")
    If rngCell.Interior.ColorIndex = 15 Then
        rngCell.Offset(, -1).Interior.ColorIndex = 15
    End If
Next
'start and end
With Range("M17:M20").Interior
        .Pattern = xlUp
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .PatternTintAndShade = 0
End With
Range("N19").Interior.ColorIndex = 3
Range("N19") = "P2"
Range("N20").Interior.ColorIndex = 8
Range("N20") = "P1"
End Sub

Sub StoreSpeed()
'I'm storing speed and position in cells on the sheet as I don't have a global variable for them
Range("A100") = 0
Range("A101") = 0
Range("A102") = 20
Range("A103") = 14
Range("A200") = 0
Range("A201") = 0
Range("A202") = 19
Range("A203") = 14
End Sub


Sub Button1_Click()
    MsgBox ("This will create a new gameboard")
    Application.ScreenUpdating = False
    Range("A1:Z24").ClearContents
    Range("A1:Z24").ClearFormats
    CreateGrid
    FillOuterGrid
    FillInnerCircle
    StoreSpeed
    Application.ScreenUpdating = True
    Instruct.Show
    GameControl.Show
End Sub
