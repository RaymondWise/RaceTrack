VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameControl 
   Caption         =   "UserForm1"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   -465
   ClientWidth     =   7065
   OleObjectBlob   =   "GameControl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGo_Click()
On Error GoTo errHandler
Dim Vx As Integer
Vx = cmbVx.Value
Dim Vy As Integer
Vy = cmbVy.Value

Dim x As Integer
Dim y As Integer
Dim intCase As Integer
Dim MoveMe As Range

If LabelP1.Visible = True Then
    intCase = 1
    Else: intCase = 2
End If

Select Case intCase
    Case 1
    'Speed
    x = GameBoardSheet.Range("A100") + Vx
    y = GameBoardSheet.Range("A101") + Vy
    GameBoardSheet.Range("A100") = x
    GameBoardSheet.Range("A101") = y
    
    'Move
    With Cells(Int(CurrentRow.Value), Int(CurrentCol.Value))
        .ClearContents
        .Interior.ColorIndex = xlNone
        'Excel uses (rows,cols) notation, so Y direction is first
        'We're using (-y) so that positive 1 moves upward
        Set MoveMe = .Offset(-y, x)
    End With
        
            If MoveMe.Interior.ColorIndex = xlNone Then
                MoveMe = "P1"
                MoveMe.Interior.ColorIndex = 8
                Range("A102") = MoveMe.Row
                Range("A103") = MoveMe.Column
            Else: GoTo WinLose
            End If
    
    'set up form for next player
    LabelP1.Visible = False
    LabelP2.Visible = True
    CurrentX.Text = Range("A200")
    CurrentY.Text = Range("A201")
    CurrentRow.Text = Range("A202")
    CurrentCol.Text = Range("A203")
    Exit Sub
    
    'Player 2 turn
    Case 2
    'Speed
    x = GameBoardSheet.Range("A200") + Vx
    y = GameBoardSheet.Range("A201") + Vy
    GameBoardSheet.Range("A200") = x
    GameBoardSheet.Range("A201") = y
    
    'Move
    With Cells(Int(CurrentRow.Value), Int(CurrentCol.Value))
        .ClearContents
        .Interior.ColorIndex = xlNone
        Set MoveMe = .Offset(-y, x)
    End With
        
            If MoveMe.Interior.ColorIndex = xlNone Then
                MoveMe = "P2"
                MoveMe.Interior.ColorIndex = 3
                Range("A202") = MoveMe.Row
                Range("A203") = MoveMe.Column
            Else: GoTo WinLose
            End If
    
    'set up form for next player
    LabelP2.Visible = False
    LabelP1.Visible = True
    CurrentX.Text = Range("A100")
    CurrentY.Text = Range("A101")
    CurrentRow.Text = Range("A102")
    CurrentCol.Text = Range("A103")
    Exit Sub
    End Select

'TODO: Create function
WinLose:
    If MoveMe.Interior.ColorIndex = xlAutomatic Then
        MsgBox ("You Win!")
        MoveMe = "P1"
        MoveMe.Interior.ColorIndex = 6
    Else: MsgBox ("Whoops, you crashed!")
    End If
    Unload GameControl
    Exit Sub
'TODO: Create Function
errHandler:
    MsgBox ("Please select your values")
End Sub

Private Sub UserForm_Initialize()
    'Placement of Form - works well on some machines, not perfect on others
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + 30
    
    'Populate the combobox lists with an array upon initialization - this way they will always retain the values I set here
    cmbVx.List = Array("-1", "0", "1")
    cmbVy.List = Array("-1", "0", "1")
    
    'Player1 goes first
    LabelP1.Visible = True
    LabelP2.Visible = False
    CurrentRow.Text = Range("A102").Value
    CurrentCol.Text = Range("A103").Value
    CurrentX.Text = Range("A100").Value
    CurrentY.Text = Range("A101").Value
End Sub
