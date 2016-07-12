Public roundCount As Integer
Public missMatchCount As Integer
Public emptyCmdBxCount As Integer
Public RunPauseInt As Integer
Public RunWhen As Double
Public cmdbut1Str As String
Public cmdbut2Str As String
Public cmdbut3Str As String
Public cmdbut4Str As String
Public cmdbutCurrStr As String
Public RunPauseLabel As String
Public Const cRunWhat = "myProcedure"

Sub StartTimer()
RunWhen = Now + TimeSerial(0, 0, 1)
Application.OnTime earliesttime:=RunWhen, procedure:=cRunWhat, _
     schedule:=True
End Sub

Public Sub StopTimer()
   On Error Resume Next
   Application.OnTime earliesttime:=RunWhen, _
       procedure:=cRunWhat, schedule:=False
End Sub

Sub myProcedure()
    'MsgBox "hello world"
    If roundCount < 1 Then
    putValInClip
    End If
    Call CountEmpty
    Call clibB
    Call StartTimer
    Call RunPauseLabelVal
    Call cmdBoxColorMatching
    Call ifBoxHasSpcOrZero
    'Call missingBlock
End Sub

Public Sub clibB()


UserForm1.TextBox1.Value = ""
cmdbutCurrStr = ""
UserForm1.TextBox1.Paste

If roundCount < 1 Then
UserForm1.TextBox2.Value = UserForm1.TextBox1.Value
End If
UserForm1.Label3 = emptyCmdBxCount




If Not UserForm1.TextBox1.Value = UserForm1.TextBox3.Value Then
UserForm1.TextBox2.Value = UserForm1.TextBox1.Value


missMatchCount = missMatchCount + 1

If UserForm1.CommandButton1.Caption = "" Then
        cmdbut1Str = Trim(UCase(UserForm1.TextBox1.Value))
        UserForm1.CommandButton1.Caption = cmdbut1Str
        
    ElseIf UserForm1.CommandButton2.Caption = "" Then
        cmdbut2Str = Trim(UCase(UserForm1.TextBox1.Value))
        UserForm1.CommandButton2.Caption = cmdbut2Str
        
    ElseIf UserForm1.CommandButton3.Caption = "" Then
        cmdbut3Str = Trim(UCase(UserForm1.TextBox1.Value))
        UserForm1.CommandButton3.Caption = cmdbut3Str
    
    ElseIf UserForm1.CommandButton4.Caption = "" Then
        cmdbut4Str = Trim(UCase(UserForm1.TextBox1.Value))
        UserForm1.CommandButton4.Caption = cmdbut4Str
    
End If


End If


roundCount = roundCount + 1
UserForm1.Label1 = roundCount
UserForm1.Label2 = missMatchCount

Dim objData As New MSForms.DataObject
        Dim strText
        A = UserForm1.TextBox3.Value
        strText = UserForm1.TextBox3.Value
        objData.SetText strText
        objData.PutInClipboard
        
cmdbutCurrStr = UserForm1.TextBox2.Value

End Sub

Public Sub clearClip()
UserForm1.CommandButton1.Caption = ""
UserForm1.CommandButton2.Caption = ""
UserForm1.CommandButton3.Caption = ""
UserForm1.CommandButton4.Caption = ""
UserForm1.TextBox1.Value = ""
UserForm1.TextBox2.Value = ""
UserForm1.TextBox5.Value = ""
UserForm1.Label1 = ""
UserForm1.Label2 = ""
UserForm1.CommandButton1.BackColor = &H80000005
UserForm1.CommandButton2.BackColor = &H80000005
UserForm1.CommandButton3.BackColor = &H80000005
UserForm1.CommandButton4.BackColor = &H80000005
roundCount = 0
missMatchCount = 0
cmdbut1Count = 0
UserForm1.CheckBox1 = False
End Sub

Public Sub putValInClip()
UserForm1.TextBox3.Value = "~"
Dim objData As New MSForms.DataObject
        Dim strText
        A = UserForm1.TextBox3.Value
        strText = UserForm1.TextBox3.Value
        objData.SetText strText
        objData.PutInClipboard
End Sub

Public Sub CountEmpty()
emptyCmdBxCount = 0
    If UserForm1.CommandButton1.Caption = "" Then
        emptyCmdBxCount = emptyCmdBxCount + 1
    End If
    
    If UserForm1.CommandButton2.Caption = "" Then
        emptyCmdBxCount = emptyCmdBxCount + 1
    End If
    If UserForm1.CommandButton3.Caption = "" Then
        emptyCmdBxCount = emptyCmdBxCount + 1
    End If
    If UserForm1.CommandButton4.Caption = "" Then
        emptyCmdBxCount = emptyCmdBxCount + 1
    End If
    emptyCmdBxCount = emptyCmdBxCount
End Sub

Public Sub VINonlyLenCount()

If Not UserForm1.CommandButton1.Caption = "" And Not Len(UserForm1.CommandButton1.Caption) = 17 Then
UserForm1.CommandButton1.BackColor = &HFF&
End If

If Not UserForm1.CommandButton2.Caption = "" And Not Len(UserForm1.CommandButton2.Caption) = 17 Then
UserForm1.CommandButton2.BackColor = &HFF&
End If

If Not UserForm1.CommandButton3.Caption = "" And Not Len(UserForm1.CommandButton3.Caption) = 17 Then
UserForm1.CommandButton3.BackColor = &HFF&
End If

If Not UserForm1.CommandButton4.Caption = "" And Not Len(UserForm1.CommandButton4.Caption) = 17 Then
UserForm1.CommandButton4.BackColor = &HFF&
End If

End Sub

Public Sub cmdBoxColorMatching()
'IF THEY ALL MATCH
'UserForm1.CommandButton5.BackColor = RGB(0, 51, 120)
'UserForm1.CommandButton5.BackColor = RGB(0, 0, 120)




'IF THERE BLANK, THE BACK COLOR WILL BE WHITE
If UserForm1.CommandButton1.Caption = "" Then
    'UserForm1.CommandButton1.BackColor = RGB(0, 0, 120)
    UserForm1.CommandButton1.BackColor = RGB(0, 0, 120)
End If
If UserForm1.CommandButton2.Caption = "" Then
    UserForm1.CommandButton2.BackColor = RGB(0, 0, 120)
    'UserForm1.CommandButton1.BackColor = RGB(0, 51, 205)
End If
If UserForm1.CommandButton3.Caption = "" Then
    UserForm1.CommandButton3.BackColor = RGB(0, 0, 120)
    'UserForm1.CommandButton1.BackColor = RGB(0, 51, 120)
End If
If UserForm1.CommandButton4.Caption = "" Then
    UserForm1.CommandButton4.BackColor = RGB(0, 0, 120)
    'UserForm1.CommandButton1.BackColor = RGB(0, 51, 120)
End If

'IF PAUSE BUTTON IS SELECTED AND BUCKET IS EMPTY, THEY STAY SOME COLOR AS USER FORM
If UserForm1.CommandButton1.Caption = "" And UserForm1.ToggleButton1 = True Then
    UserForm1.CommandButton1.BackColor = RGB(0, 191, 255)
End If
If UserForm1.CommandButton2.Caption = "" And UserForm1.ToggleButton1 = True Then
    UserForm1.CommandButton2.BackColor = RGB(0, 191, 255)
End If
If UserForm1.CommandButton3.Caption = "" And UserForm1.ToggleButton1 = True Then
    UserForm1.CommandButton3.BackColor = RGB(0, 191, 255)
End If
If UserForm1.CommandButton4.Caption = "" And UserForm1.ToggleButton1 = True Then
    UserForm1.CommandButton4.BackColor = RGB(0, 191, 255)
End If



'IF THEY MATCH ANOTHER BOX, THERE GREEN
If Not UserForm1.CommandButton1.Caption = "" Then
    If UserForm1.CommandButton1.Caption = UserForm1.CommandButton2.Caption And UserForm1.CommandButton1.Caption = UserForm1.CommandButton3.Caption And UserForm1.CommandButton1.Caption = UserForm1.CommandButton4.Caption Then
        UserForm1.CommandButton1.BackColor = &HFF00&
        UserForm1.CommandButton2.BackColor = &HFF00&
        UserForm1.CommandButton3.BackColor = &HFF00&
        UserForm1.CommandButton4.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton1.Caption = UserForm1.CommandButton2.Caption And UserForm1.CommandButton1.Caption = UserForm1.CommandButton3.Caption Then
        UserForm1.CommandButton1.BackColor = &HFF00&
        UserForm1.CommandButton2.BackColor = &HFF00&
        UserForm1.CommandButton3.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton1.Caption = UserForm1.CommandButton2.Caption Then
        UserForm1.CommandButton1.BackColor = &HFF00&
        UserForm1.CommandButton2.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton1.Caption = UserForm1.CommandButton3.Caption And UserForm1.CommandButton1.Caption = UserForm1.CommandButton4.Caption Then
        UserForm1.CommandButton1.BackColor = &HFF00&
        UserForm1.CommandButton3.BackColor = &HFF00&
        UserForm1.CommandButton4.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton1.Caption = UserForm1.CommandButton2.Caption And UserForm1.CommandButton1.Caption = UserForm1.CommandButton4.Caption Then
        UserForm1.CommandButton1.BackColor = &HFF00&
        UserForm1.CommandButton2.BackColor = &HFF00&
        UserForm1.CommandButton4.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton1.Caption = UserForm1.CommandButton3.Caption Then
        UserForm1.CommandButton1.BackColor = &HFF00&
        UserForm1.CommandButton3.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton1.Caption = UserForm1.CommandButton4.Caption Then
        UserForm1.CommandButton1.BackColor = &HFF00&
        UserForm1.CommandButton4.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton2.Caption = UserForm1.CommandButton4.Caption And Not UserForm1.CommandButton2.Caption = "" Then
        UserForm1.CommandButton2.BackColor = &HFF00&
        UserForm1.CommandButton4.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton3.Caption = UserForm1.CommandButton4.Caption And Not UserForm1.CommandButton4.Caption = "" Then
        UserForm1.CommandButton3.BackColor = &HFF00&
        UserForm1.CommandButton4.BackColor = &HFF00&
    ElseIf UserForm1.CommandButton2.Caption = UserForm1.CommandButton3.Caption And Not UserForm1.CommandButton2.Caption = "" Then
        UserForm1.CommandButton2.BackColor = &HFF00&
        UserForm1.CommandButton3.BackColor = &HFF00&
    End If
End If



'IF THERE NOT BLANK AND THERE NOT GREEN BY NOW, THERE YELLOW
If Not UserForm1.CommandButton1.Caption = "" And Not UserForm1.CommandButton1.BackColor = &HFF00& Then
    UserForm1.CommandButton1.BackColor = &HFFFF&
End If
If Not UserForm1.CommandButton2.Caption = "" And Not UserForm1.CommandButton2.BackColor = &HFF00& Then
    UserForm1.CommandButton2.BackColor = &HFFFF&
End If
If Not UserForm1.CommandButton3.Caption = "" And Not UserForm1.CommandButton3.BackColor = &HFF00& Then
    UserForm1.CommandButton3.BackColor = &HFFFF&
End If
If Not UserForm1.CommandButton4.Caption = "" And Not UserForm1.CommandButton4.BackColor = &HFF00& Then
    UserForm1.CommandButton4.BackColor = &HFFFF&
End If

'IF VIN CHECKBOX IS CHECKED THEN IF THE BOX ISNT BLANK AND THE BOX CHAR COUNT IS NOT 17, ITS RED
If UserForm1.CheckBox2 = True Then
    If Not UserForm1.CommandButton1.Caption = "" And Not Len(UserForm1.CommandButton1.Caption) = 17 Then
        UserForm1.CommandButton1.BackColor = &HFF&
    End If
    If Not UserForm1.CommandButton2.Caption = "" And Not Len(UserForm1.CommandButton2.Caption) = 17 Then
        UserForm1.CommandButton2.BackColor = &HFF&
    End If
    If Not UserForm1.CommandButton3.Caption = "" And Not Len(UserForm1.CommandButton3.Caption) = 17 Then
        UserForm1.CommandButton3.BackColor = &HFF&
    End If
    If Not UserForm1.CommandButton4.Caption = "" And Not Len(UserForm1.CommandButton4.Caption) = 17 Then
        UserForm1.CommandButton4.BackColor = &HFF&
    End If
End If
End Sub

Public Sub OpenUF()
UserForm1.Show vbModeless
End Sub

Public Sub RunPauseLabelVal()
RunPauseInt = RunPauseInt + 1
If RunPauseInt = 1 Then
    RunPauseLabel = "Running ."
    ElseIf RunPauseInt = 2 Then
        RunPauseLabel = "Running . ."
    ElseIf RunPauseInt = 3 Then
        RunPauseLabel = "Running . . ."
        RunPauseInt = 0
End If
UserForm1.Label4 = RunPauseLabel
End Sub

Public Sub ifBoxHasSpcOrZero()
Dim Label1 As String
Dim Label2 As String
Dim Label3 As String
Dim Label4 As String

'LABEL 1....
If UserForm1.CheckBox2 = True Then

        If InStr(UserForm1.CommandButton1.Caption, "O") > 0 And InStr(UserForm1.CommandButton1.Caption, "I") > 0 And InStr(UserForm1.CommandButton1.Caption, " ") > 0 Then
        Label1 = "Contains: O, I, space"
        UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton1.Caption, "O") > 0 And InStr(UserForm1.CommandButton1.Caption, "I") > 0 Then
        Label1 = "Contains O and I"
        UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton1.Caption, "O") > 0 And InStr(UserForm1.CommandButton1.Caption, " ") > 0 Then
        Label1 = "Contains O and Space"
        UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton1.Caption, "I") > 0 And InStr(UserForm1.CommandButton1.Caption, " ") > 0 Then
        Label1 = "Contains I and Space"
        UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton1.Caption, "O") > 0 Then
        Label1 = "Contains O"
        UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton1.Caption, "I") > 0 Then
        Label1 = "Contains I"
        UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton1.Caption, " ") > 0 Then
        Label1 = "Contains a Space"
        'UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
        UserForm1.CommandButton1.ForeColor = RGB(128, 0, 255)
    Else
    Label1 = ""
    End If
    
'LABEL 2...
    If InStr(UserForm1.CommandButton2.Caption, "O") > 0 And InStr(UserForm1.CommandButton2.Caption, "I") > 0 And InStr(UserForm1.CommandButton2.Caption, " ") > 0 Then
        Label2 = "Contains: O, I, space"
        UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton2.Caption, "O") > 0 And InStr(UserForm1.CommandButton2.Caption, "I") > 0 Then
        Label2 = "Contains O and I"
        UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton2.Caption, "O") > 0 And InStr(UserForm1.CommandButton2.Caption, " ") > 0 Then
        Label2 = "Contains O and Space"
        UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton2.Caption, "I") > 0 And InStr(UserForm1.CommandButton2.Caption, " ") > 0 Then
        Label2 = "Contains I and a Space"
        UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton2.Caption, "O") > 0 Then
        Label2 = "Contains O"
        UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton2.Caption, "I") > 0 Then
        Label2 = "Contains I"
        UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton2.Caption, " ") > 0 Then
        Label2 = "Contains a Space"
        'UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
        UserForm1.CommandButton2.ForeColor = RGB(128, 0, 255)
    Else
    Label2 = ""
    End If
'LABEL 3...
    If InStr(UserForm1.CommandButton3.Caption, "O") > 0 And InStr(UserForm1.CommandButton3.Caption, "I") > 0 And InStr(UserForm1.CommandButton3.Caption, " ") > 0 Then
        Label3 = "Contains: O, I, space"
        UserForm1.CommandButton3.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton3.Caption, "O") > 0 And InStr(UserForm1.CommandButton3.Caption, "I") > 0 Then
        Label3 = "Contains O and I"
        UserForm1.CommandButton3.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton3.Caption, "O") > 0 And InStr(UserForm1.CommandButton3.Caption, " ") > 0 Then
        Label3 = "Contains O and Space"
        UserForm1.CommandButton3.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton3.Caption, "I") > 0 And InStr(UserForm1.CommandButton3.Caption, " ") > 0 Then
        Label3 = "Contains I and a Space"
        UserForm1.CommandButton3.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton3.Caption, "O") > 0 Then
        Label3 = "Contains O"
        UserForm1.CommandButton3.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton3.Caption, "I") > 0 Then
        Label3 = "Contains I"
        UserForm1.CommandButton3.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton3.Caption, " ") > 0 Then
        Label3 = "Contains a Space"
        UserForm1.CommandButton3.ForeColor = RGB(128, 0, 255)
    Else
    Label3 = ""
    End If
'LABEL 4...
    If InStr(UserForm1.CommandButton4.Caption, "O") > 0 And InStr(UserForm1.CommandButton4.Caption, "I") > 0 And InStr(UserForm1.CommandButton4.Caption, " ") > 0 Then
        Label4 = "Contains: O, I, space"
        UserForm1.CommandButton4.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton4.Caption, "O") > 0 And InStr(UserForm1.CommandButton4.Caption, "I") > 0 Then
        Label4 = "Contains O and I"
        UserForm1.CommandButton4.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton4.Caption, "O") > 0 And InStr(UserForm1.CommandButton4.Caption, " ") > 0 Then
        Label4 = "Contains O and Space"
        UserForm1.CommandButton4.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton4.Caption, "I") > 0 And InStr(UserForm1.CommandButton4.Caption, " ") > 0 Then
        Label4 = "Contains I and a Space"
        UserForm1.CommandButton4.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton4.Caption, "O") > 0 Then
        Label4 = "Contains O"
        UserForm1.CommandButton4.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton4.Caption, "I") > 0 Then
        Label4 = "Contains I"
        UserForm1.CommandButton4.ForeColor = RGB(128, 0, 255)
    ElseIf InStr(UserForm1.CommandButton4.Caption, " ") > 0 Then
        Label4 = "Contains a Space"
        UserForm1.CommandButton4.ForeColor = RGB(128, 0, 255)
        Else
        Label4 = ""
    End If
    
    If Not UserForm1.CommandButton1.Caption = "" And Not Len(UserForm1.CommandButton1.Caption) = 17 Then
        Label1 = Label1 + " Not 17 Characters"
    End If

    If Not UserForm1.CommandButton2.Caption = "" And Not Len(UserForm1.CommandButton2.Caption) = 17 Then
        Label2 = Label2 + " Not 17 Characters"
    End If

    If Not UserForm1.CommandButton3.Caption = "" And Not Len(UserForm1.CommandButton3.Caption) = 17 Then
        Label3 = Label3 + " Not 17 Characters"
    End If

    If Not UserForm1.CommandButton4.Caption = "" And Not Len(UserForm1.CommandButton4.Caption) = 17 Then
        Label4 = Label4 + " Not 17 Characters"
    End If
    
Else
Label1 = ""
Label2 = ""
Label3 = ""
Label4 = ""
End If
    UserForm1.labelcmdbx1 = Label1
    UserForm1.labelcmdbx2 = Label2
    UserForm1.labelcmdbx3 = Label3
    UserForm1.labelcmdbx4 = Label4
End Sub

