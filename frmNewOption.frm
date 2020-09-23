VERSION 5.00
Begin VB.Form frmNewOption 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   3945
   ControlBox      =   0   'False
   Icon            =   "frmNewOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3945
   Begin VB.CommandButton cmdCustom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Return to Options screen without saving"
      Top             =   2595
      Width           =   750
   End
   Begin VB.CommandButton cmdCustom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Save all entries and leave"
      Top             =   2595
      Width           =   750
   End
   Begin VB.Frame fraCustom 
      Height          =   2505
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   3705
      Begin VB.PictureBox picCustom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2340
         Left            =   75
         ScaleHeight     =   2340
         ScaleWidth      =   3540
         TabIndex        =   12
         Top             =   150
         Width           =   3540
         Begin VB.TextBox txtCustom 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   150
            MaxLength       =   30
            TabIndex        =   0
            Top             =   315
            Width           =   3255
         End
         Begin VB.TextBox txtCustom 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   195
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1905
            Width           =   435
         End
         Begin VB.TextBox txtCustom 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   885
            MaxLength       =   15
            TabIndex        =   4
            Top             =   1905
            Width           =   435
         End
         Begin VB.TextBox txtCustom 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1590
            MaxLength       =   15
            TabIndex        =   5
            Top             =   1905
            Width           =   435
         End
         Begin VB.TextBox txtCustom 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1905
            Width           =   435
         End
         Begin VB.TextBox txtCustom 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   2985
            MaxLength       =   15
            TabIndex        =   7
            Top             =   1905
            Width           =   435
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Up to 5 patterns"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   1
            Top             =   720
            Value           =   -1  'True
            Width           =   3165
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Multiples of one pattern"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   2
            Top             =   1005
            Width           =   3165
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   16
            Top             =   75
            Width           =   3135
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   2
            Left            =   180
            TabIndex        =   15
            Top             =   1350
            Width           =   3240
         End
         Begin VB.Label lblAsterik 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   630
            TabIndex        =   14
            Top             =   2025
            Width           =   240
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier (max 99)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1455
            TabIndex        =   13
            Top             =   1950
            Width           =   1575
         End
      End
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   2835
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmNewOption
'
' Description:   This form is displayed when a user wants to define a custom
'                wipe algorithm.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Verified screen would not be displayed during initial load
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module variables
' ***************************************************************************
  Private mblnMultiple As Boolean

Private Sub cmdCustom_Click(Index As Integer)

    Dim intIndex As Integer  ' loop counter
    Dim strData  As String   ' Display pattern in hex

    gstrCustom = vbNullString
    strData = vbNullString

    ' based on which button was clicked
    Select Case Index
           Case 0  ' save the input data

                ' see if a description has been entered
                If Len(Trim$(txtCustom(0).Text)) = 0 Then
                    InfoMsg "A description is required"
                    txtCustom(0).SetFocus
                    Exit Sub
                End If

                If mblnMultiple Then

                    If Len(Trim$(txtCustom(2).Text)) = 0 Or _
                       Val(txtCustom(2).Text) = 0 Then

                        InfoMsg "Multiplier value must be a value of 1-99."
                        txtCustom(2).Text = vbNullString
                        txtCustom(2).SetFocus
                        Exit Sub
                    End If

                    If Val(txtCustom(2).Text) > 99 Then
                        InfoMsg "Multiplier value cannot exceed 99."
                        txtCustom(2).Text = vbNullString
                        txtCustom(2).SetFocus
                        Exit Sub
                    End If

                    ' loop thru the pattern values and load
                    ' them into the appropriate strings
                    For intIndex = 1 To 2
                        If Len(Trim$(txtCustom(intIndex).Text)) > 0 Then
                            If intIndex = 1 Then
                                If Val(txtCustom(intIndex).Text) > 255 Then
                                    strData = "Random * "
                                Else
                                    strData = "0x" & Right$("00" & Hex(Val(txtCustom(intIndex).Text)), 2) & " * "
                                End If
                            Else
                                strData = strData & txtCustom(intIndex).Text
                            End If
                        End If
                    Next intIndex
                
                    strData = TrimStr(strData)
                
                Else
                    ' loop thru the pattern values and load
                    ' them into the  appropriate strings
                    For intIndex = 1 To 5
                        If Len(Trim$(txtCustom(intIndex).Text)) > 0 Then
                            If Val(txtCustom(intIndex).Text) > 255 Then
                                strData = strData & "Random, "
                            Else
                                strData = strData & "0x" & Right$("00" & Hex(Val(txtCustom(intIndex).Text)), 2) & ", "
                            End If
                        End If
                    Next intIndex
                    
                    strData = TrimStr(strData)
                    strData = Left$(strData, Len(strData) - 1)  ' remove last comma
                    
                End If

                ' see if there was any data available
                If Len(Trim$(strData)) = 0 Then
                    InfoMsg "At least one value is required"
                    txtCustom(1).SetFocus
                    Exit Sub
                End If
                   
                ' save the custom data string
                gstrCustom = gstrItemNbr & "| " & TrimStr(txtCustom(0).Text) & "|Unknown |" & strData

                frmNewOption.Hide              ' hide the screen first because it is displayed a vbModal
                Reset_frmNewOption False       ' empty input boxes
                frmOptions.Reset_frmOptions

           Case Else  ' just leave
                gstrItemNbr = vbNullString               ' reinitialize variables
                gstrCustom = vbNullString
                frmNewOption.Hide              ' hide the screen first because it is displayed a vbModal
                Reset_frmNewOption False       ' empty input boxes
                frmOptions.Reset_frmOptions
    End Select

End Sub

Public Function Reset_frmNewOption(ByVal blnModify As Boolean) As Boolean

    Dim astrData()   As String     ' array to hold hex display
    Dim astrValues() As String     ' array to hold actual values
    Dim strTemp      As String
    Dim intIndex     As Integer    ' loop counter

    On Error GoTo Reset_frmNewOption_Error

    For intIndex = 0 To 5
        frmNewOption.txtCustom(intIndex).Text = vbNullString
    Next intIndex

    If blnModify Then
        ' This is an EDIT request
        Erase astrData()
        Erase astrValues()

        If Len(Trim(gstrCustom)) = 0 Then
            Reset_frmNewOption = False
            Exit Function
        End If

        astrData = Split(gstrCustom, "|")
        strTemp = astrData(3)
        strTemp = Replace(strTemp, "Random", "999")

        intIndex = InStr(1, strTemp, "*")
        If intIndex = 0 Then
            astrValues = Split(strTemp, ",")
            mblnMultiple = False
        Else
            mblnMultiple = True
            ReDim astrValues(1)

            If IsNumeric(TrimStr(Left$(strTemp, intIndex - 1))) Then
                astrValues(0) = TrimStr(Left$(strTemp, intIndex - 1))
            Else
                astrValues(0) = "999"
            End If

            astrValues(1) = TrimStr(Mid$(strTemp, intIndex + 1))

        End If

        With frmNewOption
             .txtCustom(0).Text = astrData(1)

             For intIndex = 0 To UBound(astrValues)
                 .txtCustom(intIndex + 1).Text = astrValues(intIndex)
             Next intIndex
        End With
    Else
        frmNewOption.txtCustom(1).Text = "0"
    End If

    If mblnMultiple Then
        optChoice_Click 1
    Else
        optChoice_Click 0
    End If

    Reset_frmNewOption = True

Reset_frmNewOption_CleanUp:
    On Error GoTo 0
    Exit Function

Reset_frmNewOption_Error:
    Err.Clear
    Reset_frmNewOption = False
    Resume Reset_frmNewOption_CleanUp

End Function

Private Sub Form_Initialize()
    
    ' Make sure this form is hidden
    ' during initial load
    frmNewOption.Hide
    DoEvents
    
End Sub

Private Sub Form_Load()

    DisableX frmNewOption   ' Disable "X" in upper right corner of form
    
    With frmNewOption
        .Caption = PGM_NAME & " - Customize"
        
        .lblOption(2).Caption = "Valid values are 0-255.  Values greater than 255 will designate random data"
        .txtCustom(0).Text = vbNullString
        .txtCustom(1).Text = "0"
        .txtCustom(2).Text = vbNullString
        
        optChoice_Click 0
        gobjKeyEdit.CenterCaption frmNewOption   ' Center form window caption

        ' center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Hide
    End With

    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Based on the unload code the system passes, we determine what to do.
    '
    ' Unloadmode codes
    '     0 - Close from the control-menu box or Upper right "X"
    '     1 - Unload method from code elsewhere in the application
    '     2 - Windows Session is ending
    '     3 - Task Manager is closing the application
    '     4 - MDI Parent is closing
    Select Case UnloadMode
           Case 0    ' return to main form
                frmMain.Show
                frmNewOption.Hide
    
           Case Else
                ' Fall thru. Something else is shutting us down.
    End Select

End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub optChoice_Click(Index As Integer)

    With frmNewOption
        Select Case Index
               Case 0
                    .optChoice(0).Value = True
                    .optChoice(1).Value = False
                    .lblAsterik.Visible = False
                    .lblOption(0).Visible = False
                    .txtCustom(3).Visible = True
                    .txtCustom(4).Visible = True
                    .txtCustom(5).Visible = True
                    mblnMultiple = False
               Case Else
                    .optChoice(0).Value = False
                    .optChoice(1).Value = True
                    .lblAsterik.Visible = True
                    .lblOption(0).Visible = True
                    .txtCustom(3).Visible = False
                    .txtCustom(4).Visible = False
                    .txtCustom(5).Visible = False
                    mblnMultiple = True
        End Select
    End With

End Sub

'==================================================================
' Custom pattern text boxes
'==================================================================
Private Sub txtCustom_Change(Index As Integer)

    ' Prevent user from pasting a non-numeric value
    ' into this textbox
    If Index <> 0 Then
        If Not IsNumeric(txtCustom(Index).Text) Then
            txtCustom(Index).Text = vbNullString
        End If
    End If
    
End Sub

Private Sub txtCustom_GotFocus(Index As Integer)
    gobjKeyEdit.TextBoxFocus txtCustom(Index)
End Sub

Private Sub txtCustom_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    gobjKeyEdit.TextBoxKeyDown txtCustom(Index), KeyCode, Shift
End Sub

Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
           Case 0   ' Description
                gobjKeyEdit.ProcessAlphaNumeric KeyAscii
           Case Else  ' pattern data
                gobjKeyEdit.ProcessNumericOnly KeyAscii
    End Select
End Sub

Private Sub txtCustom_LostFocus(Index As Integer)

    Select Case Index
           Case 0   ' Description
                ' no need to test the edit at this time
           Case Else  ' pattern data
                If Len(Trim$(txtCustom(Index).Text)) > 0 Then
                    If Val(txtCustom(Index).Text) < 0 Then
                        InfoMsg "Value must be a positive number"
                        txtCustom(Index).Text = vbNullString
                        txtCustom(Index).SetFocus
                    End If
                End If
    End Select
End Sub
