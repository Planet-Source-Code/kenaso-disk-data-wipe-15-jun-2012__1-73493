VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6495
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   8100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8100
   Begin VB.PictureBox picManifest 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   0
      ScaleHeight     =   6660
      ScaleWidth      =   8205
      TabIndex        =   13
      Top             =   -120
      Width           =   8205
      Begin VB.Frame fraOptions 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4275
         Left            =   45
         TabIndex        =   16
         Top             =   1680
         Width           =   7965
         Begin MSComctlLib.ImageList imgImages 
            Left            =   6255
            Top             =   540
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":12FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":140C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvwOptions 
            Height          =   3990
            Left            =   75
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   195
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   7038
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.CommandButton cmdChoice 
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
         Index           =   4
         Left            =   7230
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Do not save any changes"
         Top             =   6060
         Width           =   750
      End
      Begin VB.CommandButton cmdChoice 
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
         Index           =   3
         Left            =   6450
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save all changes and return to main window"
         Top             =   6060
         Width           =   750
      End
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Delete"
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
         Index           =   2
         Left            =   5670
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Delete an existing wipe agorithm"
         Top             =   6060
         Width           =   750
      End
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Modify"
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
         Left            =   4905
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Modify an existing wipe algorithm"
         Top             =   6060
         Width           =   750
      End
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
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
         Left            =   4125
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Add a new wipe algorithm"
         Top             =   6060
         Width           =   750
      End
      Begin VB.Frame fraLogFile 
         Height          =   1560
         Left            =   45
         TabIndex        =   14
         Top             =   120
         Width           =   7980
         Begin VB.CheckBox chkLogEncrypt 
            Caption         =   "Save parameters  (Encryption only)"
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
            Left            =   480
            TabIndex        =   2
            Top             =   1020
            Width           =   1935
         End
         Begin VB.CheckBox chkZeroes 
            Caption         =   "Zeroes on the last pass"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   2580
            TabIndex        =   3
            ToolTipText     =   "Create a log file"
            Top             =   180
            Width           =   2310
         End
         Begin VB.CheckBox chkVerify 
            Caption         =   "Verify last pass"
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
            Left            =   2565
            TabIndex        =   4
            Top             =   705
            Width           =   1665
         End
         Begin VB.CheckBox chkLogResults 
            Caption         =   "Use log file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   1
            ToolTipText     =   "Create a log file"
            Top             =   705
            Width           =   1455
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset defaults"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   6795
            MaskColor       =   &H8000000F&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   375
            Width           =   1065
         End
         Begin VB.CheckBox chkDisplayMsgs 
            Caption         =   "Do not display verify error messages"
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
            Left            =   2760
            TabIndex        =   5
            Top             =   1020
            Width           =   2160
         End
         Begin VB.TextBox txtPasses 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5895
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "99"
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox chkFinishMsg 
            Caption         =   "Display completion message box"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   180
            TabIndex        =   0
            Top             =   180
            Value           =   1  'Checked
            Width           =   1950
         End
         Begin VB.Label lblMain 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Number of passes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   5670
            TabIndex        =   15
            Top             =   345
            Width           =   1095
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
         Left            =   150
         TabIndex        =   17
         Top             =   6180
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmOptions
'
' Description:   This form is displayed when a user wants to select a wipe
'                algorithm or other features.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Verified screen would not be displayed during initial load
' 01-Apr-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added visual pointer to show which method was selected.
'              - Updated documentation
' 01-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Updated Form_Load(), LoadGrid(), ResetIni(),
'                CheckMethodSelection(), ResizeListviewColumns() routines
'              - Updated documentation
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME              As String = "frmOptions"
  Private Const vbVerticalBar            As String = "|"
  Private Const LVM_FIRST                As Long = &H1000
  Private Const LVM_SETCOLUMNWIDTH       As Long = (LVM_FIRST + 30)
  Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
  ' Designates which image to use from ImageList control
  Private Const UNCHECKED                As Long = 1  ' Empty checkbox
  Private Const CHECKED                  As Long = 2  ' Selected item
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The SendMessage function sends the specified message to a window or
  ' windows. The function calls the window procedure for the specified
  ' window and does not return until the window procedure has processed
  ' the message.
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
          (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
          lParam As Any) As Long

  ' The LockWindowUpdate function disables or enables drawing in the
  ' specified window. Only one window can be locked at a time. Used to
  ' reduce flicker during group updates.
  Private Declare Function LockWindowUpdate Lib "user32" _
          (ByVal hwndLock As Long) As Long

' ***************************************************************************
' Module Variables
'
' Variable name:     mabytOption
' Naming standard:   m a byt Option
'                    - - --- -------
'                    | |  |   |_____ Variable subname
'                    | |  |_________ Data type (Byte)
'                    | |____________ Array designator
'                    |______________ Module level designator
'
' ***************************************************************************
  Private mlngIndex      As Long
  Private mabytOption(6) As Byte    ' options selected
  Private mastrKeys(100) As String
  
Private Sub chkFinishMsg_Click()

    If chkFinishMsg.Value = vbChecked Then
        mabytOption(1) = 1   ' set finish message to TRUE
    Else
        mabytOption(1) = 0   ' set finish message to FALSE
    End If
    
End Sub

Private Sub chkLogResults_Click()
    
    ' if checked, then display the other log options
    If chkLogResults.Value = vbChecked Then
        mabytOption(0) = 1   ' set log results to TRUE
    Else
        mabytOption(0) = 0   ' set log results to FALSE
    
        If chkLogEncrypt.Enabled = True Then
            chkLogEncrypt.Value = vbUnchecked  ' Set to FALSE
            chkLogEncrypt.Enabled = False      ' Disable checkbox
        End If
    End If
    
    CheckMethodSelection
    
End Sub

Private Sub chkDisplayMsgs_Click()
    
    ' set flag to display/hide verify messages
    If chkDisplayMsgs.Value = vbChecked Then
        mabytOption(3) = 0   ' set verify messages to FALSE
    Else
        mabytOption(3) = 1   ' set verify messages to TRUE
    End If
    
    gblnDisplayVerifyMsgs = CBool(mabytOption(3))
    
End Sub

Private Sub chkVerify_Click()
    
    If chkVerify.Value = vbChecked Then
        chkDisplayMsgs.Enabled = True  ' Enable verify messages
        mabytOption(2) = 1             ' set verify to TRUE
    Else
        chkDisplayMsgs.Value = vbUnchecked
        chkDisplayMsgs.Enabled = False ' Disable verify messages
        mabytOption(2) = 0             ' set verify flag to FALSE
        chkDisplayMsgs_Click           ' Reset verify messages to FALSE
    End If
    
End Sub

Private Sub chkZeroes_Click()

    ' set flag to write zeroes as the last pass
    ' only valid with options 2-7
    If chkZeroes.Value = vbChecked Then
        mabytOption(4) = 1   ' set to TRUE
    Else
        mabytOption(4) = 0   ' set to FALSE
    End If
    
End Sub

Private Sub chkLogEncrypt_Click()

    ' set flag to save encrypt parameters to log
    If chkLogEncrypt.Value = vbChecked Then
        mabytOption(5) = 1   ' set to TRUE
    Else
        mabytOption(5) = 0   ' set to FALSE
    End If
    
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
           Case 0    ' Add new data
                gstrItemNbr = FindFirstEmptyRow   ' Get row number of first empty row
                
                If Val(gstrItemNbr) > 99 Then
                    InfoMsg "Maximum allowed wipe patterns is 99."
                    Exit Sub
                End If
                                 
                UpdateINI
                gstrCustom = vbNullString   ' empty data string
                frmOptions.Hide   ' hide this form
                
                With frmNewOption
                    .Reset_frmNewOption (False)
                    .Show
                End With
                
           Case 1    ' Edit selected item
                If glngWipeMethod < 1 Then
                    InfoMsg "Cannot identify wipe option:  " & CStr(glngWipeMethod)
                    Exit Sub
                End If
                
                ' Warning message about permanent options
                If glngWipeMethod >= 1 Then
                    If glngWipeMethod <= PROTECTED_ITEMS Then
                            
                        InfoMsg "Cannot modify a permanent setting."
                        Exit Sub
                    
                    End If
                End If
                
                UpdateINI         ' Update INI file
                frmOptions.Hide   ' hide this form

                With frmNewOption
                    .Reset_frmNewOption (True)
                    .Show
                End With
                
           Case 2    ' Delete selected item
                ' if not a selection
                If glngWipeMethod < 1 Then
                    InfoMsg "Cannot identify wipe option:  " & CStr(glngWipeMethod)
                    Exit Sub
                End If
                
                ' Warning message about permanent options
                If glngWipeMethod >= 1 Then
                    If glngWipeMethod <= PROTECTED_ITEMS Then
                            
                        InfoMsg "Cannot delete a permanent setting."
                        Exit Sub
                    
                    End If
                End If
                
                ' ask for verification
                If ResponseMsg("Are you sure?", vbExclamation Or vbDefaultButton2 Or vbOKCancel) = vbOK Then
                    DeleteRow    ' Update grid
                End If
                
                UpdateINI
                
           Case 3    ' Save data to INI file
                UpdateINI              ' Update INI settings
                frmOptions.Hide        ' hide this form
                frmMain.Reset_frmMain  ' show the main form
                
           Case Else    ' Do not save new data
                LoadGrid               ' reload the grid with previous settings
                frmOptions.Hide        ' hide this form
                frmMain.Reset_frmMain  ' show the main form
    End Select
    
End Sub

Private Sub cmdReset_Click()

    ' Reset INI file options and load grid
    LoadGrid True

End Sub

Private Sub Form_Initialize()

    ' Make sure this form is hidden
    ' during initial load
    frmOptions.Hide
    DoEvents
    
End Sub

Private Sub Form_Load()
    
    Erase mastrKeys()  ' Always start with an empty array
    mlngIndex = 0
    
    DisableX frmOptions   ' Disable "X" in upper right corner of form
    
    ' Load grid with wiping options.
    If Not LoadGrid Then
        
        ' Try loading grid again.
        If Not LoadGrid(True) Then
            
            ' If grid does not load properly
            ' then terminate this application
            InfoMsg "Failed to load grid with wiping options." & vbNewLine & _
                    "Terminating application."
            TerminateProgram
            Exit Sub
            
        End If
    
    End If
            
    With frmOptions
        .Caption = PGM_NAME & " - Available options"
        gobjKeyEdit.CenterCaption frmOptions   ' Center form window caption
        
        ' center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Hide
    End With
            
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Erase mastrKeys()    ' Always empty arrays when not needed
    Erase mabytOption()
        
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
                frmOptions.Hide
    
           Case Else
                ' Fall thru. Something else is shutting us down.
    End Select

End Sub

Public Sub Reset_frmOptions()

    ' if changes occurred, update the grid
    If Len(Trim$(gstrCustom)) > 0 Then
        UpdateListview
        Unload frmNewOption
        Set frmNewOption = Nothing
    End If
    
    frmOptions.Show vbModeless
    frmOptions.Refresh
    
End Sub

Public Function LoadGrid(Optional ByVal blnResetIni As Boolean = False) As Boolean

    Dim lngIndex     As Long
    Dim lngPosition  As Long
    Dim lngRow       As Long
    Dim strLastEntry As String
    Dim astrData()   As String
    Dim astrKeys()   As String         ' INI key names and values
    Dim lvwListItem  As ListItem       ' grid items in listview control
    Dim lvwColHeader As ColumnHeader   ' column headers in listview control
    
    On Error GoTo LoadGrid_Error

    LoadGrid = False
    
    ' Reset INI file options
    ' and rebuild INI file
    If blnResetIni Then
        ResetIni      ' Initialize options
        BuildIniFile  ' Recreate INI file
    End If
    
    ' Create INI file if it does not exist
    If Not IsPathValid(gstrINI) Then
        ResetIni      ' Initialize options
        BuildIniFile  ' Recreate INI file
    End If
    
    gstrItemNbr = vbNullString
    lngRow = 1
    Erase astrData()   ' Always start with empty arrays
    Erase astrKeys()
    
    ' Create the column headers
    With lvwOptions
        .SmallIcons = imgImages   ' Associate ImageList control with ListView control
        .View = lvwList           ' Each ListItem is represented by a small icon and a
                                  '    text label that appears to the right of the icon
        .LabelWrap = False        ' Keep text on a single line
        .ColumnHeaders.Clear      ' Clear header titles
        .ListItems.Clear          ' Clear grid data area
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Item"
        .ColumnHeaders.Add 2, , "Description"
        .ColumnHeaders.Add 3, , "Security"
        .ColumnHeaders.Add 4, , "Comments"
    End With
    
    ' get default values
    gobjINIMgr.GetAllSectionData gstrINI, INI_PERMANENT, astrKeys()
    
    If IsArrayInitialized(astrKeys()) Then
        
        For lngIndex = 0 To 999
        
            If IsEmpty(astrKeys) Or Len(astrKeys(lngIndex)) = 0 Then
                Exit For    ' exit For..Next loop
            Else
                astrData() = Split(astrKeys(lngIndex), "|")
                
                ' load a row of new data into listview grid
                lngPosition = InStr(1, astrData(0), "=")
                
                ' Item # column with empty blue checkbox
                Set lvwListItem = lvwOptions.ListItems.Add(, , Right$("0" & Mid$(astrData(0), lngPosition + 1), 2), , UNCHECKED)
                
                With lvwListItem
                    .SubItems(1) = astrData(1)              ' Description
                    .SubItems(2) = astrData(2)              ' Security Level
                    .SubItems(3) = astrData(3) & Space$(6)  ' Comments
                End With
                
                strLastEntry = TrimStr(astrData(1))   ' Save last description
                lngRow = lngRow + 1                 ' Increment row pointer
                Erase astrData()                    ' Remove old data from array
                
            End If
        Next lngIndex
        
    End If
    
    '========================================================
    ' See if INI files needs to be updated with a newer
    ' version.  See basMain, Declare section.
    '
    ' Are number of protected items the same?
    If lngIndex <> PROTECTED_ITEMS Then
        Exit Function
    End If
    
    ' Check for same name regardless of case.
    If StrComp(LAST_ENTRY, Left$(strLastEntry, Len(LAST_ENTRY)), vbTextCompare) <> 0 Then
        Exit Function
    End If
    '========================================================
    
    ' Get custom values
    Erase astrKeys()  ' Start with empty array
    gobjINIMgr.GetAllSectionData gstrINI, INI_CUSTOM, astrKeys()
    
    ' Is there any data to work with?
    If IsArrayInitialized(astrKeys()) Then
    
        For lngIndex = 0 To 999
        
            If IsEmpty(astrKeys) Or Len(astrKeys(lngIndex)) = 0 Then
                Exit For    ' exit For..Next loop
            Else
                astrData() = Split(astrKeys(lngIndex), "|")
                    
                ' load a row of new data into the listview grid
                lngPosition = InStr(1, astrData(0), "=")
                
                ' Item # column with empty blue checkbox
                Set lvwListItem = lvwOptions.ListItems.Add(, , Right$("0" & CStr(lngRow), 2), , UNCHECKED)
                lvwListItem.SubItems(1) = astrData(1)              ' Description
                lvwListItem.SubItems(2) = astrData(2)              ' Security Level
                lvwListItem.SubItems(3) = astrData(3) & Space$(6)  ' Comments
                lngRow = lngRow + 1
                Erase astrData()
                
            End If
        Next lngIndex
    End If
    
    ReadIniFile mabytOption()     ' Get INI file settings
    txtPasses.Text = glngPasses   ' Display number of passes
    
    If gblnLogData Then
        chkLogResults.Value = vbChecked
        
        If gblnLogEncryptParms Then
            chkLogEncrypt.Value = vbChecked
        End If
    Else
        chkLogResults.Value = vbUnchecked
        chkLogEncrypt.Value = vbUnchecked
    End If
        
    If gblnVerifyData Then
        chkVerify.Value = vbChecked
        
        If gblnDisplayVerifyMsgs Then
            chkDisplayMsgs.Value = vbChecked
        Else
            chkDisplayMsgs.Value = vbUnchecked
        End If
    Else
        chkVerify.Value = vbUnchecked
        chkDisplayMsgs.Value = vbUnchecked
    End If
    
    If gblnDisplayFinishMsg Then
        chkFinishMsg.Value = vbChecked
    Else
        chkFinishMsg.Value = vbUnchecked
    End If
    
    ResizeListviewColumns
    chkLogResults_Click
    chkVerify_Click
    chkFinishMsg_Click
    UpdateINI
    
    ' Display options form if
    ' called from main form
    If Not blnResetIni Then
        frmOptions.Show
        frmOptions.Refresh
    End If
    
    LoadGrid = True
    
LoadGrid_CleanUp:
    Set lvwListItem = Nothing   ' Always free objects from memory when not needed
    Set lvwColHeader = Nothing
    Erase astrData()            ' Always empty arrays when not needed
    Erase astrKeys()
    
    On Error GoTo 0
    Exit Function

LoadGrid_Error:
    ErrorMsg MODULE_NAME, "LoadGrid", Err.Description
    Resume LoadGrid_CleanUp
    
End Function

Private Sub UpdateListview()

    Dim lngRow      As Long
    Dim astrData()  As String      ' array to hold custom data
    Dim lvwListItem As ListItem    ' grid components of listview control
    Dim objItem     As MSComctlLib.ListItem
    
    On Error GoTo UpdateListview_Error
        
    If Len(Trim$(gstrCustom)) > 0 Then
        
        Erase astrData()   ' Always start with an empty array
        
        astrData = Split(gstrCustom, vbVerticalBar)  ' load array from string of new data
        lngRow = lvwOptions.ListItems.Count + 1
            
        ' load a row of new data into the listview grid
        Set lvwListItem = lvwOptions.ListItems.Add(, , Right$("0" & CStr(lngRow), 2), , CHECKED)  ' Item # with selected blue box
        lvwListItem.SubItems(1) = astrData(1)                                                      ' Description
        lvwListItem.SubItems(2) = astrData(2)                                                      ' Security Level
        lvwListItem.SubItems(3) = astrData(3) & Space$(6)                                          ' Comments
        
        ClearCheckboxes
        
        With lvwOptions.ListItems(lngRow)
            .Selected = True     ' highlight row
            .CHECKED = True      ' Update internal checkbox
    
            For Each objItem In lvwOptions.ListItems
                If objItem.Selected Then
                    objItem.SmallIcon = CHECKED   ' Display selected blue checkbox
                    Exit For
                End If
            Next objItem
        End With
    
        glngWipeMethod = lngRow
        gstrDescription = TrimStr(lvwOptions.ListItems(lngRow).ListSubItems(1))
        chkVerify.Enabled = True
        
        DoEvents
        Select Case lngRow
               Case 2       ' Random data stream
                    chkZeroes.Enabled = True
               Case 3 To 8  ' Enable verification and zero last write
                    chkZeroes.Enabled = True
                    chkVerify.Value = vbChecked
               Case Else
                    chkVerify.Value = vbUnchecked
                    chkZeroes.Value = vbUnchecked
                    chkZeroes.Enabled = False
        End Select
                    
    End If
                
UpdateListview_CleanUp:
    gstrCustom = vbNullString
    ResizeListviewColumns
    
    DoEvents
    frmOptions.Refresh
    DoEvents
    
    Set lvwListItem = Nothing  ' Always free objects from memory when not needed
    Set objItem = Nothing
    Erase astrData()           ' Always empty arrays when not needed
    
    On Error GoTo 0
    Exit Sub

UpdateListview_Error:
    ErrorMsg MODULE_NAME, "UpdateListview", Err.Description
    Resume UpdateListview_CleanUp
        
End Sub

Private Sub DeleteRow()

    Dim lngRow  As Long
    Dim objItem As MSComctlLib.ListItem
    
    On Error GoTo DeleteRow_Error
    
    ' Remove the selected item(s) from listView starting
    ' from the bottom of the list so the data does not
    ' have to be reindexed immediately.
    With lvwOptions
        
        For lngRow = .ListItems.Count To 1 Step -1
            
            ' remove selected item(s)
            If .ListItems.Item(lngRow).Selected Then
                mastrKeys(mlngIndex) = "Item_" & CStr(Val(.ListItems.Item(lngRow).Text))
                mlngIndex = mlngIndex + 1  ' Increment module row counter
                .ListItems.Remove lngRow   ' Remove selected row
                Exit For                   ' exit For..Next loop
            End If
            
        Next lngRow
        
        ' renumber the first column
        For lngRow = 1 To .ListItems.Count
            .ListItems(lngRow).Text = Right$("0" & CStr(lngRow), 2)
        Next lngRow
    
        ' Default to first row of options
        With .ListItems(1)
            .Selected = True  ' highlight row
            .CHECKED = True   ' Update internal checkbox
        End With

        DoEvents
        
        For Each objItem In lvwOptions.ListItems
            If objItem.Selected Then
                objItem.SmallIcon = CHECKED   ' Display selected blue checkbox
                Exit For
            End If
        Next objItem
        
        glngWipeMethod = 1
        gstrDescription = TrimStr(.ListItems(1).ListSubItems(1))
        chkVerify.Enabled = True
    
    End With
    
    ResizeListviewColumns
         
    With frmOptions
        .Hide
        DoEvents
        .Refresh
        DoEvents
        .Show vbModeless
    End With
    
    UpdateINI

DeleteRow_CleanUp:
    Set objItem = Nothing  ' Always free objects from memory when not needed
    On Error GoTo 0
    Exit Sub

DeleteRow_Error:
    ErrorMsg MODULE_NAME, "DeleteRow", Err.Description
    Resume DeleteRow_CleanUp

End Sub

Private Function FindFirstEmptyRow() As String

    ' Identify first available row
    FindFirstEmptyRow = Right$("0" & CStr(lvwOptions.ListItems.Count + 1), 2)
    
End Function

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub lvwOptions_BeforeLabelEdit(Cancel As Integer)
    ' This is so no editing will be done
    ' to the individual column items
    Cancel = 1
End Sub

Private Sub lvwOptions_Click()
    
    Dim lngIndex As Long
    
    On Error GoTo lvwOptions_Click_Error
    
    With lvwOptions
        For lngIndex = 1 To .ListItems.Count
            .ListItems(lngIndex).CHECKED = False
        Next lngIndex
        
        For lngIndex = 1 To .ListItems.Count
            
            If .ListItems.Item(lngIndex).Selected = True Then
                
                With .ListItems(lngIndex)
                    .CHECKED = True
                    
                    ' save the wipe method selected
                    glngWipeMethod = Val(.Text)
                    gstrItemNbr = .Text
                    
                    ' save data from selected item
                    gstrCustom = .Text & vbVerticalBar
                    gstrCustom = gstrCustom & .ListSubItems(1) & vbVerticalBar
                    gstrCustom = gstrCustom & .ListSubItems(2) & vbVerticalBar
                    gstrCustom = gstrCustom & TrimStr(.ListSubItems(3))
                    gstrDescription = TrimStr(.ListSubItems(1))
                End With
                
                Exit For  ' exit For..Next loop
                
            End If
                
        Next lngIndex
    End With
    
    CheckMethodSelection
    
lvwOptions_Click_CleanUp:
    On Error GoTo 0
    Exit Sub

lvwOptions_Click_Error:
    Err.Clear
    Resume lvwOptions_Click_CleanUp

End Sub

Private Sub lvwOptions_ItemClick(ByVal Item As MSComctlLib.ListItem)

    For Each Item In lvwOptions.ListItems
        
        With Item
            If .CHECKED Then
                .SmallIcon = UNCHECKED   ' Display empty blue checkbox
            End If
            If .Selected Then
                .SmallIcon = CHECKED     ' Display selected blue checkbox
            End If
        End With
    
    Next Item

End Sub

Private Sub UpdateINI()

    Dim lngRow      As Long
    Dim strKey      As String      ' ini key name
    Dim lvwListItem As ListItem    ' grid components of listview control
    
    On Error GoTo UpdateINI_Error

    ' loop thru the array of items to delete
    For lngRow = 0 To mlngIndex - 1
        
        ' if no more items then exit this loop
        If Len(mastrKeys(lngRow)) = 0 Then
            Exit For    ' exit For..Next loop
        Else
            ' delete this item from the INI file
            strKey = mastrKeys(lngRow)
            gobjINIMgr.DeleteOneKey gstrINI, INI_CUSTOM, strKey
        End If
        
    Next lngRow
    
    ' Loop thru the listview grid and add all items
    ' with an item number greater than the number
    ' of protected items.
    If lvwOptions.ListItems.Count > PROTECTED_ITEMS Then
         
        gobjINIMgr.DeleteSection gstrINI, INI_CUSTOM
        gobjINIMgr.SaveSectionTitle gstrINI, INI_CUSTOM
        
        For lngRow = (PROTECTED_ITEMS + 1) To lvwOptions.ListItems.Count
                
            strKey = "Item_" & CStr(lngRow)
            
            With lvwOptions.ListItems(lngRow)
                gstrCustom = .Text & vbVerticalBar
                gstrCustom = gstrCustom & .ListSubItems(1) & vbVerticalBar
                gstrCustom = gstrCustom & .ListSubItems(2) & vbVerticalBar
                gstrCustom = gstrCustom & TrimStr(.ListSubItems(3))
            End With
            
            gobjINIMgr.SaveOneKeyValue gstrINI, INI_CUSTOM, strKey, gstrCustom
            
        Next lngRow
            
    End If
    
    ' save rest of options
    WriteIniFile mabytOption()
    
    If gblnLogData Then
        gobjINIMgr.SaveOneKeyValue gstrINI, INI_DEFAULT, "LogName", gstrLogFile
    Else
        gobjINIMgr.SaveOneKeyValue gstrINI, INI_DEFAULT, "LogName", "NO_FILE"
    End If
    
    GetDefaultPattern
    
UpdateINI_CleanUp:
    Set lvwListItem = Nothing  ' Always free objects from memory when not needed
    Erase mastrKeys()          ' Empty arrays when not needed
    mlngIndex = 0
    On Error GoTo 0
    Exit Sub

UpdateINI_Error:
    ErrorMsg MODULE_NAME, "UpdateINI", Err.Description
    Resume UpdateINI_CleanUp

End Sub

Private Sub GetDefaultPattern()
    
    Dim lngRow  As Long
    Dim objItem As MSComctlLib.ListItem
    
    On Error GoTo GetDefaultPattern_Error

    gstrCustom = vbNullString
    ClearCheckboxes
    
    With lvwOptions
    
        ' Find default pattern
        For lngRow = 1 To .ListItems.Count
            
            If Val(.ListItems(lngRow).Text) = glngWipeMethod Then
                            
                With .ListItems(lngRow)
                    .Selected = True     ' highlight row
                    .CHECKED = True      ' Update internal checkbox
        
                    For Each objItem In lvwOptions.ListItems
                        If objItem.Selected Then
                            objItem.SmallIcon = CHECKED   ' Display selected blue checkbox
                            Exit For
                        End If
                    Next objItem
                    
                    gstrItemNbr = .Text  ' capture item number text
                    
                    ' save data from selected item
                    gstrCustom = .Text & vbVerticalBar
                    gstrCustom = gstrCustom & .ListSubItems(1) & vbVerticalBar
                    gstrCustom = gstrCustom & .ListSubItems(2) & vbVerticalBar
                    gstrCustom = gstrCustom & TrimStr(.ListSubItems(3))
                    gstrDescription = TrimStr(.ListSubItems(1))
                End With
                
                CheckMethodSelection
                Exit For    ' exit For..Next loop
                            
            End If
        
        Next lngRow
    End With

GetDefaultPattern_CleanUp:
    Set objItem = Nothing  ' Always free objects from memory when not needed
    On Error GoTo 0
    Exit Sub

GetDefaultPattern_Error:
    Err.Clear
    Resume GetDefaultPattern_CleanUp
    
End Sub

Private Sub ClearCheckboxes()

    Dim lngIndex As Long
    Dim objItem  As MSComctlLib.ListItem
    
    With lvwOptions
    
        ' Uncheck all internal checkboxes
        For lngIndex = 1 To .ListItems.Count
            .ListItems(lngIndex).CHECKED = False
        Next lngIndex
        
        ' Display empty blue checkboxes
        For Each objItem In lvwOptions.ListItems
            objItem.SmallIcon = UNCHECKED
        Next objItem
    
    End With
    
    Set objItem = Nothing

End Sub

Private Sub CheckMethodSelection()

    Dim lngPos As Long
    Dim strTmp As String
    
    ' Determine the display message on the form concerning
    ' the type of pattern being used.  Same message will
    ' be used in the log file.
    Select Case glngWipeMethod
           
           Case 1   ' All zeroes (Null values)
                gstrLogPattern = TrimStr(gstrDescription & "  [ Low Security ]")
                chkZeroes.Enabled = False
                chkVerify.Enabled = True
                chkVerify.Value = vbUnchecked    ' mabytOption(2) = FALSE
           
                If chkLogEncrypt.Enabled = True Then
                    chkLogEncrypt.Value = vbUnchecked  ' preset to FALSE
                    chkLogEncrypt.Enabled = False      ' Disable checkbox
                End If
           
           Case 2   ' Random data only
                If glngPasses >= 5 Then
                    ' 5 or more passes makes this high security
                    gstrLogPattern = "Random data  [ Random * " & CStr(glngPasses) & " = High Security ]"
                Else
                    gstrLogPattern = "Random data  [ Random * " & CStr(glngPasses) & " = Medium Security ]"
                End If
                
                chkZeroes.Enabled = True
                chkVerify.Enabled = True
                chkVerify.Value = vbUnchecked    ' mabytOption(2) = FALSE
                
                If chkLogEncrypt.Enabled = True Then
                    chkLogEncrypt.Value = vbUnchecked  ' preset to FALSE
                    chkLogEncrypt.Enabled = False      ' Disable checkbox
                End If
           
           Case 3  ' US DoD Short (Verification required)
                If glngPasses > 1 Then
                    ' 2 or more passes makes this high security
                    gstrLogPattern = TrimStr(gstrDescription & "  * " & CStr(glngPasses) & " = High Security ")
                Else
                    gstrLogPattern = TrimStr(gstrDescription & "  [ Medium Security ]")
                End If
                
                chkZeroes.Enabled = True
                chkVerify.Value = vbChecked     ' mabytOption(2) = TRUE
                chkVerify.Enabled = False
                
                If chkLogEncrypt.Enabled = True Then
                    chkLogEncrypt.Value = vbUnchecked  ' preset to FALSE
                    chkLogEncrypt.Enabled = False      ' Disable checkbox
                End If
           
           Case 4 To 6  ' US DoD Long, European  (Verification required)
                gstrLogPattern = TrimStr(gstrDescription & "  [ High Security ]")
                chkZeroes.Enabled = True
                chkVerify.Value = vbChecked     ' mabytOption(2) = TRUE
                chkVerify.Enabled = False
                                
                If chkLogEncrypt.Enabled = True Then
                    chkLogEncrypt.Value = vbUnchecked  ' preset to FALSE
                    chkLogEncrypt.Enabled = False      ' Disable checkbox
                End If
           
           Case 7, 8  ' Bruce Schneier, Peter Gutmann
                gstrLogPattern = TrimStr(gstrDescription & "  [ High Security ]")
                chkZeroes.Enabled = True
                chkVerify.Enabled = True
                chkVerify.Value = vbUnchecked    ' mabytOption(2) = FALSE
           
                If chkLogEncrypt.Enabled = True Then
                    chkLogEncrypt.Value = vbUnchecked  ' preset to FALSE
                    chkLogEncrypt.Enabled = False      ' Disable checkbox
                End If
           
           Case 9 To 12   ' Encryption
                gstrLogPattern = TrimStr(gstrDescription & "  [ High Security ]")
                chkVerify.Value = vbUnchecked    ' mabytOption(2) = FALSE
                chkZeroes.Enabled = False
                chkVerify.Enabled = False
                
                If chkLogResults Then
                    chkLogEncrypt.Enabled = True  ' Enable checkbox
                    
                    ' See if previously checked
                    If gblnLogEncryptParms Then
                        chkLogEncrypt.Value = vbChecked
                    Else
                        chkLogEncrypt.Value = vbUnchecked
                    End If
                End If
                
                ' Encryption limited to 1 passes
                If (Val(txtPasses.Text) <> 1) Then
                    glngPasses = 1
                    txtPasses.Text = 1
                End If
                
           Case Else   ' Custom patterns defined by user
                lngPos = InStrRev(gstrCustom, "|")
                If lngPos > 0 Then
                    strTmp = "  [ " & Mid$(gstrCustom, lngPos + 1) & " ]"
                Else
                    strTmp = vbNullString
                End If
                
                gstrLogPattern = TrimStr(gstrDescription & "  [ Unknown Security ]") & strTmp
                chkZeroes.Value = vbUnchecked    ' mabytOption(4) = FALSE
                chkZeroes.Enabled = False
                chkVerify.Value = vbUnchecked
                
                If chkLogEncrypt.Enabled = True Then
                    chkLogEncrypt.Value = vbUnchecked  ' preset to FALSE
                    chkLogEncrypt.Enabled = False      ' Disable checkbox
                End If
    End Select
    
End Sub

'==================================================================
' Number of passes textbox
'==================================================================
Private Sub txtPasses_GotFocus()
    gobjKeyEdit.TextBoxFocus txtPasses
End Sub

Private Sub txtPasses_KeyDown(KeyCode As Integer, Shift As Integer)
    gobjKeyEdit.TextBoxKeyDown txtPasses, KeyCode, Shift
End Sub

Private Sub txtPasses_KeyPress(KeyAscii As Integer)
    gobjKeyEdit.ProcessNumericOnly KeyAscii
End Sub

Private Sub txtPasses_LostFocus()
    
    ' Make sure this is numeric data
    If IsNumeric(txtPasses.Text) Then
    
        If Len(Trim$(txtPasses.Text)) = 0 Then
            txtPasses.Text = 1
        ElseIf Len(Trim$(txtPasses.Text)) > 0 Then
            
            Select Case glngWipeMethod

                   Case 9 To 12   ' Encryption limited to 10 passes
                        If (Val(txtPasses.Text) <> 1) Then
                            txtPasses.Text = 1
                            InfoMsg "Encryption is limited to one pass."
                        End If

                   Case Else
                        If Val(txtPasses.Text) < 1 Then
                            txtPasses.Text = 1
                            InfoMsg "Minimum number of passes is one."
                        ElseIf Val(txtPasses.Text) > 99 Then
                            txtPasses.Text = 99
                            InfoMsg "Maximum number of passes is 99."
                        End If
            End Select
        End If
    Else
        ' Not numeric data. Replace data
        ' with numeric value of one.
        txtPasses.Text = 1
        InfoMsg "Minimum number of passes will be set to one."
    End If
    
    glngPasses = Val(txtPasses.Text)
    
End Sub

Private Sub ResizeListviewColumns()

    Dim intCol As Integer
    
    With lvwOptions

        ' Lock ListView display to reduce
        ' flicker while being updated
        LockWindowUpdate .hwnd

        ' Autosize columns to fit longest entry
        For intCol = 0 To .ColumnHeaders.Count - 1

            SendMessage .hwnd, LVM_SETCOLUMNWIDTH, _
                        intCol, LVSCW_AUTOSIZE_USEHEADER
        Next intCol
        
    End With
    
    ' Unlock grid display when finished
    LockWindowUpdate 0&

End Sub

Private Sub ResetIni()

    mabytOption(0) = 0   ' Log process to a file
    mabytOption(1) = 0   ' Display finish msg
    mabytOption(2) = 0   ' Perform verification
    mabytOption(3) = 0   ' Display verify messages
    mabytOption(4) = 0   ' Zero last write to file
    mabytOption(5) = 0   ' Save encrypt parms
    
    gblnLogData = CBool(mabytOption(0))
    gblnDisplayFinishMsg = CBool(mabytOption(1))
    gblnVerifyData = CBool(mabytOption(2))
    gblnDisplayVerifyMsgs = CBool(mabytOption(3))
    gblnZeroLastWrite = CBool(mabytOption(4))
    gblnLogEncryptParms = CBool(mabytOption(5))
    glngWipeMethod = 1
    glngPasses = 1
    gblnLogData = False
    
    ' save rest of options
    With gobjINIMgr
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "Default", CStr(glngWipeMethod)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "Description", gstrDescription
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "LogData", mabytOption(0)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "DisplayFinishMsg", mabytOption(1)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "Verify", mabytOption(2)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "DisplayVerifyMsg", mabytOption(3)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "ZeroLastWrite", mabytOption(4)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "LogEncryptParms", mabytOption(5)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "NbrOfPasses", glngPasses
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "LogName", "NO_FILE"
    End With
    
End Sub
