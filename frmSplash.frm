VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1185
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   4230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   Icon            =   "frmSplash.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSplash 
      Interval        =   500
      Left            =   90
      Top             =   870
   End
   Begin VB.PictureBox picShred 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   825
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   0
      Width           =   840
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   855
      Picture         =   "frmSplash.frx":08D9
      ScaleHeight     =   585
      ScaleWidth      =   3360
      TabIndex        =   0
      Top             =   30
      Width           =   3360
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
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
      Height          =   465
      Left            =   975
      TabIndex        =   2
      Top             =   660
      Width           =   3135
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmSplash
'
' Description:   This is the Microsoft "About" screen that is distributed with
'                visual Basic.  It has been modified to my specifications.
'
' Special thanks to Randy Birch http://www.mvps.org/vbnet/index.html
' Great web site for code snippets with examples.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Reformatted display data on screen
' ***************************************************************************
Option Explicit

Private Sub Form_Load()
      
    ' Display this form while
    ' other folrms are being loaded
    With frmSplash
        .lblCopyright.Caption = "Copyright © Kenneth Ives" & vbNewLine & _
                                 "All rights reserved"
        .tmrSplash.Enabled = True
        
        ' center form on screen
        .Move (Screen.Width - frmSplash.Width) \ 2, (Screen.Height - frmSplash.Height) \ 2
        DoEvents
        .Show vbModeless
        .Refresh
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload frmSplash         ' Deactivate form
    Set frmSplash = Nothing  ' Unload form from memory

End Sub

Private Sub tmrSplash_Timer()

    ' gblnFormsLoaded is set to FALSE when application is
    ' first started.  When last form is finished loading,
    ' flag is set to TRUE.
    If gblnFormsLoaded Then
        tmrSplash.Enabled = False  ' Turn off timer
        frmMain.Reset_frmMain      ' Display main form
        Unload Me                  ' Unload this form
    End If
    
End Sub
