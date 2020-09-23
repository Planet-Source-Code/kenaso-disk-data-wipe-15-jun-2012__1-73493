VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5385
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6405
   Begin MSComDlg.CommonDialog dlgCD 
      Left            =   1110
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.Animation aniWork 
      Height          =   615
      Index           =   1
      Left            =   5430
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   41
   End
   Begin MSComCtl2.Animation aniWork 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   41
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   2
      Left            =   4800
      Picture         =   "frmMain.frx":12FA
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Display About Screen"
      Top             =   4620
      Width           =   650
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   3
      Left            =   5520
      Picture         =   "frmMain.frx":1604
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Terminate this application"
      Top             =   4620
      Width           =   650
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   1
      Left            =   4095
      Picture         =   "frmMain.frx":190E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Stop the active process"
      Top             =   4620
      Width           =   650
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   0
      Left            =   4080
      Picture         =   "frmMain.frx":1D50
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Start the wiping process"
      Top             =   4620
      Width           =   650
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   1695
      Picture         =   "frmMain.frx":205A
      ScaleHeight     =   585
      ScaleWidth      =   3360
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   45
      Width           =   3360
   End
   Begin VB.Frame fraMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   90
      TabIndex        =   36
      Top             =   855
      Width           =   6225
      Begin VB.Frame fraSelection 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   135
         TabIndex        =   37
         Top             =   135
         Width           =   5955
         Begin VB.PictureBox picManifest 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2715
            Index           =   0
            Left            =   75
            ScaleHeight     =   2715
            ScaleWidth      =   5760
            TabIndex        =   41
            Top             =   180
            Width           =   5760
            Begin VB.PictureBox picFolderOpt 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1050
               Left            =   810
               ScaleHeight     =   1050
               ScaleWidth      =   4875
               TabIndex        =   45
               Top             =   1170
               Width           =   4875
               Begin VB.OptionButton optFolders 
                  Caption         =   "Keep folder structure.  Remove files only."
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   2
                  Left            =   0
                  TabIndex        =   11
                  Top             =   525
                  Width           =   3705
               End
               Begin VB.OptionButton optFolders 
                  Caption         =   "Remove top level folder"
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
                  Left            =   0
                  TabIndex        =   9
                  Top             =   0
                  Width           =   2250
               End
               Begin VB.OptionButton optFolders 
                  Caption         =   "Keep top level folder"
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
                  Left            =   0
                  TabIndex        =   10
                  Top             =   255
                  Width           =   2310
               End
               Begin VB.Label lblPath 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name of top level folder to be removed"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   150
                  TabIndex        =   46
                  Top             =   825
                  Width           =   4635
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.CheckBox chkAlternateMethod 
               Caption         =   "Use DoD Short pattern"
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
               Left            =   2355
               TabIndex        =   6
               Top             =   795
               Width           =   3165
            End
            Begin VB.CommandButton cmdFolders 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   60
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Browse for a folder"
               Top             =   705
               Width           =   435
            End
            Begin VB.CommandButton cmdFiles 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   60
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Browse for one or more files"
               Top             =   0
               Width           =   435
            End
            Begin VB.ComboBox cboUSB 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3825
               TabIndex        =   3
               Text            =   "cboUSB"
               ToolTipText     =   "Select the target drive"
               Top             =   15
               Width           =   1920
            End
            Begin VB.OptionButton optTarget 
               Caption         =   "Wipe USB drive"
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
               Index           =   2
               Left            =   2100
               TabIndex        =   2
               ToolTipText     =   "Wipe all data on USB drive"
               Top             =   90
               Width           =   1650
            End
            Begin VB.CommandButton cmdOptions 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Options"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   3870
               TabIndex        =   12
               TabStop         =   0   'False
               ToolTipText     =   "Wiping options available"
               Top             =   2295
               Width           =   930
            End
            Begin VB.CommandButton cmdLogFile 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Log files"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4845
               Picture         =   "frmMain.frx":2604
               TabIndex        =   13
               TabStop         =   0   'False
               ToolTipText     =   "Review entries in the log file"
               Top             =   2295
               Width           =   930
            End
            Begin VB.ComboBox cboDrive 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3825
               TabIndex        =   5
               Text            =   "cboDrive"
               ToolTipText     =   "Select the target drive"
               Top             =   450
               Width           =   1920
            End
            Begin VB.OptionButton optTarget 
               Caption         =   "Wipe free space"
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
               Index           =   3
               Left            =   2100
               TabIndex        =   4
               ToolTipText     =   "Wipe the free space on one or more drives"
               Top             =   510
               Width           =   1650
            End
            Begin VB.OptionButton optTarget 
               Caption         =   "Folder"
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
               Index           =   1
               Left            =   600
               TabIndex        =   8
               ToolTipText     =   "Select a folder to be wiped"
               Top             =   855
               Width           =   855
            End
            Begin VB.OptionButton optTarget 
               Caption         =   "File(s)"
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
               Left            =   600
               TabIndex        =   1
               ToolTipText     =   "Select files to be wiped"
               Top             =   90
               Value           =   -1  'True
               Width           =   1215
            End
         End
      End
      Begin VB.Label lblPathFile 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Full path name (max 55 chars)"
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
         Left            =   120
         LinkTimeout     =   0
         TabIndex        =   38
         Top             =   3240
         Width           =   5985
      End
   End
   Begin VB.Frame fraProgress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   90
      TabIndex        =   18
      Top             =   855
      Width           =   6225
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   5895
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1845
         Width           =   5955
      End
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   5895
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5955
      End
      Begin MSComCtl2.Animation aniWork 
         Height          =   1395
         Index           =   2
         Left            =   4725
         TabIndex        =   49
         Top             =   2250
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   2461
         _Version        =   393216
         Center          =   -1  'True
         FullWidth       =   94
         FullHeight      =   93
      End
      Begin VB.Label lblProgMsg 
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
         Height          =   255
         Index           =   14
         Left            =   1575
         TabIndex        =   48
         Top             =   3060
         Width           =   3195
      End
      Begin VB.Label lblProgMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File system:"
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
         Index           =   13
         Left            =   45
         TabIndex        =   47
         Top             =   3060
         Width           =   1440
      End
      Begin VB.Label lblProgMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "99 days  12 hrs 59 mins 59 secs"
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
         Index           =   6
         Left            =   1590
         TabIndex        =   35
         Top             =   2295
         Width           =   3165
      End
      Begin VB.Label lblProgMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time elapsed:"
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
         Index           =   5
         Left            =   0
         TabIndex        =   34
         Top             =   2295
         Width           =   1485
      End
      Begin VB.Label lblProgMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time remaining:"
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
         Index           =   9
         Left            =   0
         TabIndex        =   33
         Top             =   2550
         Width           =   1485
      End
      Begin VB.Label lblProgMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data rate:"
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
         Index           =   10
         Left            =   45
         TabIndex        =   32
         Top             =   2805
         Width           =   1440
      End
      Begin VB.Label lblPattern 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Current pattern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         TabIndex        =   31
         Top             =   3330
         Width           =   4665
      End
      Begin VB.Label lblProgMsg 
         BackStyle       =   0  'Transparent
         Caption         =   " "
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
         Index           =   12
         Left            =   1590
         TabIndex        =   30
         Top             =   2805
         Width           =   3165
      End
      Begin VB.Label lblProgMsg 
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
         Height          =   255
         Index           =   11
         Left            =   1590
         TabIndex        =   29
         Top             =   2550
         Width           =   3165
      End
      Begin VB.Label lblProgMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "999,999,999"
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
         Index           =   8
         Left            =   1770
         TabIndex        =   28
         Top             =   1620
         Width           =   4230
      End
      Begin VB.Label lblProgMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblProgMsg(7)"
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
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   1620
         Width           =   1590
      End
      Begin VB.Label lblPathFile 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Full path name (max 55 chars)"
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
         Index           =   1
         Left            =   135
         LinkTimeout     =   0
         TabIndex        =   26
         Top             =   315
         Width           =   5955
      End
      Begin VB.Label lblProgMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pass 99 of 99"
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
         Left            =   4710
         TabIndex        =   25
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label lblProgMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "lblProgMsg(1)"
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
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   735
         Width           =   1590
      End
      Begin VB.Label lblProgMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "lblProgMsg(2)"
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
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1020
         Width           =   1590
      End
      Begin VB.Label lblProgMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "999,999,999"
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
         Index           =   3
         Left            =   1770
         TabIndex        =   22
         Top             =   735
         Width           =   3060
      End
      Begin VB.Label lblProgMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "999,999,999"
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
         Index           =   4
         Left            =   1770
         TabIndex        =   21
         Top             =   1020
         Width           =   4230
      End
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Height          =   660
      Left            =   180
      TabIndex        =   44
      Top             =   4605
      Width           =   2550
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   2370
      TabIndex        =   40
      Top             =   645
      Width           =   1680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' ***************************************************************************
' Module:        frmMain
'
' Description:   This is the main form for the DiscDataWipe application
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to disallow the user to wipe a drive
'              starting at the root if that drive contains the Windows
'              operating system.
'              Verified screen would not be displayed during initial load.
'              Rearranged folder options on display.
'              Updated flash drive size to 32GB.
' 24-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  If wiping free space and a drive letter is not
'              selected then recapture drive letter.
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
'              See ProgressBar() routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME             As String = "frmMain"
  Private Const TEMP_FOLDER             As String = "DD_Temp"
  Private Const EVAL_DR                 As String = " ... Evaluating drive"
  Private Const WIPE_USB                As String = "Completely overwrite USB drive "
  Private Const WIPE_FREESPACE_SPECIFIC As String = "Wipe free space on drive "
  Private Const WIPE_FREESPACE_ALL      As String = "Wipe free space on all local hard drives"
  Private Const WARNING_MSG             As String = "This application is limited to floppy, " & _
                                                    "USB drives and local fixed disk drives."
  
  Private Const MB_1                    As Long = 1048576  ' 1 mb
  Private Const PATH_LEN                As Long = 55
  
' ***************************************************************************
' Module variables
' ***************************************************************************
  Private mblnAviReady         As Boolean   ' flag to designate AVI file are ready
  Private mblnUSB_Drive        As Boolean
  Private mblnFlashDrive       As Boolean
  Private mblnDoSubfolders     As Boolean
  Private mblnWipeFreeSpace    As Boolean   ' Flag for wiping freespace only
  Private mblnAllFixedDrives   As Boolean
  Private mblnRemoveTopFolder  As Boolean
  Private mblnAlternateMethod  As Boolean   ' Flag denotes using DoD Short to wipe free space
  Private mblnKeepDirStructure As Boolean
  Private mlngPrevCurrent      As Long
  Private mlngPrevOverall      As Long
  Private mlngEncryptAlgo      As Long
  Private mlngTotalProgress    As Long
  Private mcurByteCount        As Currency  ' Total byte count
  Private mcurFileCount        As Currency  ' Total file count
  Private mcurFreeSpaceStart   As Currency
  Private mstrDrive            As String    ' selected drive letter
  Private mstrFileSys          As String
  Private mstrCaption          As String    ' Centered main caption title
  Private mstrPathFile         As String
  Private mastrAVI()           As String
  Private mastrFileList()      As String    ' List of files to be wiped
  Private mastrFixedDrives()   As String

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The GetDriveType function determines whether a disk drive is a removable,
  ' fixed, CD-ROM, RAM disk, or network drive.
  Private Declare Function GetDriveType Lib "kernel32" _
          Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

' ***************************************************************************
' Events
' ***************************************************************************
  Private WithEvents mobjWipe As kiWipe.cWipe
Attribute mobjWipe.VB_VarHelpID = -1
  
  
Private Sub cboDrive_Click()

    Dim strTest As String
    
    mblnAllFixedDrives = False
    lblPath.Caption = vbNullString
    strTest = TrimStr(Left$(cboDrive.Text, 4))
    
    If StrComp("All", strTest, vbTextCompare) = 0 Then
        
        mstrDrive = QualifyPath(mastrFixedDrives(0))
        mstrPathFile = WIPE_FREESPACE_ALL
        lblPathFile(0).Caption = WIPE_FREESPACE_ALL
        lblPathFile(1).Caption = WIPE_FREESPACE_ALL
        mblnAllFixedDrives = True
    
    ElseIf StrComp("Removable", strTest, vbTextCompare) = 0 Then
        
        mstrDrive = QualifyPath(Left$(cboDrive.Text, 2))
        lblPathFile(0).Caption = WIPE_FREESPACE_SPECIFIC & mstrDrive
        lblPathFile(1).Caption = WIPE_FREESPACE_SPECIFIC & mstrDrive
        mstrPathFile = WIPE_FREESPACE_SPECIFIC & mstrDrive
    
    Else
        
        mstrDrive = QualifyPath(Left$(cboDrive.Text, 2))
        lblPathFile(0).Caption = WIPE_FREESPACE_SPECIFIC & mstrDrive
        lblPathFile(1).Caption = WIPE_FREESPACE_SPECIFIC & mstrDrive
        mstrPathFile = WIPE_FREESPACE_SPECIFIC & mstrDrive

    End If
    
End Sub

Private Sub cboUSB_Click()

    lblPath.Caption = vbNullString
    Erase mastrFileList()
    ReDim mastrFileList(1)
    
    If Len(Trim$(cboUSB.Text)) = 0 Then
        InfoMsg "USB device is not available"
        Exit Sub
    End If
    
    mstrDrive = QualifyPath(Left$(cboUSB.Text, 2))
    mastrFileList(0) = mstrDrive
    
    lblPathFile(0).Caption = WIPE_USB & mstrDrive
    lblPathFile(1).Caption = WIPE_USB & mstrDrive
    mstrPathFile = WIPE_USB & mstrDrive
      
End Sub

Private Sub chkAlternateMethod_Click()

    ' Toggle to use an alternate wiping
    ' patttern to overwrite free space
    mblnAlternateMethod = CBool(chkAlternateMethod.Value)
    
End Sub

' ***************************************************************************
' Routine:       cmdChoice_Click
'
' Description:   Performs the main functions of this application.  There is
'                STOP, GO, EXIT.
'
' Parameters:    Index - indicates which command button was clicked
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub cmdChoice_Click(Index As Integer)

    Dim intIndex As Integer
    
    Select Case Index

           Case 0    ' GO button
                gblnStopProcessing = False       ' reset the STOP flag
                mobjWipe.StopProcessing = gblnStopProcessing
                
                ' make sure we are to make at least one pass
                ' and somthing to wipe
                If Len(Trim$(lblPathFile(0).Caption)) = 0 Then
                
                    ' display error message
                    InfoMsg "Cannot identify target to securely wipe!"
                    Reset_Screen
                    Exit Sub
                End If
                                    
                AlwaysOnTop frmMain, True   ' Set this app to be on top
                RefreshCaptions
                
                ' Make the STOP button the topmost
                ' button and hide the GO button
                With frmMain
                    .cmdChoice(0).Visible = False    ' hide go
                    .cmdChoice(1).Visible = True     ' show stop
                    .cmdChoice(2).Enabled = False    ' disable help
                    .cmdChoice(3).Enabled = False    ' disable Exit
                    
                    .fraMain.Enabled = False         ' disable input areas
                    .fraMain.Visible = False         ' hide input areas
                    .fraProgress.Visible = True      ' show the progression frame
                End With
                
                ' wiping the free space
                If mblnWipeFreeSpace Then
                    
                    ' If drive destination is missing, then get it.
                    If Len(Trim$(mstrDrive)) <= 0 Then
                        cboDrive_Click
                    End If
                            
                    If mblnAllFixedDrives Then
                        
                        For intIndex = 0 To UBound(mastrFixedDrives) - 1
                            
                            mstrDrive = QualifyPath(mastrFixedDrives(intIndex))
                            
                            ' Pass required parms to mobjWipe
                            If SaveRequiredData Then
                            
                                ' Start wiping the freespace on selected drive
                                If Not mobjWipe.WipeTheFreeSpace(mstrDrive) Then
                                    gblnStopProcessing = True
                                    mobjWipe.StopProcessing = gblnStopProcessing
                                    Exit For    ' exit For..Next loop
                                End If
                            Else
                                gblnStopProcessing = True
                                mobjWipe.StopProcessing = gblnStopProcessing
                                Exit For    ' exit For..Next loop
                            End If
                            
                            RefreshCaptions
                        
                        Next intIndex
                                                
                        DoEvents
                        If Not gblnStopProcessing Then
                            DisplayFinishMsg True
                        End If
                    
                    Else
                        ' Pass required parms to mobjWipe
                        If SaveRequiredData Then
                        
                            ' Start wiping the freespace on selected drive
                            If mobjWipe.WipeTheFreeSpace(mstrDrive) Then
                                DisplayFinishMsg True
                            Else
                                gblnStopProcessing = True
                            End If
                        Else
                            gblnStopProcessing = True
                        End If
                            
                        mobjWipe.StopProcessing = gblnStopProcessing
                    
                    End If
                    
                Else
                    If SaveRequiredData Then
                    
                        ' are we using a custom pattern?
                        If mobjWipe.WipeMethod > PROTECTED_ITEMS Then
                            CustomWipe
                        Else
                            NormalWipe
                        End If
                    End If
                End If
                
                If mblnWipeFreeSpace Then
                    RemoveWorkFolder
                End If
                
                RefreshCaptions
                Reset_Screen
                AlwaysOnTop frmMain, False
                
           Case 1   ' STOP button
                If mblnWipeFreeSpace Then
                    RemoveWorkFolder
                End If
                
                gblnStopProcessing = True
                mobjWipe.StopProcessing = gblnStopProcessing
                RefreshCaptions
                lblPathFile(0).Caption = "Cleaning work area ..."
                DoEvents
                
                Reset_Screen
                AlwaysOnTop frmMain, False
                                
           Case 2 ' about window
                frmMain.Hide
                frmAbout.DisplayAbout
           
           ' Shutdown this application
           Case Else
                If mblnWipeFreeSpace Then
                    RemoveWorkFolder
                End If
                
                gblnStopProcessing = True
                mobjWipe.StopProcessing = gblnStopProcessing
                DoEvents
                
                AlwaysOnTop frmMain, False
                WriteIniFile
                TerminateProgram
    End Select
    
End Sub

Private Sub RefreshCaptions()

    With frmMain
        .lblProgMsg(2).Caption = vbNullString
        .lblProgMsg(3).Caption = vbNullString
        .lblProgMsg(4).Caption = vbNullString
        .lblProgMsg(6).Caption = vbNullString
        .lblProgMsg(8).Caption = vbNullString
        .lblProgMsg(11).Caption = vbNullString
        .lblProgMsg(12).Caption = vbNullString
        .lblProgMsg(14).Caption = vbNullString
    End With

End Sub

Private Sub RemoveWorkFolder()

    Dim intIndex As Integer
    Dim strPath  As String
    
    CloseAllFiles
    
    For intIndex = 0 To UBound(mastrFixedDrives) - 1
        
        strPath = mastrFixedDrives(intIndex)
        strPath = QualifyPath(strPath) & TEMP_FOLDER
        
        If IsPathValid(strPath) Then
            mobjWipe.FolderCleanUp strPath, True
        End If
        
    Next intIndex
    
End Sub


Private Sub Reset_Screen()

    ' Make the GO button the topmost
    ' button and hide the STOP button
    With frmMain
        .cmdChoice(0).Enabled = True    ' Enable buttons
        .cmdChoice(1).Enabled = True
        .cmdChoice(2).Enabled = True
        .cmdChoice(3).Enabled = True
        ' Show buttons
        .cmdChoice(0).Visible = True     ' show go
        .cmdChoice(1).Visible = False    ' hide stop
    
        ' stop animation
        DoEvents
        If mblnAviReady Then
            .aniWork(0).Stop
            .aniWork(1).Stop
            .aniWork(2).Stop
            .aniWork(2).Visible = False
        End If
        DoEvents
        
    End With
    
    ReInitialize
    
    If mblnWipeFreeSpace Then
        
        mstrDrive = Right$(mobjWipe.WipePath, 3)
        
        If mblnAllFixedDrives Then
            lblPathFile(0).Caption = WIPE_FREESPACE_ALL
            lblPathFile(1).Caption = WIPE_FREESPACE_ALL
        Else
            lblPathFile(0).Caption = WIPE_FREESPACE_SPECIFIC & mstrDrive
            lblPathFile(1).Caption = WIPE_FREESPACE_SPECIFIC & mstrDrive
        End If
    End If
           
    If mblnUSB_Drive Then
        cboUSB_Click
    End If
        
End Sub

' ***************************************************************************
' Routine:       SaveRequiredData
'
' Description:   Captures and saves flags and data prior to performing the
'                requested wipe operation.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 02-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added reference to encryption algorithms
' ***************************************************************************
Private Function SaveRequiredData() As Boolean

    Dim strDrive    As String
    Dim strDrvType  As String
    Dim lngDrvType  As Long
    Dim objDiskInfo As cDiskInfo
    
    Const ROUTINE_NAME As String = "SaveRequiredData"
    
    On Error GoTo SaveRequiredData_Error
    
    Set objDiskInfo = New cDiskInfo  ' Instantiate class object
        
    ' start animation
    DoEvents
    If mblnAviReady Then
        aniWork(0).Play
        aniWork(1).Play
    End If
    
    mblnUSB_Drive = False    ' Preset to FALSE
    mblnFlashDrive = False

    strDrive = QualifyPath(Left$(mstrDrive, 2))             ' Drive letter and colon
    lngDrvType = GetDriveType(strDrive)                     ' Determine drive type
    strDrvType = objDiskInfo.GetDriveDescription(strDrive)  ' Partition description
        
    Select Case lngDrvType
           
           Case 2   ' Removable drive
                ' Test for drive description
                If StrComp("USB drive", Left$(strDrvType, 9), vbBinaryCompare) = 0 Then
                   
                    ' Test for availablility
                    If objDiskInfo.IsDeviceReady(strDrive) Then
                            
                        ' Capture file system type (FAT32, NTFS, etc)
                        objDiskInfo.GetVolumeInfo strDrive, mstrFileSys
                    
                        mblnUSB_Drive = True
                        mblnFlashDrive = True
                        mblnRemoveTopFolder = True
                    End If
                End If
       
           Case 3   ' Fixed drive
                ' Test for availablility
                If objDiskInfo.IsDeviceReady(strDrive) Then
                        
                    ' Capture file system type (FAT32, NTFS, etc)
                    objDiskInfo.GetVolumeInfo strDrive, mstrFileSys
                End If
    
           Case Else
                GoTo SaveRequiredData_CleanUp  ' Invalid drive type
    End Select
           
    If StrComp("NTFS", mstrFileSys, vbTextCompare) = 0 Then
        lblProgMsg(14).Caption = mstrFileSys
    Else
        lblProgMsg(14).Caption = mstrFileSys & "  (very slow)"
    End If
    
    WriteIniFile   ' update ini file
    
    ' Common properties
    With mobjWipe
        .StopProcessing = gblnStopProcessing
        .LogData = gblnLogData
        .LogEncryptParms = gblnLogEncryptParms
        .CurrentPattern = gstrLogPattern
        .USB_Drive = mblnUSB_Drive
        .FlashDrive = mblnFlashDrive
        .DoFolders = mblnDoSubfolders
        .DoSubFolders = mblnDoSubfolders
        .RemoveTopFolder = mblnRemoveTopFolder
        .KeepDirStructure = mblnKeepDirStructure
        .WipeFreeSpace = mblnWipeFreeSpace
        .AlternateMethod = mblnAlternateMethod
        .ProtectiveItems = PROTECTED_ITEMS
    End With
    
    If mblnWipeFreeSpace Then
    
        ' Wiping free space
        With mobjWipe
            .WipePath = mstrPathFile
            ' min of three passes to zero drives
            .Passes = IIf(glngPasses > 3, glngPasses, 3)
        End With
    
        ' capture the most recent path visited
        gstrPreviousPath = mstrDrive
        mcurFreeSpaceStart = objDiskInfo.GetDiskSpaceInfo(strDrive)
    
    Else
        ' Wiping data
        With mobjWipe
            .VerifyData = gblnVerifyData
            .ZeroLastWrite = gblnZeroLastWrite
            .DisplayMsgs = gblnDisplayVerifyMsgs
            .Passes = glngPasses
            .WipeMethod = glngWipeMethod
            
            ' Encryption only
            Select Case glngWipeMethod
                   Case 9:  .EncryptAlgo = eAlgo_Rijndael
                   Case 10: .EncryptAlgo = eAlgo_Blowfish
                   Case 11: .EncryptAlgo = eAlgo_Twofish
                   Case 12: .EncryptAlgo = eAlgo_ArcFour
            End Select
        End With
    
        ' capture the most recent path visited
        If IsPathValid(mastrFileList(0)) Then
            gstrPreviousPath = QualifyPath(mastrFileList(0))
        Else
            gstrPreviousPath = GetPath(mastrFileList(0))
        End If
        
    End If
        
    lblPathFile(0).Caption = mstrPathFile & EVAL_DR
    lblPathFile(1).Caption = lblPathFile(0).Caption
    lblProgMsg(2).Caption = "Current Drive Is " & mstrDrive
    
    ' keep a record of the most recently visited path
    gobjINIMgr.SaveOneKeyValue gstrINI, INI_DEFAULT, "PreviousPath", gstrPreviousPath
    SaveRequiredData = True  ' Good finish
    
SaveRequiredData_CleanUp:
    Set objDiskInfo = Nothing  ' Always free objects when not needed
    On Error GoTo 0
    Exit Function

SaveRequiredData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    FillDriveLetterComboBox
    SaveRequiredData = False
    Resume SaveRequiredData_CleanUp

End Function

' ***************************************************************************
' Routine:       Reset_frmMain
'
' Description:   Checks specific boolean flags and performs pertinent processes
'                based on these flags.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Sub Reset_frmMain()
           
    On Error GoTo Reset_frmMain_Error

    With frmMain
        If mblnWipeFreeSpace Then
            .lblPattern.Caption = FREE_SPACE_PATTERN
        Else
            .lblPattern.Caption = gstrLogPattern
        End If
        .Show
    End With

Reset_frmMain_CleanUp:
    On Error GoTo 0
    Exit Sub

Reset_frmMain_Error:
    ErrorMsg MODULE_NAME, "Reset_frmMain", Err.Description
    Resume Reset_frmMain_CleanUp
    
End Sub

' ***************************************************************************
' Routine:       NormalWipe
'
' Description:   Performs one of the default wipe operations as displayed
'                in the options table.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub NormalWipe()
                        
    ' start processing
    If mobjWipe.BeginProcessing(mastrFileList()) Then
        
        If Not gblnStopProcessing Then
            DisplayFinishMsg False
        End If
        
        fraProgress.Visible = False
        ReInitialize
                            
    End If
                
End Sub

' ***************************************************************************
' Routine:       CustomWipe
'
' Description:   Performs a custom defined wipe operation as defined by the
'                user.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub CustomWipe()
                        
    Dim intIndex       As Integer
    Dim intPosition    As Integer
    Dim strTmp         As String
    Dim astrTemp()     As String
    Dim astrPatterns() As String
    
    Erase astrTemp()      ' Always start with empty arrays
    Erase astrPatterns()
    
    gstrCustom = gstrCustom & "|"
    astrTemp() = Split(gstrCustom, "|")
    intPosition = InStr(1, astrTemp(3), "*")
    
    If intPosition = 0 Then
    
        ' add trailing comma to so we get the last value
        astrTemp(3) = astrTemp(3) & ","
        astrPatterns() = Split(astrTemp(3), ",")
            
        For intIndex = 0 To UBound(astrPatterns) - 1
             
            strTmp = astrPatterns(intIndex)
            intPosition = InStr(1, strTmp, "x")
            
            If intPosition > 0 Then
                ' convert pattern to numeric value
                strTmp = Mid$(strTmp, intPosition + 1)
                astrPatterns(intIndex) = CStr(Val("&H" & strTmp))
            Else
                astrPatterns(intIndex) = "999"     ' designate random data
            End If
        
        Next intIndex
        
    Else
        ReDim astrPatterns(3)
        strTmp = TrimStr(Mid$(astrTemp(3), 3, 3))   ' Get pattern to be duplicated
        
        ' convert pattern to numeric value
        astrPatterns(0) = IIf(strTmp = "999", strTmp, CStr(Val("&H" & strTmp)))
        astrPatterns(1) = TrimStr(Mid$(astrTemp(3), intPosition + 1))
        astrPatterns(2) = "Multiple"
    End If
    
    mobjWipe.WipePatterns = astrPatterns()
    
    Erase astrTemp()      ' Empty arrays when not needed
    Erase astrPatterns()
    
    ' Start processing
    If mobjWipe.BeginProcessing(mastrFileList()) Then
        
        If Not gblnStopProcessing Then
            DisplayFinishMsg False
        End If
        
        fraProgress.Visible = False
        ReInitialize
    End If
                                
End Sub

Private Sub DisplayFinishMsg(ByVal blnWipeFreeSpace As Boolean)

    Dim strMsg          As String
    Dim strDisplayFmt   As String
    Dim curFreeSpaceNow As Currency
    Dim objDrive        As Drive
    Dim objFSO          As Scripting.FileSystemObject
    
    If Len(Trim$(mstrDrive)) = 0 Then
        Exit Sub
    End If
    
    Set objFSO = New Scripting.FileSystemObject
    Set objDrive = objFSO.GetDrive(mstrDrive)
    
    strDisplayFmt = "!" & String$(20, "@")
    
    If mblnAllFixedDrives Then
        
        strMsg = "You have successfully completed wiping the free space on all local fixed drives."
        strMsg = strMsg & vbNewLine & vbNewLine
        strMsg = strMsg & "Security recommendations are:"
        strMsg = strMsg & vbNewLine & vbNewLine
        strMsg = strMsg & Space$(7) & "Step 1.  Stop all active applications."
        strMsg = strMsg & vbNewLine
        strMsg = strMsg & Space$(7) & "Step 2.  Defragment this PC."
        strMsg = strMsg & vbNewLine
        strMsg = strMsg & Space$(7) & "Step 3.  Shutdown and reboot."
    
    Else
    
        curFreeSpaceNow = objDrive.FreeSpace
        
        If blnWipeFreeSpace Then
            strMsg = "You have successfully completed wiping the free space on " & mstrDrive & " drive."
            strMsg = strMsg & vbNewLine & vbNewLine
            strMsg = strMsg & "Security recommendations are:"
            strMsg = strMsg & vbNewLine & vbNewLine
            strMsg = strMsg & Space$(7) & "Step 1.  Stop all active applications."
            strMsg = strMsg & vbNewLine
            strMsg = strMsg & Space$(7) & "Step 2.  Defragment this PC."
            strMsg = strMsg & vbNewLine
            strMsg = strMsg & Space$(7) & "Step 3.  Shutdown and reboot."
        Else
            strMsg = "You have successfully completed wiping data on " & mstrDrive & " drive."
            strMsg = strMsg & vbNewLine & vbNewLine
            strMsg = strMsg & "Security recommendations are:"
            strMsg = strMsg & vbNewLine & vbNewLine
            strMsg = strMsg & Space$(7) & "Step 1.  Stop all active applications."
            strMsg = strMsg & vbNewLine
            strMsg = strMsg & Space$(7) & "Step 2.  Wipe free space on this drive with at least 1 pass."
            strMsg = strMsg & vbNewLine
            strMsg = strMsg & Space$(7) & "Step 3.  Defragment this PC."
            strMsg = strMsg & vbNewLine
            strMsg = strMsg & Space$(7) & "Step 4.  Shutdown and reboot."
        End If
           
        strMsg = strMsg & vbNewLine & vbNewLine
        
    End If
    
    If gblnDisplayFinishMsg Then
        InfoMsg strMsg
    End If
    
    Set objFSO = Nothing
    Set objDrive = Nothing

End Sub

' ***************************************************************************
' Routine:       cmdFiles_Click
'
' Description:   Displays the file open dialog box.  Allows the user to select
'                one or more files to be wiped.
'
' Parameters:    None.
'
' Returns:       None.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub cmdFiles_Click()

    Dim strFiles As String   ' Hold list of files from dialog box
    Dim strDrive As String   ' Drive letter
    Dim objDrive As Drive
    Dim objFSO   As Scripting.FileSystemObject
    
    On Error GoTo Cancel_Pressed
    
    Set objFSO = New Scripting.FileSystemObject
            
    gblnStopProcessing = False       ' reset the STOP flag
    mobjWipe.StopProcessing = gblnStopProcessing
                
    strFiles = vbNullString
    lblPathFile(0).Caption = vbNullString
    lblPathFile(1).Caption = vbNullString
    Erase mastrFileList()
    
    ' always set to TRUE first time into the dialog box
    dlgCD.CancelError = True
    
    ' Prepare and show the File Open dialog box
    With dlgCD
        .Filter = "All files (*.*)|*.*"
        .FileName = vbNullString
        .MaxFileSize = 32767        ' allow enough memory for selected file names
        .InitDir = gstrPreviousPath
        .Flags = cdlOFNAllowMultiselect Or _
                 cdlOFNLongNames Or _
                 cdlOFNHideReadOnly Or _
                 cdlOFNExplorer
        .ShowOpen
    End With
    
    ' Save the list of files selected
    strFiles = TrimStr(dlgCD.FileName)
    
    ' if we have any files selected, then place them in an array
    If Len(strFiles) = 0 Then
        exit sub
    end if
        
    ' Save the list of files selected and append a null character
    ' to make sure all items are placed into the array
    strFiles = strFiles & Chr$(0)
    strDrive = QualifyPath(Left$(strFiles, 2))
    Set objDrive = objFSO.GetDrive(strDrive)
    
    If objDrive.DriveType = Removable Or _
       objDrive.DriveType = Fixed Then
       
        mstrDrive = UCase$(strDrive)
        mastrFileList() = Split(strFiles, Chr$(0), Len(strFiles))  ' Load the array
        
        ' see if multiple files were selected
        If UBound(mastrFileList) > 1 Then
            ' Show first file only
            mastrFileList(0) = QualifyPath(mastrFileList(0))
            lblPathFile(0).Caption = ShrinkToFit(mastrFileList(0) & mastrFileList(1), PATH_LEN) & "  ..."
        Else
            ' only one file selected
            lblPathFile(0).Caption = ShrinkToFit(mastrFileList(0), PATH_LEN)
        End If
        
        lblPathFile(1).Caption = lblPathFile(0).Caption
        mstrPathFile = lblPathFile(0).Caption
    Else
        InfoMsg WARNING_MSG
    End If
    
CleanUp:
    Set objFSO = Nothing
    Set objDrive = Nothing
    On Error GoTo 0
    Exit Sub
    
Cancel_Pressed:
    ' This usually means the user selected CANCEL on the dialog box
    Resume CleanUp
    
End Sub

' ***************************************************************************
' Routine:       cmdFolders_Click
'
' Description:   Displays the folder dialog box.  Allows the user to select
'                a folder to be wiped.  At this point, only the files in the
'                upper level folder will be targeted.  If there are no
'                subfolders or no files, then the folder will be deleted.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 04-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Update the amount of folder name to be displayed
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to disallow the user to wipe a drive
'              starting at the root if that drive contains the Windows
'              operating system.
'              Added title when browsing for a folder.
' ***************************************************************************
Private Sub cmdFolders_Click()
    
    Dim strMsg    As String
    Dim strDrive  As String
    Dim strFolder As String     ' holds the folder name selected
    Dim blnWinDrv As Boolean
    Dim intPos    As Integer
    Dim intCount  As Integer
    Dim objBrowse As cBrowse    ' class to display the folder dialog box
    Dim objDrive  As Drive
    Dim objFSO    As Scripting.FileSystemObject
    
    On Error GoTo Cancel_Pressed
    
    Set objFSO = New Scripting.FileSystemObject
    
    gblnStopProcessing = False       ' reset the STOP flag
    mobjWipe.StopProcessing = gblnStopProcessing
                
    intPos = 0
    intCount = 0
    strFolder = vbNullString
    lblPathFile(0).Caption = vbNullString
    lblPathFile(1).Caption = vbNullString
    Erase mastrFileList()
    
    Set objBrowse = New cBrowse   ' instantiate the class module
    strFolder = objBrowse.BrowseForFolder(frmMain, "Select folder to wipe")
    Set objBrowse = Nothing       ' Free class object from memory
    
    ' see if a folder was selected
    If Len(Trim$(strFolder)) > 0 Then
        
        strDrive = QualifyPath(Left$(strFolder, 2))
        mstrDrive = UCase$(strDrive)
        Set objDrive = objFSO.GetDrive(mstrDrive)
        
        ' is this a floppy or flash
        If objDrive.DriveType = Removable Then
           
            ReDim mastrFileList(1)          ' resize the array for for 1 element
            mastrFileList(0) = strFolder    ' add folder name to array
            lblPathFile(0).Caption = ShrinkToFit(strFolder, PATH_LEN)
            lblPathFile(1).Caption = lblPathFile(0).Caption
            
        ' Must be fixed disk
        ElseIf objDrive.DriveType = Fixed Then
               
            ' Further testing if this is the root of a drive
            If IsPathARoot(strFolder) Then
                
                ' See if this is the root of the drive which
                ' contains the Windows operating system
                blnWinDrv = IsWindowsFolder(mstrDrive)
                strFolder = UCase$(strFolder)
                
                ' If this is the root of the drive which
                ' contains the Windows operating system,
                ' then display a warning message.
                If blnWinDrv Then
                    
                    strMsg = "Drive " & mstrDrive
                    strMsg = strMsg & " is where your Windows operating system resides."
                    strMsg = strMsg & vbNewLine & vbNewLine
                    strMsg = strMsg & "You are not allowed to wipe this drive starting at the root."
                        
                    InfoMsg strMsg, "WARNING"

                    strFolder = vbNullString
                    mstrDrive = vbNullString
                    mstrPathFile = vbNullString
                    lblPathFile(0).Caption = vbNullString
                    lblPathFile(1).Caption = vbNullString
                    Erase mastrFileList()
                    GoTo cmdFolders_CleanUp
                
                Else
                    ' If the user has selected the root of
                    ' a fixed drive, display a warning message
                    strMsg = "You are about to permanently destroy all files "
                    strMsg = strMsg & "in the root directory of drive " & strFolder & " and possibly"
                    strMsg = strMsg & "the complete folder structure."
                    strMsg = strMsg & vbNewLine & vbNewLine
                    strMsg = strMsg & "Are you sure you want to do this?"
                    
                    If ResponseMsg(strMsg, , "WARNING") = vbYes Then
                        ReDim mastrFileList(1)          ' resize the array for for 1 element
                        mastrFileList(0) = strFolder    ' add folder name to array
                    Else
                        ' NO was selected
                        strFolder = vbNullString
                        mstrDrive = vbNullString
                        mstrPathFile = vbNullString
                        lblPathFile(0).Caption = vbNullString
                        lblPathFile(1).Caption = vbNullString
                        Erase mastrFileList()
                        GoTo cmdFolders_CleanUp
                    End If
                End If
            
            Else
                
                ' See if this folder is located
                ' within the Windows folder
                blnWinDrv = IsWindowsFolder(strFolder)
               
                If blnWinDrv Then
                    ' If the user has selected the Windows
                    ' folder or one of its subfolders.  If
                    ' so, display a warning message
                    strMsg = "You are about to permanently destroy some of the files "
                    strMsg = strMsg & "in the Windows folder."
                    strMsg = strMsg & vbNewLine & vbNewLine
                    strMsg = strMsg & "Are you sure you want to do this?"
                    
                    If ResponseMsg(strMsg, , "WARNING") = vbYes Then
                        ReDim mastrFileList(1)          ' resize the array for for 1 element
                        mastrFileList(0) = strFolder    ' add folder name to array
                    Else
                        ' NO was selected
                        strFolder = vbNullString
                        mstrDrive = vbNullString
                        mstrPathFile = vbNullString
                        lblPathFile(0).Caption = vbNullString
                        lblPathFile(1).Caption = vbNullString
                        Erase mastrFileList()
                        GoTo cmdFolders_CleanUp
                    End If
                Else
                    ReDim mastrFileList(1)          ' resize the array for for 1 element
                    mastrFileList(0) = strFolder    ' add folder name to array
                End If
            End If
            
            strFolder = UnQualifyPath(strFolder)
            mstrPathFile = strFolder
            
            ' if the top level folder checkbox is
            ' checked then update with folder name
            If mblnRemoveTopFolder Then
                
                intPos = InStr(1, strFolder, "\")
                
                If intPos > 0 Then
                    
                    intCount = intCount + 1
                    strFolder = Mid$(strFolder, intPos + 1)
                    
                    Do While intPos > 0
                        intPos = InStr(strFolder, "\")
                        
                        If intPos > 0 Then
                            intCount = intCount + 1
                            strFolder = Mid$(strFolder, intPos + 1)
                        End If
                    Loop
                End If
                
                If intCount > 1 Then
                    lblPath.Caption = ShrinkToFit(mstrDrive & "...\" & strFolder, PATH_LEN)
                Else
                    lblPath.Caption = ShrinkToFit(mstrDrive & strFolder, PATH_LEN)
                End If
            
            Else
            
                lblPath.Caption = ShrinkToFit(mstrPathFile, PATH_LEN)
            
            End If
    
            lblPathFile(0).Caption = ShrinkToFit(mstrPathFile, PATH_LEN)
            lblPathFile(1).Caption = lblPathFile(0).Caption
        
        Else
            InfoMsg WARNING_MSG
            
            strFolder = vbNullString
            mstrDrive = vbNullString
            mstrPathFile = vbNullString
            lblPathFile(0).Caption = vbNullString
            lblPathFile(1).Caption = vbNullString
            Erase mastrFileList()
            GoTo cmdFolders_CleanUp
        End If
    End If
    
cmdFolders_CleanUp:
    Set objBrowse = Nothing  ' Free class objects from memory
    Set objDrive = Nothing
    Set objFSO = Nothing
    On Error GoTo 0
    Exit Sub
    
Cancel_Pressed:
    ' Most likely the user selected
    ' CANCEL on the dialog box
    Resume cmdFolders_CleanUp
        
End Sub

' ***************************************************************************
' Routine:       cmdLogFile_Click
'
' Description:   If the log data checkbox has been checked and a log file
'                currently exist, then it will be displayed using the
'                default text editor.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub cmdLogFile_Click()

    gblnStopProcessing = False       ' reset the STOP flag
    mobjWipe.StopProcessing = gblnStopProcessing
                
    frmMain.Hide
    
    With frmLogFiles
        .Reset_frmLogFiles
        .Show
    End With
    
End Sub

Private Sub cmdOptions_Click()
    
    gblnStopProcessing = False  ' reset STOP flag
    mobjWipe.StopProcessing = gblnStopProcessing
                
    frmMain.Hide         ' Hide main form
    frmOptions.LoadGrid  ' Load and display options form
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub Form_Initialize()

    ' Make sure form is hidden
    ' during initial load
    frmMain.Hide
    DoEvents
        
End Sub

Private Sub Form_Load()
        
    Set mobjWipe = New kiWipe.cWipe   ' Instantiate dll object
    gblnStopProcessing = mobjWipe.StopProcessing
    
    If gblnStopProcessing Then
        TerminateProgram
        Exit Sub
    End If
    
    mobjWipe.StopProcessing = False
    
    Erase mastrAVI()          ' Always start with empty arrays
    Erase mastrFileList()
    Erase mastrFixedDrives()
    
    mlngEncryptAlgo = 0
    mblnAlternateMethod = False
    mblnAviReady = CreateAviFiles()
    
    DisableX frmMain         ' Disable "X" in upper right corner of form
    FillDriveLetterComboBox  ' default local fixed disks to "C:\"
    
    With frmMain
        .Caption = gstrVersion
        gobjKeyEdit.CenterCaption frmMain
        mstrCaption = .Caption
        
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                "No warranties or guarantees implied or intended."
        .cmdFiles.Picture = LoadResPicture("FIND", vbResBitmap)
        .cmdFolders.Picture = LoadResPicture("FIND", vbResBitmap)
        .Icon = LoadResPicture("MAIN", vbResIcon)
        .chkAlternateMethod.Value = vbUnchecked
        
        optFolders_Click 0              ' Default to remove top level folder
        .optFolders(0).Enabled = False
        .optFolders(1).Enabled = False
        .optFolders(2).Enabled = False
        
        optTarget_Click 0        ' Default to wiping files
        chkAlternateMethod_Click
        
        ' If AVI files successfully created
        ' then assign them to a control
        If mblnAviReady Then
            .aniWork(0).Open mastrAVI(0)  ' Scrolling monitors
            .aniWork(1).Open mastrAVI(0)
            .aniWork(2).Visible = False
            .aniWork(2).Enabled = False
        End If
        
        ' center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
    End With
    
    ReInitialize             ' initialize form controls
    Reset_Screen
    DoEvents
    frmMain.Hide
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' Verify the avi files have been stopped
    If mblnAviReady Then
        
        Dim intIdx As Integer
        
        On Error Resume Next
        mblnAviReady = False  ' reset flags
        
        For intIdx = 0 To aniWork.Count - 1
            aniWork(intIdx).Stop    ' stop playing AVI file
            aniWork(intIdx).Close   ' close AVI file
        Next intIdx
        
        CloseAllFiles
        DoEvents
        
        ' Delete AVI files
        For intIdx = 0 To UBound(mastrAVI) - 1
        
            ' if AVI file exist then delete it
            If IsPathValid(mastrAVI(intIdx)) Then
                Kill mastrAVI(intIdx)
            End If
                    
            DoEvents
        Next intIdx
                
        On Error GoTo 0  ' Reset set error trap
        
    End If
    
    Erase mastrAVI()           ' Always empty arrays when not needed
    Erase mastrFileList()
    Erase mastrFixedDrives()
    
    Set mobjWipe = Nothing     ' Free object from memory
    
End Sub

Private Sub Form_Resize()

    DoEvents
    With frmMain
        If .WindowState = vbMinimized Then
            If .fraMain.Enabled Then
                .Caption = PGM_NAME
            Else
                .Caption = CStr(mlngTotalProgress) & "% Completed"
            End If
        Else
            .Caption = mstrCaption
        End If
    End With
    DoEvents
    
End Sub



' ************************************************************************
' **                      Wipe events begin here                        **
' ************************************************************************
Private Sub mobjWipe_CountPasses(ByVal lngCurrentPass As Long, _
                                 ByVal lngMaxPasses As Long)
                              
    ' Tracks number of passes
    ' when overwriting a file
    
    frmMain.lblProgMsg(0).Caption = "Pass " & Format$(lngCurrentPass, "0") & _
                                    " of " & Format$(lngMaxPasses, "0")
    DoEvents

End Sub

Private Sub mobjWipe_CountTotals(ByVal strPathFile As String, _
                                 ByVal curFileCount As Currency, _
                                 ByVal curByteCount As Currency, _
                                 ByVal curFileSize As Currency, _
                                 ByVal blnCountOnly As Boolean, _
                                 ByVal blnTempFiles As Boolean)
    
    ' Tracks name of folder and file
    ' currently being processed
    
    DoEvents
    If blnCountOnly Then
        mcurFileCount = curFileCount
        mlngPrevOverall = 0
        mlngPrevCurrent = 0
    Else
        With frmMain
            .lblPathFile(1).Caption = ShrinkToFit(strPathFile, PATH_LEN)
            
            If blnTempFiles Then
                .lblProgMsg(3).Caption = Format$(curFileCount, "#,##0")
            Else
                .lblProgMsg(3).Caption = Format$(curFileCount, "#,##0") & " of " & _
                                         Format$(mcurFileCount, "#,##0")
            End If
            
            If mblnWipeFreeSpace Then
                Exit Sub
            End If
        
            DoEvents
            If curFileSize <> -1 Then
                .lblProgMsg(4).Caption = Format$(curByteCount, "#,##0") & " of " & _
                                         Format$(curFileSize, "#,##0")
            Else
                .lblProgMsg(4).Caption = "0 of 0"
            End If
            
        End With
    End If
    
    DoEvents
    
End Sub

Private Sub mobjWipe_CurrentPattern(ByVal strCurrentPattern As String)

    gstrLogPattern = strCurrentPattern
    frmMain.lblPattern.Caption = "Pattern:  " & strCurrentPattern
    
End Sub

Private Sub mobjWipe_CurrentProgress(ByVal curByteCount As Currency, _
                                     ByVal curMaxAmount As Currency)
    
    ' wiping only files
    DoEvents
    frmMain.lblProgMsg(4).Caption = Format$(curByteCount, "#,##0") & "  of  " & _
                                    Format$(curMaxAmount, "#,##0")
    
    If curByteCount = 0 Then
        mlngPrevCurrent = 0
    End If
    DoEvents
    
End Sub

Private Sub mobjWipe_ElapsedTime(ByVal strElapsedTime As String)
    
    DoEvents
    frmMain.lblProgMsg(6).Caption = strElapsedTime
    
End Sub

Private Sub mobjWipe_FileProgress(ByVal lngFileProgress As Long)

    DoEvents
    ProgressBar picProgressBar(0), lngFileProgress, vbBlue
    mlngPrevCurrent = lngFileProgress
    DoEvents

End Sub

Private Sub mobjWipe_OverallProgress(ByVal curByteCount As Currency, _
                                     ByVal curMaxAmount As Currency)

    DoEvents
    If curByteCount >= MB_1 Then
        frmMain.lblProgMsg(8).Caption = DisplayNumber(curByteCount, 3) & "  ( " & _
                                        DisplayNumber(curByteCount) & "  of  " & _
                                        DisplayNumber(curMaxAmount) & " )"
    Else
        frmMain.lblProgMsg(8).Caption = Format$(curByteCount, "#,##0") & "  ( " & _
                                        DisplayNumber(curByteCount) & "  of  " & _
                                        DisplayNumber(curMaxAmount) & " )"
    End If
        
    If curByteCount = 0 Then
        mlngPrevOverall = 0
    End If
    DoEvents
    
End Sub
    
Private Sub mobjWipe_TimeRemaining(ByVal strTimeRemaining As String, _
                                   ByVal strTransferRate As String)

    DoEvents
    With frmMain
        .lblProgMsg(11).Caption = strTimeRemaining
        .lblProgMsg(12).Caption = strTransferRate
    End With
    DoEvents

End Sub

Private Sub mobjWipe_TotalProgress(ByVal lngTotalProgress As Long)
    
    mlngTotalProgress = lngTotalProgress

    If frmMain.WindowState = vbMinimized Then
        Form_Resize
        Exit Sub
    End If
    
    ' prevent flickering
    DoEvents
    If lngTotalProgress > mlngPrevOverall Then
        ProgressBar picProgressBar(1), lngTotalProgress, vbRed
        mlngPrevOverall = lngTotalProgress
    End If
    DoEvents

End Sub

Private Sub mobjWipe_UpdateLogData(ByVal strLogRecord As String)
    
    UpdateLogFile strLogRecord
    strLogRecord = vbNullString
    DoEvents
    
End Sub

Private Sub mobjWipe_WaitMsg(ByVal strMsg As String, _
                             ByVal lngColor As Long)
    
    DoEvents
    With frmMain.lblPathFile(1)
        
        If StrComp(Left$(strMsg, 8), "REMOVING", vbTextCompare) = 0 Then
            .Alignment = vbCenter
        Else
            .Alignment = vbLeftJustify
        End If
        
        .BackColor = lngColor
        .Caption = strMsg
    End With
    
End Sub

Private Sub mobjWipe_WipingFiles(ByVal blnStarted As Boolean)

    DoEvents
    With frmMain.aniWork(2)
        If blnStarted Then
            .Enabled = True
            .Open mastrAVI(3)  ' Files AVI
            .Visible = True
            .Play
        Else
            .Stop
            .Visible = False
            .Enabled = False
        End If
    End With
    DoEvents
            
End Sub

Private Sub mobjWipe_WipingFreespace(ByVal blnStarted As Boolean)

    DoEvents
    With frmMain.aniWork(2)
        If blnStarted Then
            .Enabled = True
            .Visible = True
            .Open mastrAVI(2)  ' WipeFree AVI
            .Play
        Else
            .Stop
            .Visible = False
            .Enabled = False
        End If
    End With
    DoEvents
            
End Sub

Private Sub mobjWipe_WipingFolders(ByVal blnStarted As Boolean)

    DoEvents
    With frmMain.aniWork(2)
        If blnStarted Then
            .Enabled = True
            .Open mastrAVI(1)  ' Folders AVI
            .Visible = True
            .Play
        Else
            .Stop
            .Visible = False
            .Enabled = False
        End If
    End With
    DoEvents
            
End Sub
' ************************************************************************
' **                      Wipe events end here                          **
' ************************************************************************


Private Sub optFolders_Click(Index As Integer)

    Select Case Index
           Case 0    ' Remove top level folder
                optFolders(0).Value = True
                optFolders(1).Value = False
                optFolders(2).Value = False
                mblnRemoveTopFolder = True
                mblnKeepDirStructure = False
           
           Case 1    ' Keep top level folder
                optFolders(0).Value = False
                optFolders(1).Value = True
                optFolders(2).Value = False
                mblnRemoveTopFolder = False
                mblnKeepDirStructure = False
           
           Case 2    ' Keep folder structure
                optFolders(0).Value = False
                optFolders(1).Value = False
                optFolders(2).Value = True
                mblnRemoveTopFolder = False
                mblnKeepDirStructure = True
    End Select
            
    If Len(Trim$(mstrPathFile)) > 0 Then
        lblPath.Caption = ShrinkToFit(mstrPathFile, PATH_LEN)
    End If
    
End Sub

Private Sub optTarget_Click(Index As Integer)

    ' Option to determine if files,
    ' folders, or free space are to
    ' be wiped.
    
    mobjWipe.TypeTarget = Index + 1
    mblnWipeFreeSpace = False
    mblnDoSubfolders = False
    mstrPathFile = vbNullString
    
    With frmMain
        .lblPathFile(0).Caption = vbNullString
        .lblPathFile(1).Caption = vbNullString
        .lblPath.Caption = vbNullString
        .lblProgMsg(1).Caption = " Number of items"
        .chkAlternateMethod.Enabled = False
            
        Select Case Index
               Case 0  ' wipe files
                    .lblPattern.Caption = gstrLogPattern
                    .cmdFiles.Enabled = True
                    .cmdFolders.Enabled = False
                    .cboDrive.Enabled = False
                    .cboUSB.Enabled = False
                    .optFolders(0).Enabled = False
                    .optFolders(1).Enabled = False
                    .optFolders(2).Enabled = False
                                    
               Case 1  ' wipe folders
                    mblnDoSubfolders = True
                    .lblPattern.Caption = gstrLogPattern
                    .cmdFiles.Enabled = False
                    .cmdFolders.Enabled = True
                    .cboDrive.Enabled = False
                    .cboUSB.Enabled = False
                    .optFolders(0).Enabled = True
                    .optFolders(1).Enabled = True
                    .optFolders(2).Enabled = True
                    optFolders_Click 0
                    
               Case 2   ' wipe USB Drive
                    FillDriveLetterComboBox
                    .cmdFiles.Enabled = False
                    .cmdFolders.Enabled = False
                    .cboUSB.Enabled = True
                    .optFolders(0).Enabled = False
                    .optFolders(1).Enabled = False
                    .optFolders(2).Enabled = False
                    cboUSB_Click
        
               Case 3   ' wipe free space
                    FillDriveLetterComboBox
                    mblnWipeFreeSpace = True
                    .chkAlternateMethod.Enabled = True
                    .cmdFiles.Enabled = False
                    .cmdFolders.Enabled = False
                    .cboUSB.Enabled = False
                    .cboDrive.Enabled = True
                    .optFolders(0).Enabled = False
                    .optFolders(1).Enabled = False
                    .optFolders(2).Enabled = False
                    cboDrive_Click
        End Select
    End With
    
    DoEvents
    Reset_frmMain
    
End Sub

Private Sub FillDriveLetterComboBox()

    Dim lngIndex     As Long
    Dim lngCount     As Long
    Dim lngDrvType   As Long
    Dim strDrive     As String
    Dim strDrvType   As String
    Dim astrDrives() As String
    Dim objDiskInfo  As cDiskInfo
    
    Set objDiskInfo = New cDiskInfo
    Erase astrDrives()
    Erase mastrFixedDrives()
    
    cboDrive.Enabled = True
    cboUSB.Enabled = True
    cboDrive.Clear
    cboUSB.Clear
    lngCount = 0
    ReDim mastrFixedDrives(26)
    
    ' Capture list of available drive letters
    astrDrives() = objDiskInfo.GetDriveLetters()
    
    ' load combo boxes
    With cboDrive
        For lngIndex = 0 To UBound(astrDrives) - 1
            
            lngDrvType = GetDriveType(astrDrives(lngIndex))         ' Determine drive type
                
            Select Case lngDrvType
                   
                   Case 2   ' Removable drive
                        strDrive = QualifyPath(Left$(astrDrives(lngIndex), 2))  ' Drive letter and colon
                        strDrvType = objDiskInfo.GetDriveDescription(strDrive)  ' Partition description
                        
                        ' Test for drive description
                        If StrComp("USB drive", Left$(strDrvType, 9), vbBinaryCompare) = 0 Then
                           
                            ' Test for availablility
                            If objDiskInfo.IsDeviceReady(strDrive) Then
                                    
                                ' Add to combo boxes
                                .AddItem astrDrives(lngIndex) & "  Removable"
                                cboUSB.AddItem astrDrives(lngIndex) & "  Removable"
                            End If
                        End If
               
                   Case 3   ' Fixed drive
                        ' Test for availablility
                        If objDiskInfo.IsDeviceReady(astrDrives(lngIndex)) Then
                                                        
                            ' Capture file system type (FAT32, NTFS, etc)
                            objDiskInfo.GetVolumeInfo astrDrives(lngIndex), mstrFileSys
                            
                            ' Add to combo box
                            .AddItem astrDrives(lngIndex) & "  Fixed - " & mstrFileSys
                            
                            ' Update array
                            mastrFixedDrives(lngCount) = astrDrives(lngIndex)
                            lngCount = lngCount + 1
                        End If
            End Select
                   
        Next lngIndex
        
        .AddItem "All Fixed Drives"
        .ListIndex = 0
        .Enabled = False
    End With
    
    ReDim Preserve mastrFixedDrives(lngCount)  ' resize to number of available drives
    
    If cboUSB.ListCount > 0 Then
        cboUSB.ListIndex = 0
    End If
    
    mstrFileSys = vbNullString  ' Empty variable
    cboUSB.Enabled = False      ' Disable USB combo box
    Erase astrDrives()          ' Always empty arrays when not needed
    Set objDiskInfo = Nothing   ' Free object when not needed
    
End Sub

Private Function CreateAviFiles() As Boolean

    Dim hFile      As Long
    Dim abytData() As Byte
        
    Const ROUTINE_NAME As String = "CreateAviFiles"

    On Error GoTo CreateAviFiles_Error

    CreateAviFiles = False   ' Preset to FALSE
    Erase mastrAVI()         ' Always start with an empty array
    ReDim mastrAVI(4)        ' Size AVI file name array
    
    ' Load file names
    mastrAVI(0) = QualifyPath(App.Path) & "Scroll.avi"
    mastrAVI(1) = QualifyPath(App.Path) & "Folders.avi"
    mastrAVI(2) = QualifyPath(App.Path) & "WipeFree.avi"
    mastrAVI(3) = QualifyPath(App.Path) & "Files.avi"

    Erase abytData()   ' Always start with an empty array
    abytData() = LoadResData("SCROLL", "CUSTOM")
    
    hFile = FreeFile
    Open mastrAVI(0) For Binary Access Write As #hFile
    Put #hFile, , abytData()
    Close #hFile
    
    Erase abytData()   ' Always start with an empty array
    abytData() = LoadResData("FOLDERS", "CUSTOM")
    
    hFile = FreeFile
    Open mastrAVI(1) For Binary Access Write As #hFile
    Put #hFile, , abytData()
    Close #hFile
    
    Erase abytData()   ' Always start with an empty array
    abytData() = LoadResData("WIPEFREE", "CUSTOM")
    
    hFile = FreeFile
    Open mastrAVI(2) For Binary Access Write As #hFile
    Put #hFile, , abytData()
    Close #hFile
    
    Erase abytData()   ' Always start with an empty array
    abytData() = LoadResData("FILES", "CUSTOM")
    
    hFile = FreeFile
    Open mastrAVI(3) For Binary Access Write As #hFile
    Put #hFile, , abytData()
    Close #hFile
    
    CreateAviFiles = True
    
CreateAviFiles_CleanUp:
    Erase abytData()   ' Always empty arrays when not needed
    On Error GoTo 0
    Exit Function

CreateAviFiles_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Erase mastrAVI()   ' Always empty arrays when not needed
    Resume CreateAviFiles_CleanUp
    
End Function

Private Sub ReInitialize()
     
    DoEvents
    ' initialize controls on the main form
    With frmMain
        .Caption = gstrVersion
        .fraProgress.Visible = False
        .fraMain.Enabled = True
        .fraMain.Visible = True  ' show input areas
        
        .lblPathFile(0).Caption = vbNullString
        .lblPathFile(1).Caption = vbNullString
        
        .lblProgMsg(0).Caption = "Pass 0/0"
        .lblProgMsg(2).Caption = "Current progress"
        .lblProgMsg(3).Caption = vbNullString
        .lblProgMsg(4).Caption = vbNullString
        .lblProgMsg(6).Caption = vbNullString
        .lblProgMsg(7).Caption = "Overall progress"
        .lblProgMsg(8).Caption = vbNullString
        .lblProgMsg(12).Caption = vbNullString
        
        .lblPath.Caption = vbNullString
    End With
        
    ResetProgressBar
    
    DoEvents
    mstrPathFile = vbNullString
    mcurFileCount = 0
    mcurByteCount = 0
    mlngPrevOverall = 0
    mlngPrevCurrent = 0
    mlngTotalProgress = 0
    Erase mastrFileList()

End Sub

Private Sub ResetProgressBar()

    ' Resets progressbar to zero
    ' with all white background
    ProgressBar picProgressBar(0), 0, vbWhite
    ProgressBar picProgressBar(1), 0, vbWhite
        
End Sub

' ***************************************************************************
' Routine:       ProgessBar
'
' Description:   Fill a picturebox as if it were a horizontal progress bar.
'
' Parameters:    objProgBar - name of picture box control
'                lngPercent - Current percentage value
'                lngForeColor - Optional-The progression color. Default = Black.
'                           can use standard VB colors or long Integer
'                           values representing a color.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 14-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
' 05-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' ***************************************************************************
Private Sub ProgressBar(ByRef objProgBar As PictureBox, _
                        ByVal lngPercent As Long, _
               Optional ByVal lngForeColor As Long = vbBlue)

    Dim strPercent As String
    
    Const MAX_PERCENT As Long = 100
    
    ' Called by ResetProgressBar() routine
    ' to reinitialize progress bar properties.
    ' If forecolor is white then progressbar
    ' is being reset to a starting position.
    If lngForeColor = vbWhite Then
        
        With objProgBar
            .AutoRedraw = True      ' Required to prevent flicker
            .BackColor = &HFFFFFF   ' White
            .DrawMode = 10          ' Not Xor Pen
            .FillStyle = 0          ' Solid fill
            .FontName = "Arial"     ' Name of font
            .FontSize = 11          ' Font point size
            .FontBold = True        ' Font is bold.  Easier to see.
            Exit Sub                ' Exit this routine
        End With
    
    End If
        
    ' If no progress then leave
    If lngPercent < 1 Then
        Exit Sub
    End If
    
    ' Verify flood display has not exceeded 100%
    If lngPercent <= MAX_PERCENT Then

        With objProgBar
        
            ' Error trap in case code attempts to set
            ' scalewidth greater than the max allowable
            If lngPercent > .ScaleWidth Then
                lngPercent = .ScaleWidth
            End If
               
            .Cls                        ' Empty picture box
            .ForeColor = lngForeColor   ' Reset forecolor
         
            ' set picture box ScaleWidth equal to maximum percentage
            .ScaleWidth = MAX_PERCENT
            
            ' format percent into a displayable value (ex: 25%)
            strPercent = Format$(CLng((lngPercent / .ScaleWidth) * 100)) & "%"
            
            ' Calculate X and Y coordinates within
            ' picture box and and center data
            .CurrentX = (.ScaleWidth - .TextWidth(strPercent)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(strPercent)) \ 2
                
            objProgBar.Print strPercent   ' print percentage string in picture box
            
            ' Print flood bar up to new percent position in picture box
            objProgBar.Line (0, 0)-(lngPercent, .ScaleHeight), .ForeColor, BF
        
        End With
                
        DoEvents   ' allow flood to complete drawing
    
    End If

End Sub

