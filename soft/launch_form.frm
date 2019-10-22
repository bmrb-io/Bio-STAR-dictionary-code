VERSION 5.00
Begin VB.Form launch_form 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Launch the NMR-STAR dictionary analysis"
   ClientHeight    =   8796
   ClientLeft      =   6708
   ClientTop       =   1248
   ClientWidth     =   7812
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   8796
   ScaleWidth      =   7812
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Launch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Application Operation Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6132
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
      Begin VB.CheckBox Check9 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Check tags in a test file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   600
         TabIndex        =   12
         Top             =   5400
         Width           =   4332
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Write a SG stub dictionary (.csv file)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   4920
         Width           =   5055
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Update from the NMR-STAR template"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   4320
         Width           =   5055
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Generate ADIT files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   3720
         Width           =   5055
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Update keys in the Excel file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   3120
         Width           =   5055
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Update from Excel file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2520
         Width           =   5055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Write interface items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1920
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Write NMR-STAR dictionary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Check Excel file integrity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "NMR-STAR Dictionary Processing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "launch_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
run_prog = 1
run_option = 1
Close
go_yahoo
launch_form.Show
End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub

Private Sub Check1_Click()
launch_form.Check1.Value = 1
End Sub
