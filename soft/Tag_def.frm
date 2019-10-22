VERSION 5.00
Begin VB.Form Tag_def 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Tag definitions, prompts, and examples"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "Find tag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8280
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   5880
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   4680
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2520
      Width           =   6375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "Previous tag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "Next tag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7800
      TabIndex        =   10
      Text            =   "Text7"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Edit tag definitions, prompts, and examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1200
      TabIndex        =   21
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Mandatory status"
      Height          =   255
      Left            =   6840
      TabIndex        =   20
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Data type"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Example(s)"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Prompt"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Definition"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tag"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Category"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Saveframe"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Tag_def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
