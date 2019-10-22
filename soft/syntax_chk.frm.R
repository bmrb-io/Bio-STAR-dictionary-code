VERSION 5.00
Begin VB.Form syntax_check 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   Caption         =   "Syntax Check"
   ClientHeight    =   7155
   ClientLeft      =   285
   ClientTop       =   570
   ClientWidth     =   10845
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   10845
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   6960
      TabIndex        =   43
      Text            =   "Text22"
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   6960
      TabIndex        =   42
      Text            =   "Text21"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   5880
      TabIndex        =   41
      Text            =   "Text20"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   5880
      TabIndex        =   40
      Text            =   "Text19"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   8040
      TabIndex        =   37
      Text            =   "Text17"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   5640
      TabIndex        =   36
      Text            =   "Text18"
      Top             =   2520
      Width           =   5055
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   4080
      TabIndex        =   32
      Text            =   "Text16"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   8640
      TabIndex        =   20
      Text            =   "Text15"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   8640
      TabIndex        =   19
      Text            =   "Text14"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   8040
      TabIndex        =   18
      Text            =   "Text13"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2280
      TabIndex        =   17
      Text            =   "Text12"
      Top             =   4680
      Width           =   3495
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Text            =   "Text11"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Text            =   "Text10"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Text            =   "Text9"
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Text            =   "Text8"
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Text            =   "Text7"
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1080
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "ADIT super group ID"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   39
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "ADIT category group ID"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   38
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Error type"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   35
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Item count"
      Height          =   255
      Left            =   2280
      TabIndex        =   34
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Value count"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   33
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Loop count error"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Length"
      Height          =   255
      Left            =   7080
      TabIndex        =   30
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Length"
      Height          =   255
      Left            =   7080
      TabIndex        =   29
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Category No."
      Height          =   255
      Left            =   6240
      TabIndex        =   28
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Duplicate category"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Empty field"
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "No. of misformed tags"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Tag"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Category"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "No. of tags"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Saveframe"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "No. of saveframes"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "NMR-STAR Schema File Syntax and Completeness Check"
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
      Left            =   1800
      TabIndex        =   14
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "File path"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "No. of tables"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "File name"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Line error"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "syntax_check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
