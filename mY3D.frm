VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5070
   Icon            =   "mY3D.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Text            =   " CHOOSE"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Text            =   " OR"
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF0000&
      Caption         =   " Night"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FFFF&
      Caption         =   " Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Menu fmenu 
      Caption         =   "File"
      Begin VB.Menu bigmenu 
         Caption         =   "Big Screen"
      End
      Begin VB.Menu smallmenu 
         Caption         =   "Small Screen"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Option1_Click()
Option1.Value = True
Form1.Hide
Form2.Show
End Sub

Private Sub Option2_Click()
Option2.Value = True
Form1.Hide
Form2.Show
End Sub
