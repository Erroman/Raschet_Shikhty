VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Подготовка задания"
   ClientHeight    =   6396
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7812
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6396
   ScaleWidth      =   7812
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Поиск задания в базе данных"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2040
      TabIndex        =   3
      Top             =   3720
      Width           =   4092
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Подготовка задания из расчета"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   4932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.2
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   1
      Top             =   4800
      Width           =   3132
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Поиск расчета в базе данных"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1200
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   1800
      Width           =   5652
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  Окно подготовки задания"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   25.8
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   6132
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
UserForm1.Show
End Sub

Private Sub Command2_Click()
'Unload Me
Unload UserForm1
Unload UserForm2
Unload UserForm3
Unload UserForm4
Unload UserForm5
Unload Form1

End Sub

Private Sub Command3_Click()
UserForm4.Show
End Sub

Private Sub Command4_Click()
UserForm5.Show
End Sub
