VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox UserName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label3 
      Height          =   975
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "E-mail Adress:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form1.Signin UserName.Text, Password.Text
End Sub

Private Sub Command2_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Label3.Caption = "Please Fill in you Login Details"
End Sub
