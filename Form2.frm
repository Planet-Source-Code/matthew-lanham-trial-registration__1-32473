VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Register"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form2"
   ScaleHeight     =   990
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Verify"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   5
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Please Enter Your Serial Number"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim serial As String
serial = Text1.Text & Text2.Text & Text3.Text
If serial = "123451234512345" Then
MsgBox "Thank You For Registering"
Open "C:\Registered.txt" For Output As #1
Close #1
Form1.Command1.Visible = False
Form2.Visible = False
Form1.Label1.Caption = "Registered Copy"
MsgBox "Reload this program to see changes", vbInformation, "Registered"
Else
MsgBox "Incorrect Try Again"
End If
End Sub

Private Sub Command2_Click()
Form2.Visible = False
End Sub

