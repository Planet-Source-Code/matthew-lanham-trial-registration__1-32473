VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Main"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Unregister"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MsgBox"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "You have used 0 days of your 30 day Trial"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'This code apart form the CheckFile function was
'Made by ZeOn ( zeon@xoolo.net)(http://www.xoolo.net)
'if you have any problems with this feel
'free to email me or add me to msn
'--------------------------------------
Dim filefound As Boolean
Dim abc As String
Dim abd As Integer
Dim registered As Boolean
Private Sub Command1_Click()
Form2.Visible = True
End Sub
Public Function CheckFile(InFileName As String) As Boolean
On Error GoTo ErrHandler
CheckFile = False
If Dir(InFileName) <> "" Then
If (GetAttr(InFileName) And vbDirectory) = 0 Then
CheckFile = True
filefound = True
Else
filefound = False
Exit Function
End If
End If
ErrHandler:
End Function
'--------------------------------------
'this is the function to check wether a file exists
Public Function Checkreg(InFileName As String) As Boolean
On Error GoTo ErrHandler
registered = False
'--------------------------------------
' if the directory = ""
If Dir(InFileName) <> "" Then
'--------------------------------------
' if the directory has the file in it
If (GetAttr(InFileName) And vbDirectory) = 0 Then
'--------------------------------------
'set the boolean to true
registered = True
'--------------------------------------
Else
'--------------------------------------
'set the boolean to false
registered = False
'--------------------------------------
Exit Function
End If
End If
ErrHandler:
End Function
Private Sub Command2_Click()
Form1.Visible = False
End Sub

Private Sub Command3_Click()
'--------------------------------------
'set the labels caption
MsgBox "This is an example of how to make a trial program"
End Sub
'--------------------------------------
Private Sub Command4_Click()
'--------------------------------------
'this code Deletes the file so you can re register
Kill ("C:\Registered.txt")
'--------------------------------------
'Displays a msgbox
MsgBox "Reload to See Changes", vbInformation, "Reload"
'--------------------------------------
End Sub

Private Sub Form_Load()
'--------------------------------------
'Displays a msgbox
MsgBox "Register Example by ZeOn", vbInformation, "ZeOn"
'--------------------------------------
'Calls the CheckFile Function
'to check wether the file exists
Checkreg ("C:\Registered.txt")
'--------------------------------------
'Calls the CheckFile Function
'to check wether the file exists
CheckFile ("C:\ab.txt")
'--------------------------------------
'if the user is registered registered will be True
If registered = False Then
'--------------------------------------
'Set the command button so we can see it
Command1.Visible = True
'--------------------------------------
'If the File we asked to search is found
If filefound = True Then
'open the file to get text form it
Open "C:\ab.txt" For Input As #1
' set the string name to store the text
Input #1, abc
'clsoe the file
Close #1
'--------------------------------------
'This takes one data from another where "d" = Days "w" = weeks "m" = months
abd = DateDiff("d", abc, Date)
'--------------------------------------
' if the 2 dats taken form each other = or are more than 30
If abd <= 30 Then
'set the caption of the label
Label1.Caption = "You have used " & abd & " Days Of your 30 Day Trial"
'if it dosent <= 30 then
Else
'change the labels caption
Label1.Caption = "Your 30 day trial has run out"
'--------------------------------------
'if the users registration has expired then set the form to disabled
Form1.Enabled = False
'--------------------------------------
'open the registration form
Form2.Visible = True
' make the command button enabled for registering
Command1.Visible = True
'--------------------------------------
End If
Else
'--------------------------------------
' if the file dosent exist make it
' and put todays date in it
Open "C:\ab.txt" For Output As #1
Print #1, Date
Close #1
'--------------------------------------
'get the text form the file
Open "C:\ab.txt" For Input As #1
Input #1, abc
Close #1
'--------------------------------------
End If
Else
'--------------------------------------
'set the labels caption
Label1.Caption = "Registered Copy"
'--------------------------------------
'make the command visable
Command4.Visible = True
End If
End Sub
