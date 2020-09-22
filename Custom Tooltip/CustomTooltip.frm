VERSION 5.00
Begin VB.Form CustomTooltip 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3000
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "CustomTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create a form with 3 Command Buttons and a Text Box

Option Explicit

Private Sub Form_Load()
    ' Add the custom tooltip to the controls
    AddCustomToolTip Command1, "This is an example" & vbCrLf & "of a Custom ToolTip Window" & _
        vbCrLf & "With multiline text", Me
    AddCustomToolTip Command2, "This is another" & vbCrLf & "custom ToolTip", Me
    AddCustomToolTip Command3, "Hi! I'm a Tip", Me
    AddCustomToolTip Text1, "TextBox ToolTip", Me
End Sub


