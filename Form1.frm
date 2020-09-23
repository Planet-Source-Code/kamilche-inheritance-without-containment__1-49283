VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   6225
      Left            =   210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   165
      Width           =   7785
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Display(ByVal s As String)
    ' Display text in the text window
    txtOutput.SelStart = Len(txtOutput.Text)
    txtOutput.SelText = s & vbCrLf
End Sub

Private Sub Form_Resize()
    ' Resize interior controls to fit
    On Error Resume Next
   txtOutput.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


Private Sub mnuFileExit_Click()
    ' Quit
    Unload Me
End Sub
