VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmStatus 
   Caption         =   "Cabral Status:"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   4575
   End
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4048
      _Version        =   327681
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmStatus.frx":0CCA
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'intro text
    txtStatus.SelBold = True
    txtStatus.SelColor = RGB(140, 0, 0)
    txtStatus.SelText = "Cabral IRC Client"
    txtStatus.SelColor = vbBlack
    txtStatus.SelText = " by Jaime Cabral" & vbCrLf
    txtStatus.SelBold = False
    txtStatus.SelColor = RGB(0, 140, 0)
    txtStatus.SelText = "important message"
    txtStatus.SelColor = vbBlack
    txtStatus.SelText = ": Want to help? E-mail me.  Help wanted for colors in channel and private query windows, dns, DCC, and other small features." & vbCrLf & "Please rebort bugs to cabral@n-link.com.  Have comments, additions, or suggestions?  E-mail me." & vbCrLf
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtStatus.Move txtStatus.Left, txtStatus.Top, frmStatus.Width - 150, frmStatus.Height - 700
    txtSend.Move txtStatus.Left, txtStatus.Height - 10, frmStatus.Width - 150, 300
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo xError
    If KeyCode = 13 Then
        mdiMain.tcp.SendData txtSend & vbCrLf
        txtStatus.SelText = "> " & txtSend & vbCrLf
        txtSend = ""
    End If
xError:
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
End Sub


Private Sub txtStatus_Change()
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub
