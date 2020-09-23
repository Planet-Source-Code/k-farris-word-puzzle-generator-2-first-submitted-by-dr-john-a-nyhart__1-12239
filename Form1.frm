VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Puzzle"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   7680
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
rtf.SelPrint Printer.hDC
End Sub

Private Sub Form_Load()


rtf.SelAlignment = rtfCenter

rtf.Text = rtf.Text + title + vbCrLf + vbCrLf


rtf.Text = rtf.Text + frmMain.txtWord.Text
rtf.Text = rtf.Text + vbCrLf + vbCrLf + vbCrLf
For x = 0 To frmMain.lstInputWord.ListCount - 1
    rtf.Text = rtf.Text + Str(x + 1) + ". " + frmMain.lstInputWord.List(x) + "  "
    Next x
    
End Sub
