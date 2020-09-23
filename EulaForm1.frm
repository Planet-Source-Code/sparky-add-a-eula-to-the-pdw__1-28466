VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmEula 
   Caption         =   "Eula"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   Icon            =   "EulaForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2850
      Left            =   45
      TabIndex        =   5
      Top             =   1080
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   5027
      _Version        =   393217
      BackColor       =   12632256
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"EulaForm1.frx":0442
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   4500
      TabIndex        =   0
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next >"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3165
      TabIndex        =   4
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Yes, I &accept all of the terms of the preceeding license agreement"
      Height          =   240
      Left            =   195
      TabIndex        =   3
      Top             =   4020
      Width           =   5325
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   5715
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   5715
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Label Label2 
      Caption         =   "End User License Agreement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   75
      Width           =   4530
   End
   Begin VB.Label Label1 
      Caption         =   $"EulaForm1.frx":0517
      Height          =   675
      Left            =   105
      TabIndex        =   1
      Top             =   360
      Width           =   5610
   End
End
Attribute VB_Name = "FrmEula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CancelExitFlg As Boolean

Private Sub Check1_Click()
CmdNext.Enabled = Not CmdNext.Enabled
End Sub

Private Sub CmdCancel_Click()
ExitSetup Me, gintRET_EXIT

End Sub

Private Sub CmdNext_Click()
Unload Me
End Sub

Private Sub Form_Load()
    'Load Resource item 103 (a text file) into the textbox
    Dim sFileName As String
    Dim iFileNum  As Integer
    Dim sText     As String
    Dim Margins As Integer
    
    CancelExitFlg = False
    
    If GetTempFile("", "~rs", 0, sFileName) Then
        If Not SaveResItemToDisk(101, "Custom", sFileName) Then
           RichTextBox1.LoadFile sFileName
           Margins = 100
           With RichTextBox1
              .SelStart = 1
              .SelLength = Len(RichTextBox1.Text)
              .SelIndent = Margins
              .SelStart = 0
           End With
            'Delete the temp file
            Kill sFileName
        Else
            MsgBox "Unable to save resource item to disk!", vbCritical
            'Show/Hide Controls
        End If
    Else
        MsgBox "Unable to get temp file name!", vbCritical
        'Show/Hide Controls
    End If



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If CancelExitFlg = True Then
'   ExitSetup Me, gintRET_EXIT
'   CancelExitFlg = False
'   Exit Sub
'End If

HandleFormQueryUnload UnloadMode, Cancel, Me

End Sub

Private Sub RichTextBox1_Click()
Check1.SetFocus
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Check1.SetFocus
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Check1.SetFocus
End Sub


