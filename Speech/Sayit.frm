VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Say it"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4950
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3120
      Top             =   3360
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tips and tricks!"
      Height          =   375
      Left            =   2680
      TabIndex        =   12
      Top             =   3340
      Width           =   2250
   End
   Begin VB.Frame Frame1 
      Caption         =   "Loading Voices..."
      Height          =   2020
      Left            =   2450
      TabIndex        =   9
      Top             =   0
      Width           =   2499
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.Timer Timer5 
         Interval        =   1000
         Left            =   945
         Top             =   600
      End
      Begin VB.Label Selected 
         Alignment       =   2  'Center
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   1540
         Width           =   2145
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Convert Text to .wav here!"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3340
      Width           =   2685
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   375
      Left            =   730
      TabIndex        =   7
      Top             =   2955
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2955
      Width           =   735
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   1080
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   750
      Left            =   3960
      Top             =   2880
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   3540
      Max             =   250
      Min             =   50
      TabIndex        =   3
      Top             =   3000
      Value           =   150
      Width           =   1410
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   2880
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   885
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Sayit.frx":0000
      Top             =   2040
      Width           =   4950
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Say it"
      Height          =   375
      Left            =   1470
      TabIndex        =   1
      Top             =   2955
      Width           =   1215
   End
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   795
      Left            =   120
      OleObjectBlob   =   "Sayit.frx":0007
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Sayit.frx":002B
      ForeColor       =   &H0000FFFF&
      Height          =   885
      Left            =   600
      TabIndex        =   14
      Top             =   4320
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tips"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2040
      TabIndex        =   13
      Top             =   3840
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   1080
      X2              =   240
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1440
      X2              =   2280
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      Height          =   135
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   960
      Shape           =   2  'Oval
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   375
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   375
      Left            =   360
      Shape           =   2  'Oval
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Shape           =   2  'Oval
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "50"
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   3045
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
      Height          =   195
      Left            =   2715
      TabIndex        =   4
      Top             =   3045
      Width           =   510
   End
   Begin VB.Shape Shape8 
      BorderWidth     =   3
      Height          =   60
      Left            =   120
      Shape           =   2  'Oval
      Top             =   405
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BorderWidth     =   3
      Height          =   60
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   405
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DropDown As Boolean
Dim DropUp As Boolean
Sub LoadText(Lst As TextBox, file As String)
On Error GoTo error
Dim mystr As String
Open file For Input As #1
Do While Not EOF(1)
            Line Input #1, a$
            texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
        Loop
        Lst = texto$
Close #1
Exit Sub
error:
X = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub
Private Sub Command1_Click()
CommonDialog1.ShowSave
If CommonDialog1.CancelError Then
MsgBox "Nothing Selected!"
Exit Sub
Else
Call SaveText(Text1, CommonDialog1.FileName)
End If
End Sub
Private Sub Command2_Click()
CommonDialog1.ShowOpen
If CommonDialog1.CancelError = True Then
MsgBox "Nothing selected"
Exit Sub
Else
Call LoadText(Text1, CommonDialog1.FileName)
End If
End Sub

Sub SaveText(Lst As TextBox, file As String)
On Error GoTo error
Dim mystr As String
Open file For Output As #1
Print #1, Lst
Close 1
Exit Sub
error:
X = MsgBox("There has been a error!", vbOKOnly, "Error")
End Sub
Private Sub Command3_Click()
If Command3.Caption = "Say it" Then
  If Text1.Text = "" Then
   TextToSpeech1.Speak "Box empty"
   Exit Sub
  Else
   TextToSpeech1.Speak Text1.Text
  End If
Else
  TextToSpeech1.StopSpeaking
End If
End Sub

Private Sub Command4_Click()
Shell App.Path & "\" & "TextToWav.exe", vbNormalFocus
End Sub

Private Sub Command5_Click()
Timer6.Enabled = True
End Sub

Private Sub Form_Load()
CommonDialog1.Filter = "Text Files (*.log;*.ini;*.txt)|*.log;*.ini;*.txt"
  TextToSpeech1.Speaker (1)
  TextToSpeech1.Speed = 150
  Me.Top = Screen.Width / 6
  Me.Left = Screen.Height / 3
  TextToSpeech1.MouthHeight = 0
  TextToSpeech1.TeethLowerVisible = 0
  TextToSpeech1.TeethUpperVisible = 0
  TextToSpeech1.Speak "Welcome to say it!"

End Sub

Private Sub Form_Unload(Cancel As Integer)
TextToSpeech1.Speed = 150
 TextToSpeech1.Speak "Bye bye,"
 MsgBox "Bye bye!", vbApplicationModal, ""
End Sub

Private Sub HScroll1_Change()
TextToSpeech1.Speed = HScroll1.Value
Label2.Caption = HScroll1.Value
Label2.Visible = True
Timer2.Enabled = True
End Sub

Private Sub HScroll1_Scroll()
TextToSpeech1.Speed = HScroll1.Value
Label2.Caption = HScroll1.Value
Label2.Visible = True
Timer2.Enabled = True
End Sub


Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BorderStyle = 1
Shape4.Visible = False
Shape5.Visible = False
Shape7.Visible = False
End Sub

Private Sub Label6_Click()
TextToSpeech1.Speak "Stop peeking my nose!"
End Sub

Private Sub Label7_Click()
TextToSpeech1.Speak "Rororororororororororororo!"
End Sub

Private Sub List1_DblClick()
Selected.Caption = "Loading"
On Error Resume Next
Dim a, b
a = List1.ListIndex
b = a + 1
TextToSpeech1.Select b
Selected.Caption = TextToSpeech1.MfgName(List1.ListIndex + 1)
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If tru = False Then
   If KeyCode = vbKeyF1 Then
      If Text1.Text = "" Then
      TextToSpeech1.Speak "Box empty"
      Exit Sub
     Else
      TextToSpeech1.Speak Text1.Text
      End If
   End If
 Else
  If KeyCode = vbKeyF1 Then
        If Text1.Text = "" Then
        Genie.Speak "Box empty"
        Exit Sub
     Else
      Genie.Speak Text1.Text
  End If
End If
End If
End Sub

Private Sub TextToSpeech1_ClickIn(ByVal X As Long, ByVal Y As Long)
TextToSpeech1.Speak "Whaat!"
End Sub

Private Sub Timer1_Timer()
If TextToSpeech1.IsSpeaking Then
Command3.Caption = "Stop"
Else
Command3.Caption = "Say it"
End If
End Sub

Private Sub Timer2_Timer()
Label2.Visible = False
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Shape1.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Line2.Y2 = 240
Line1.Y1 = 240
Line2.Y1 = 240
Line1.Y2 = 240
Shape9.Visible = True
Shape8.Visible = True
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Shape1.Visible = True
Shape3.Visible = True
Shape4.Visible = True
Shape5.Visible = True
Shape6.Visible = True
Shape7.Visible = True
Line2.Y2 = 120
Line2.Y1 = 120
Line1.Y1 = 120
Line1.Y2 = 120
Timer4.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer5_Timer()
Dim NumVoices As Integer
Dim Count As Integer
If TextToSpeech1.initialized <> 1 Then
    MsgBox "Speech engine failed to load", vbExclamation, "Error"
    Unload Main
    End
End If
NumVoices = TextToSpeech1.CountEngines
For Count = 1 To NumVoices
    List1.AddItem TextToSpeech1.ModeName(Count), Count - 1
    List1.ListIndex = 0
Next
Selected.Caption = TextToSpeech1.MfgName(List1.ListIndex + 1)
Frame1.Caption = "Available voices:"
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
If DropUp = False Then
  Form1.Height = Form1.Height + 50
   If Form1.Height > 6040 Then
    Timer6.Enabled = False
    DropUp = True
    Command5.Caption = "Close tips and tricks"
   End If
Else
   Form1.Height = Form1.Height - 50
    If Form1.Height < 4100 Then
     Timer6.Enabled = False
     Form1.Height = 4095
     DropUp = False
     Command5.Caption = "Tips and tricks"
    End If
End If
End Sub
