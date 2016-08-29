VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CLOCK [ADDTEAM]"
   ClientHeight    =   4695
   ClientLeft      =   15600
   ClientTop       =   840
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Digital-7"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "clock1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "clock1.frx":324A
   ScaleHeight     =   4695
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "&SWITCH"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "ADVANCE"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command3 
         BackColor       =   &H00404040&
         Caption         =   "&HELP"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404040&
         Caption         =   "&DONE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   720
         Left            =   1200
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "clock1.frx":3B43
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Digital-7"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   1560
         TabIndex        =   10
         Text            =   "10"
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "ON"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Ur Msg :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FOR EACH          TH MINs"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REMINDER :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "&ADVANCE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   10
      X1              =   0
      X2              =   4800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      X1              =   1560
      X2              =   4800
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "day"
      BeginProperty Font 
         Name            =   "Balls on the rampage"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   465
      Index           =   4
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "month"
      BeginProperty Font 
         Name            =   "Balls on the rampage"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   465
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dd"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1485
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1080
      Index           =   1
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1200
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   -120
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STATUS As Boolean
Dim MN As Integer
Dim CHK As Integer
Dim MMIN As Integer
Dim SSEC As Integer
Dim MMSG As String


Private Sub Command1_Click()

If Command1.Caption = "&ADVANCE" Then
   Frame1.Visible = True
   Form1.Width = 4905
   Form1.Height = 5130
   Command1.Caption = "&HIDE"
Else
   Frame1.Visible = False
   Form1.Width = 4905
   Form1.Height = 2970
   Command1.Left = 3840
   Command1.Top = 960
   Command1.Caption = "&ADVANCE"
End If

End Sub

Private Sub Command2_Click()
    
     If Val(Text1.Text) <> 0 Then
        MN = Text1.Text
     Else
        MN = 10
     End If
     
     If Text2.Text = "" Or Text2.Text = "Text2" Then
        MMSG = "PLEASE CHEACK YOUR TASK EVENT."
     Else
        MMSG = Text2.Text
     End If
  
  Frame1.Visible = False
   Form1.Width = 4905
   Form1.Height = 2970
   Command1.Left = 3840
   Command1.Top = 960
   Command1.Caption = "&ADVANCE"
  
End Sub

Private Sub Command3_Click()
Me.Enabled = False
Form2.Show
End Sub

Private Sub Command4_Click()
Unload Me
Load Form3
Form3.Show
End Sub

Private Sub Form_Load()
Form1.Width = 4905
Form1.Height = 2970
Command1.Left = 3840
Command1.Top = 960
Command1.Caption = "&ADVANCE"
STATUS = True
CHK = 0
MN = 10
MMSG = "PLEASE CHEACK YOUR TASK EVENT."
End Sub

Private Sub Option1_Click(Index As Integer)

Text1.Visible = True
Text2.Visible = True
Label3.Visible = True
Label4.Visible = True
Command2.Visible = True
Option1(0).Value = True
Option2.Value = False
STATUS = True
MN = Text1.Text

End Sub

Private Sub Option1_GotFocus(Index As Integer)
Option1(0).Value = True
Option2.Value = False
STATUS = True
MN = Text1.Text

End Sub

Private Sub Option2_Click()

Text1.Visible = False
Text2.Visible = False
Label3.Visible = False
Label4.Visible = False
Command2.Visible = False

Option1(0).Value = False
Option2.Value = True
STATUS = False

End Sub

Private Sub Text1_GotFocus()
    MN = Text1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
     If Val(Text1.Text) <> 0 Then
      MN = Text1.Text
     Else
      MN = 10
     End If
  Else
     MN = 10
     Text1.Text = 10
     MsgBox "ONLY NUMBERS ARE ALLOWED, PLEASE PROVIDE NUMBERS.", vbCritical, "INVAILID INPUT"
  End If
End Sub

Private Sub Text1_LostFocus()
   MN = Text1.Text
End Sub

Private Sub Text2_Change()
MMSG = Text2.Text
End Sub

Private Sub Text2_LostFocus()
  If Text2.Text = "" Or Text2.Text = "Text2" Then
     MMSG = "PLEASE CHEACK YOUR TASK EVENT."
  Else
     MMSG = Text2.Text
  End If
End Sub

Private Sub Timer1_Timer()
Label1(0).Caption = Format(Now, "hh:mm:ss")
Label1(1).Caption = Format(Now, "ampm")
Label1(2).Caption = Format(Now, "DD")
Label1(3).Caption = Format(Now, "mmmm, yy")
Label1(4).Caption = Format(Now, "dddd")
 
MMIN = Format(Now, "n")
SSEC = Format(Now, "s")

If Option1(0).Value = True Then
    MMSG = Text2.Text
    If MMIN Mod MN = 0 And SSEC = 0 Then
      MsgBox MMSG, vbInformation, "REMINDER"
    End If
End If

End Sub
