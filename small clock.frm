VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SMALL-CLOCK"
   ClientHeight    =   1290
   ClientLeft      =   15540
   ClientTop       =   1065
   ClientWidth     =   4605
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "small clock.frx":0000
   ScaleHeight     =   1290
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Index           =   0
      Left            =   4080
      TabIndex        =   7
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[1.0.3v]"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   210
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1005
      Index           =   5
      Left            =   2685
      TabIndex        =   5
      ToolTipText     =   "Double Click to change Mode of clock."
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1005
      Index           =   4
      Left            =   1245
      TabIndex        =   4
      ToolTipText     =   "Double Click to change Mode of clock."
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Index           =   3
      Left            =   4125
      TabIndex        =   3
      ToolTipText     =   "Double Click to change Mode of clock."
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1005
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Double Click to change Mode of clock."
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1005
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Double Click to change Mode of clock."
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1005
      Index           =   0
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Double Click to change Mode of clock."
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hh, mm, ss As Integer
Dim res As Integer

Private Sub Command1_Click()

End Sub

Private Sub Label1_DblClick(Index As Integer)
Unload Me
Load Form1
Form1.Show
End Sub


Private Sub Label3_Click(Index As Integer)
Shape1.BackColor = &HC0& ''red
Shape1.BorderColor = &HC0&


 If Index = 0 Then
   CreateObject("SAPI.SpVoice").Speak "Now, The Time is " & hh Mod 12 & " " & mm & " " & Format(Now, "ampm")
 ElseIf Index = 1 Then
   CreateObject("SAPI.SpVoice").Speak "Version 1.0.3, Developer, Ashutosh Mishra."
 End If
 
Shape1.BackColor = &HFF00&      ''green
End Sub

Private Sub Timer1_Timer()
Label1(0).Caption = Format(Now, "hh")
        hh = Format(Now, "hh")
Label1(1).Caption = Format(Now, "nn")
        mm = Format(Now, "n")
Label1(2).Caption = Format(Now, "ss")
        ss = Format(Now, "s")
Label1(3).Caption = Format(Now, "ampm")

Me.Caption = Format(Now, "dddd, dd mmmm yy")

    If (mm = 0) And (ss = 0) Then
       
       CreateObject("SAPI.SpVoice").Speak "Now, The Time is " & hh & " " & Format(Now, "ampm")
       
    End If


End Sub
