VERSION 5.00
Object = "{C0F23CC2-0EE2-4CD5-975D-EDA99765C6F1}#1.0#0"; "Light_Transparent_Button.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mghan Wire & Cable Co."
   ClientHeight    =   5010
   ClientLeft      =   1215
   ClientTop       =   1560
   ClientWidth     =   12435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   12435
   Begin VB.PictureBox Pic_moghan 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5070
      Left            =   -30
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5010
      ScaleWidth      =   12540
      TabIndex        =   6
      Top             =   0
      Width           =   12600
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   6420
         TabIndex        =   0
         Top             =   330
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   6420
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   930
         Width           =   2055
      End
      Begin Light_ButtonTP.LSButtonT LSButtonT_input 
         Height          =   375
         Index           =   0
         Left            =   9165
         TabIndex        =   2
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Input"
      End
      Begin Light_ButtonTP.LSButtonT LSButtonT_exit 
         Height          =   375
         Index           =   1
         Left            =   9165
         TabIndex        =   4
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exit"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5085
         TabIndex        =   5
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5085
         TabIndex        =   3
         Top             =   360
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Music.Show

   Set xlApp = New excel.Application
   Set wb = xlApp.Workbooks.Open(App.Path & "\Excel\tables.xls")

End Sub

Private Sub LSButtonT_exit_Click(Index As Integer)
Z = MsgBox("  Are You Sure?", vbYesNo, "Exit")
If Z = 6 Then
    wb.Close
    xlApp.Quit
    Unload Me
    Unload Music
End If
End Sub

Private Sub LSButtonT_input_Click(Index As Integer)
Set ws_pass = wb.Worksheets("sheet1")
pass = ws_pass.Cells(1, 1)
If ((Text1.Text = "moghan") And (Text2.Text = pass)) Then
Form2.Show
Unload Me
Else
MsgBox ("Username or Password is Uncorrect! ")
Text1.Text = ""
Text2.Text = ""
End If
End Sub




