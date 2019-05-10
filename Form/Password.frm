VERSION 5.00
Object = "{5E347DF1-C0CC-11D9-84B3-9100B9C7DC45}#6.0#0"; "BZ_Frame2.ocx"
Object = "{C0F23CC2-0EE2-4CD5-975D-EDA99765C6F1}#1.0#0"; "Light_Transparent_Button.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3225
   ClientLeft      =   5760
   ClientTop       =   2610
   ClientWidth     =   4155
   ControlBox      =   0   'False
   Icon            =   "Password.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4155
   Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_XP_Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5741
      BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty footerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ContainerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderText      =   ""
      FooterText      =   ""
      Appearence      =   6
      HeaderText      =   ""
      FooterText      =   ""
      ZOrderOnFocus   =   0   'False
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1755
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
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1177
         Width           =   2055
      End
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   645
         Width           =   2055
      End
      Begin Light_ButtonTP.LSButtonT LSButtonT_cancel 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin Light_ButtonTP.LSButtonT LSButtonT_change 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Change"
      End
      Begin VB.Label L_new2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   8
         Top             =   1830
         Width           =   1275
      End
      Begin VB.Label L_new1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   7
         Top             =   1290
         Width           =   1275
      End
      Begin VB.Label L_old 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   6
         Top             =   735
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LSButtonT_cancel_Click()
Unload Me
End Sub

Private Sub LSButtonT_change_Click()
If ((Text1.Text = pass) And (Text2.Text = Text3.Text)) Then
MsgBox ("Password changed succesfully ")

Set ws_pass = wb.Worksheets("sheet1")
ws_pass.Cells(1, 1) = Text3.Text
wb.Save
Unload Me
ElseIf Text1.Text <> pass Then
MsgBox ("Old Password is Uncorrect!")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Else
MsgBox ("New Password is Uncorrect!")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End If
End Sub

