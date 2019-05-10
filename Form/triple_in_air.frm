VERSION 5.00
Object = "{5E347DF1-C0CC-11D9-84B3-9100B9C7DC45}#6.0#0"; "BZ_Frame2.ocx"
Object = "{C0F23CC2-0EE2-4CD5-975D-EDA99765C6F1}#1.0#0"; "Light_Transparent_Button.ocx"
Begin VB.Form triple_in_air 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Method of Installation"
   ClientHeight    =   6720
   ClientLeft      =   2085
   ClientTop       =   1800
   ClientWidth     =   11655
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11655
   Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_XP_Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11880
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
      HeaderBackColor =   16711935
      Appearence      =   6
      HeaderText      =   ""
      FooterText      =   ""
      ZOrderOnFocus   =   0   'False
      Begin VB.PictureBox Pic_6 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   7957
         Picture         =   "triple_in_air.frx":0000
         ScaleHeight     =   1665
         ScaleWidth      =   3165
         TabIndex        =   21
         Top             =   3135
         Width           =   3225
         Begin VB.PictureBox Pic_tic_6 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   0
            Picture         =   "triple_in_air.frx":09BE
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   22
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.ComboBox Combo_cable 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "triple_in_air.frx":0E36
         Left            =   8618
         List            =   "triple_in_air.frx":0E4C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   5880
         Width           =   1305
      End
      Begin VB.ComboBox Combo_tray 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "triple_in_air.frx":0E62
         Left            =   1703
         List            =   "triple_in_air.frx":0E6F
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   5880
         Width           =   1305
      End
      Begin VB.PictureBox Pic_4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   4207
         Picture         =   "triple_in_air.frx":0E7C
         ScaleHeight     =   1875
         ScaleWidth      =   3135
         TabIndex        =   14
         Top             =   3135
         Width           =   3195
         Begin VB.PictureBox Pic_tic_4 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   0
            Picture         =   "triple_in_air.frx":17F5
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   472
         Picture         =   "triple_in_air.frx":1C6D
         ScaleHeight     =   1650
         ScaleWidth      =   3165
         TabIndex        =   12
         Top             =   3135
         Width           =   3225
         Begin VB.PictureBox Pic_tic_2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   0
            Picture         =   "triple_in_air.frx":2568
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_5 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   7939
         Picture         =   "triple_in_air.frx":29E0
         ScaleHeight     =   1725
         ScaleWidth      =   3135
         TabIndex        =   10
         Top             =   1200
         Width           =   3195
         Begin VB.PictureBox Pic_tic_5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   0
            Picture         =   "triple_in_air.frx":353D
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_1 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   521
         Picture         =   "triple_in_air.frx":39B5
         ScaleHeight     =   1590
         ScaleWidth      =   3165
         TabIndex        =   8
         Top             =   1200
         Width           =   3225
         Begin VB.PictureBox Pic_tic_1 
            BackColor       =   &H8000000B&
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   0
            Picture         =   "triple_in_air.frx":4416
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   9
            Top             =   -45
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_3 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   4230
         Picture         =   "triple_in_air.frx":488E
         ScaleHeight     =   1635
         ScaleWidth      =   3165
         TabIndex        =   6
         Top             =   1200
         Width           =   3225
         Begin VB.PictureBox Pic_tic_3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   0
            Picture         =   "triple_in_air.frx":5229
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame IA 
         Height          =   390
         Index           =   0
         Left            =   7762
         TabIndex        =   16
         Top             =   570
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   688
         BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty footerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ContainerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderText      =   "Cables on Ladder support Cleats"
         FooterText      =   ""
         HeaderText      =   "Cables on Ladder support Cleats"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame ISWD 
         Height          =   390
         Left            =   4012
         TabIndex        =   17
         Top             =   570
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   688
         BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty footerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ContainerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderText      =   "Cables on Vert.  Perforated Trays"
         FooterText      =   ""
         HeaderText      =   "Cables on Vert.  Perforated Trays"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame bdig 
         Height          =   390
         Index           =   0
         Left            =   262
         TabIndex        =   18
         Top             =   570
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   688
         BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty footerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ContainerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderText      =   "    Cables on Perforated Trays"
         FooterText      =   ""
         HeaderText      =   "    Cables on Perforated Trays"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_troed 
         Height          =   390
         Left            =   7343
         TabIndex        =   19
         Top             =   5445
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   688
         BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty footerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ContainerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderText      =   "                                  Number of Cables"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "                                  Number of Cables"
         FooterText      =   ""
         HeaderPictureSize=   12
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_gt 
         Height          =   390
         Left            =   413
         TabIndex        =   20
         Top             =   5445
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   688
         BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty footerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ContainerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderText      =   "                                  Number of Trays"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "                                  Number of Trays"
         FooterText      =   ""
         HeaderPictureSize=   14
         ZOrderOnFocus   =   0   'False
      End
      Begin Light_ButtonTP.LSButtonT LSButton_back 
         Height          =   495
         Left            =   4545
         TabIndex        =   4
         Top             =   5775
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Back"
      End
      Begin Light_ButtonTP.LSButtonT LSButton_reset 
         Height          =   495
         Left            =   6390
         TabIndex        =   3
         Top             =   5775
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Reset"
      End
      Begin Light_ButtonTP.LSButtonT LSButtonT_ok 
         Height          =   495
         Left            =   5467
         TabIndex        =   2
         Top             =   5775
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ok"
      End
   End
End
Attribute VB_Name = "triple_in_air"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Combo_tray.Text = 1
Combo_cable.Text = 1

End Sub

Private Sub LSButton_back_Click()
dd = 1
Unload Me

End Sub

Private Sub LSButton_reset_Click()
Combo_tray.Text = 1
Combo_cable.Text = 1

Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
End Sub

Private Sub LSButtonT_ok_Click()

    x = Combo_tray.Text
    Select Case x
      Case 1
        h = 0
      Case 2
        h = 1
      Case 3
        h = 2
    End Select
    
    x = Combo_cable.Text
    Select Case x
      Case 1
      i = 4
      Case 2
      i = 5
      Case 3
      i = 6
      Case 4
      i = 7
      Case 6
      i = 8
      Case 9
      i = 9
    End Select
           
         
If Pic_tic_1.Visible = True Or Pic_tic_2.Visible = True Or Pic_tic_3.Visible = True Or Pic_tic_4.Visible = True Or Pic_tic_5.Visible = True Or Pic_tic_6.Visible = True Then
dd = 2
Unload Me
Else
MsgBox ("Please select one of Pictures")
End If





End Sub

Private Sub Pic_1_Click()
j = 4
Pic_tic_1.Visible = True
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False

End Sub

Private Sub Pic_2_Click()
j = 8
Pic_tic_1.Visible = False
Pic_tic_2.Visible = True
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
End Sub

Private Sub Pic_3_Click()
j = 12
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = True
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
End Sub

Private Sub Pic_4_Click()
j = 16
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = True
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
End Sub

Private Sub Pic_5_Click()
j = 20
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = True
Pic_tic_6.Visible = False
End Sub

Private Sub Pic_6_Click()
j = 24
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = True
End Sub

