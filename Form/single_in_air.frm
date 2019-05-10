VERSION 5.00
Object = "{5E347DF1-C0CC-11D9-84B3-9100B9C7DC45}#6.0#0"; "BZ_Frame2.ocx"
Object = "{C0F23CC2-0EE2-4CD5-975D-EDA99765C6F1}#1.0#0"; "Light_Transparent_Button.ocx"
Begin VB.Form single_in_air_trefoil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Method of Installation"
   ClientHeight    =   5775
   ClientLeft      =   1680
   ClientTop       =   1995
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   11640
   Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_XP_Frame1 
      Height          =   5805
      Left            =   -30
      TabIndex        =   5
      Top             =   -15
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10239
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
         Height          =   1830
         Left            =   180
         Picture         =   "single_in_air.frx":0000
         ScaleHeight     =   1770
         ScaleWidth      =   3165
         TabIndex        =   11
         Top             =   1815
         Width           =   3225
         Begin VB.PictureBox Pic_tic_1 
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
            Left            =   -15
            Picture         =   "single_in_air.frx":092F
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   12
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
         Height          =   1740
         Left            =   4260
         Picture         =   "single_in_air.frx":0DA7
         ScaleHeight     =   1680
         ScaleWidth      =   3165
         TabIndex        =   9
         Top             =   1815
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
            Picture         =   "single_in_air.frx":182A
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   10
            Top             =   0
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
         Height          =   1860
         Left            =   8235
         Picture         =   "single_in_air.frx":1CA2
         ScaleHeight     =   1800
         ScaleWidth      =   3165
         TabIndex        =   7
         Top             =   1815
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
            Picture         =   "single_in_air.frx":26A8
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
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
         ItemData        =   "single_in_air.frx":2B20
         Left            =   1703
         List            =   "single_in_air.frx":2B2D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   4590
         Width           =   1305
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
         ItemData        =   "single_in_air.frx":2B3A
         Left            =   8618
         List            =   "single_in_air.frx":2B47
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4590
         Width           =   1305
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame IA 
         Height          =   345
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   1335
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   609
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
         HeaderText      =   "                      Perforated Trays"
         FooterText      =   ""
         HeaderText      =   "                      Perforated Trays"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_troed 
         Height          =   390
         Left            =   7343
         TabIndex        =   13
         Top             =   4155
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
         HeaderText      =   "         Number of three-phase Circuits"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "         Number of three-phase Circuits"
         FooterText      =   ""
         HeaderPictureSize=   12
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_gt 
         Height          =   390
         Left            =   413
         TabIndex        =   14
         Top             =   4155
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
         Left            =   4538
         TabIndex        =   4
         Top             =   4485
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
         Top             =   4485
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
         Left            =   5453
         TabIndex        =   2
         Top             =   4485
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
      Begin Babak_Zawari_Frame.BZ_XP_Frame IA 
         Height          =   345
         Index           =   1
         Left            =   8220
         TabIndex        =   15
         Top             =   1335
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   609
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
         HeaderText      =   "          Ladder Supports Cleats"
         FooterText      =   ""
         HeaderText      =   "          Ladder Supports Cleats"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame IA 
         Height          =   345
         Index           =   2
         Left            =   4230
         TabIndex        =   16
         Top             =   1335
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   609
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
         HeaderText      =   "      Vertical Perforated Trays"
         FooterText      =   ""
         HeaderText      =   "      Vertical Perforated Trays"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
   End
End
Attribute VB_Name = "single_in_air_trefoil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
      i = 3
      Case 2
      i = 4
      Case 3
      i = 5
    End Select
           
         
If Pic_tic_1.Visible = True Or Pic_tic_2.Visible = True Or Pic_tic_3.Visible = True Then
dd = 2
Unload Me
Else
MsgBox ("Please select one of Pictures")
End If
End Sub

Private Sub Pic_1_Click()
j = 12
Pic_tic_1.Visible = True
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
End Sub

Private Sub Pic_2_Click()
j = 16
Pic_tic_1.Visible = False
Pic_tic_2.Visible = True
Pic_tic_3.Visible = False
End Sub

Private Sub Pic_3_Click()
j = 20
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = True
End Sub
