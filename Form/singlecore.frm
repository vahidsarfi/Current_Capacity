VERSION 5.00
Object = "{5E347DF1-C0CC-11D9-84B3-9100B9C7DC45}#6.0#0"; "BZ_Frame2.ocx"
Object = "{C0F23CC2-0EE2-4CD5-975D-EDA99765C6F1}#1.0#0"; "Light_Transparent_Button.ocx"
Begin VB.Form singlecore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Single Core"
   ClientHeight    =   6720
   ClientLeft      =   2085
   ClientTop       =   1800
   ClientWidth     =   11640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11640
   Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_XP_Frame1 
      Height          =   6735
      Left            =   -30
      TabIndex        =   10
      Top             =   -15
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
      Begin VB.ComboBox Combo_troed 
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
         ItemData        =   "singlecore.frx":0000
         Left            =   8618
         List            =   "singlecore.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5880
         Width           =   1305
      End
      Begin VB.ComboBox Combo_tros 
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
         ItemData        =   "singlecore.frx":0010
         Left            =   8618
         List            =   "singlecore.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4995
         Width           =   1305
      End
      Begin VB.ComboBox Combo_dol 
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
         ItemData        =   "singlecore.frx":0052
         Left            =   8633
         List            =   "singlecore.frx":0074
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4095
         Width           =   1305
      End
      Begin VB.ComboBox Combo_gt 
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
         ItemData        =   "singlecore.frx":00A6
         Left            =   1703
         List            =   "singlecore.frx":00C5
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   5880
         Width           =   1305
      End
      Begin VB.ComboBox Combo_aat 
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
         ItemData        =   "singlecore.frx":00ED
         Left            =   1703
         List            =   "singlecore.frx":010C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4995
         Width           =   1305
      End
      Begin VB.ComboBox Combo_mct 
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
         ItemData        =   "singlecore.frx":0134
         Left            =   1703
         List            =   "singlecore.frx":013B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4110
         Width           =   1305
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame current 
         Height          =   1335
         Left            =   4680
         TabIndex        =   21
         Top             =   4350
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   2355
         BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
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
         HeaderText      =   ""
         FooterText      =   ""
         HeaderText      =   ""
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
         Begin VB.TextBox result 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   345
            TabIndex        =   22
            Top             =   525
            Width           =   1620
         End
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame IA 
         Height          =   345
         Left            =   6690
         TabIndex        =   20
         Top             =   800
         Width           =   4680
         _ExtentX        =   8255
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
         HeaderText      =   "                                                                   In Air"
         FooterText      =   ""
         HeaderText      =   "                                                                   In Air"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame ISWD 
         Height          =   345
         Left            =   3450
         TabIndex        =   19
         Top             =   800
         Width           =   3105
         _ExtentX        =   5477
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
         HeaderText      =   "                 In Single way ducts"
         FooterText      =   ""
         HeaderText      =   "                 In Single way ducts"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin VB.PictureBox Pic_1c_ia_fs 
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
         Height          =   2085
         Left            =   9870
         Picture         =   "singlecore.frx":0143
         ScaleHeight     =   2025
         ScaleWidth      =   1455
         TabIndex        =   18
         Top             =   1185
         Width           =   1515
         Begin VB.PictureBox Pic_tic_7 
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
            Picture         =   "singlecore.frx":3950
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_1c_ia_ft 
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
         Height          =   2085
         Left            =   8275
         Picture         =   "singlecore.frx":3DC8
         ScaleHeight     =   2025
         ScaleWidth      =   1440
         TabIndex        =   17
         Top             =   1200
         Width           =   1500
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
            Picture         =   "singlecore.frx":74A5
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_1c_ia_tf 
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
         Height          =   2085
         Left            =   6668
         Picture         =   "singlecore.frx":791D
         ScaleHeight     =   2025
         ScaleWidth      =   1455
         TabIndex        =   16
         Top             =   1200
         Width           =   1515
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
            Picture         =   "singlecore.frx":B14A
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_1c_iswd_ftd 
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
         Height          =   2085
         Left            =   5061
         Picture         =   "singlecore.frx":B5C2
         ScaleHeight     =   2025
         ScaleWidth      =   1440
         TabIndex        =   15
         Top             =   1200
         Width           =   1500
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
            Picture         =   "singlecore.frx":F182
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_1c_iswd_tfd 
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
         Height          =   2085
         Left            =   3454
         Picture         =   "singlecore.frx":F5FA
         ScaleHeight     =   2025
         ScaleWidth      =   1410
         TabIndex        =   14
         Top             =   1215
         Width           =   1470
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
            Picture         =   "singlecore.frx":12E04
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame bdig 
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   800
         Width           =   3105
         _ExtentX        =   5477
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
         HeaderText      =   "Buried Direct into the ground"
         FooterText      =   ""
         HeaderText      =   "Buried Direct into the ground"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin VB.PictureBox Pic_1c_bdig_tf 
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
         Height          =   2085
         Left            =   240
         Picture         =   "singlecore.frx":1327C
         ScaleHeight     =   2025
         ScaleWidth      =   1500
         TabIndex        =   12
         Top             =   1200
         Width           =   1560
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
            Picture         =   "singlecore.frx":169B0
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   29
            Top             =   -45
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.PictureBox Pic_1c_bdig_fs 
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
         Height          =   2085
         Left            =   1847
         Picture         =   "singlecore.frx":16E28
         ScaleHeight     =   2025
         ScaleWidth      =   1500
         TabIndex        =   11
         Top             =   1200
         Width           =   1560
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
            Picture         =   "singlecore.frx":1A4AE
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   30
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_troed 
         Height          =   390
         Left            =   7343
         TabIndex        =   23
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
         HeaderText      =   "Thermal resitivity earthenware ducts"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "Thermal resitivity earthenware ducts"
         FooterText      =   ""
         HeaderPictureSize=   12
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_mct 
         Height          =   390
         Left            =   413
         TabIndex        =   24
         Top             =   3645
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
         HeaderText      =   "Maximum conductor tempersture ('C)"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "Maximum conductor tempersture ('C)"
         FooterText      =   ""
         HeaderPictureSize=   14
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_aat 
         Height          =   390
         Left            =   413
         TabIndex        =   25
         Top             =   4545
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
         HeaderText      =   "                Ambient air tempersture ('C)"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "                Ambient air tempersture ('C)"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_gt 
         Height          =   390
         Left            =   413
         TabIndex        =   26
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
         HeaderText      =   "       Ambient ground tempersture ('C)"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "       Ambient ground tempersture ('C)"
         FooterText      =   ""
         HeaderPictureSize=   14
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_tros 
         Height          =   390
         Left            =   7350
         TabIndex        =   27
         Top             =   4530
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
         HeaderText      =   "      Thermal resistivity of soil (Km/W)"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "      Thermal resistivity of soil (Km/W)"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_dol 
         Height          =   390
         Left            =   7343
         TabIndex        =   28
         Top             =   3645
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
         HeaderText      =   "                               Depth of laying (m)"
         FooterText      =   ""
         FooterBackColor =   16711680
         HeaderText      =   "                               Depth of laying (m)"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
      End
      Begin Light_ButtonTP.LSButtonT LSButton_calculate 
         Height          =   495
         Left            =   4635
         TabIndex        =   0
         Top             =   3780
         Width           =   2370
         _ExtentX        =   4180
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
         Caption         =   "Calculate"
      End
      Begin Light_ButtonTP.LSButtonT LSButton_back 
         Height          =   495
         Left            =   4545
         TabIndex        =   9
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
         TabIndex        =   8
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
      Begin Light_ButtonTP.LSButtonT LSButtonT_exit 
         Height          =   495
         Left            =   5467
         TabIndex        =   7
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
         Caption         =   "Exit"
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   -15
         X2              =   11625
         Y1              =   3510
         Y2              =   3495
      End
   End
End
Attribute VB_Name = "singlecore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
'Set wb = xlApp.Workbooks.Open(App.path & "\Excel\tables.xls")
Combo_mct.Text = 90
Combo_aat.Text = 30
Combo_gt.Text = 20
Combo_troed.Text = 1.2
Combo_tros.Text = 1.5
Combo_dol.Text = 0.8

singlecore.Caption = "  1 x " + xx + "  " + zz + " / " + yy
current.HeaderText = "  1 x " + xx + "  " + zz + " / " + yy

End Sub

Private Sub LSButton_back_Click()
Form2.Show
Unload Me
End Sub

Private Sub LSButton_calculate_Click()
c = 1
 x = Combo_aat.Text
    Select Case x
      Case 20
        c = 1.08
      Case 25
        c = 1.04
      Case 30
        c = 1
      Case 35
        c = 0.96
      Case 40
        c = 0.91
      Case 45
        c = 0.87
      Case 50
        c = 0.82
      Case 55
        c = 0.76
      Case 60
        c = 0.71
      End Select

d = 1
 y = Combo_gt.Text
    Select Case y
      Case 10
        d = 1.07
      Case 15
        d = 1.04
      Case 20
        d = 1
      Case 25
        d = 0.96
      Case 30
        d = 0.93
      Case 35
        d = 0.89
      Case 40
        d = 0.85
      Case 45
        d = 0.8
      Case 50
        d = 0.76
      End Select
      

    Select Case aa
      Case 1
        Set ws = wb.Worksheets("1c,xlpe,cu")
      Case 2
        Set ws = wb.Worksheets("1c,xlpe,al")
      Case 3
        Set ws = wb.Worksheets("1c,epr,cu")
      Case 4
        Set ws = wb.Worksheets("1c,epr,al")
      Case 5
        Set ws = wb.Worksheets("3c,xlpe,cu")
      Case 6
        Set ws = wb.Worksheets("3c,xlpe,al")
      Case 7
        Set ws = wb.Worksheets("3c,epr,cu")
      Case 8
        Set ws = wb.Worksheets("3c,epr,al")
      End Select


If a <= 11 Then
e = 2
Else
e = 3
End If

f = 6
 w = Combo_dol.Text
    Select Case w
      Case 0.5
        f = 4
      Case 0.6
        f = 5
      Case 0.8
        f = 6
      Case 1
        f = 7
      Case 1.25
        f = 8
      Case 1.5
        f = 9
      Case 1.75
        f = 10
      Case 2
        f = 11
      Case 2.5
        f = 12
      Case 3
        f = 13
      End Select

g = 6
 u = Combo_tros.Text
    Select Case u
      Case 0.7
        g = 2
      Case 0.8
        g = 3
      Case 0.9
        g = 4
      Case 1
        g = 5
      Case 1.5
        g = 6
      Case 2
        g = 7
      Case 2.5
        g = 8
      Case 3
        g = 9
      End Select

If Pic_tic_1.Visible = True Or Pic_tic_2.Visible = True Then
Set wss = wb.Worksheets("cor_fac_depths_buried_cable")
Set wsss = wb.Worksheets("cor_fac_soil_ter_resis_burid_1c")
bb = wss.Cells(f, e)
cc = wsss.Cells(a, g)
ElseIf Pic_tic_3.Visible = True Or Pic_tic_4.Visible = True Then
Set wss = wb.Worksheets("cor_fac_cable_in_duct")
Set wsss = wb.Worksheets("cor fac soil ter resis duct 1c")
bb = wss.Cells(f, e)
cc = wsss.Cells(a, g)
Else
cc = 1
bb = 1
End If

If dd <> 1 Then
Set ws_air = wb.Worksheets("1c_nasb")
ddd = ws_air.Cells(h + j, i)
Else
ddd = 1
End If

If Pic_tic_5.Visible = True Or Pic_tic_6.Visible = True Or Pic_tic_7.Visible = True Then
result.Text = ws.Cells(a, b) * c * ddd
ElseIf Pic_tic_1.Visible = True Or Pic_tic_2.Visible = True Or Pic_tic_3.Visible = True Or Pic_tic_4.Visible = True Then
result.Text = ws.Cells(a, b) * d * bb * cc
Else
MsgBox ("Please select one of Pictures")
End If

End Sub

Private Sub LSButton_reset_Click()
Combo_mct.Text = 90
Combo_aat.Text = 30
Combo_gt.Text = 20
Combo_troed.Text = 1.2
Combo_tros.Text = 1.5
Combo_dol.Text = 0.8
result.Text = ""

Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
Pic_tic_7.Visible = False
End Sub

Private Sub LSButtonT_exit_Click()
Z = MsgBox("  Are You Sure?", vbYesNo, "Exit")
If Z = 6 Then
    wb.Close
    xlApp.Quit
    Unload Me
    Unload Music
End If
End Sub

Private Sub Pic_1c_bdig_fs_Click()
b = 3
Pic_tic_1.Visible = False
Pic_tic_2.Visible = True
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
Pic_tic_7.Visible = False

End Sub

Private Sub Pic_1c_bdig_tf_Click()
b = 2
Pic_tic_1.Visible = True
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
Pic_tic_7.Visible = False

End Sub

Private Sub Pic_1c_ia_fs_Click()
b = 8
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
Pic_tic_7.Visible = True

single_in_air_horiz.Show
End Sub

Private Sub Pic_1c_ia_ft_Click()
b = 7
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = True
Pic_tic_7.Visible = False

single_in_air_horiz.Show

End Sub

Private Sub Pic_1c_ia_tf_Click()
b = 6
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = False
Pic_tic_5.Visible = True
Pic_tic_6.Visible = False
Pic_tic_7.Visible = False
single_in_air_trefoil.Show

End Sub

Private Sub Pic_1c_iswd_ftd_Click()
b = 5
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = False
Pic_tic_4.Visible = True
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
Pic_tic_7.Visible = False

End Sub

Private Sub Pic_1c_iswd_tfd_Click()
b = 4
Pic_tic_1.Visible = False
Pic_tic_2.Visible = False
Pic_tic_3.Visible = True
Pic_tic_4.Visible = False
Pic_tic_5.Visible = False
Pic_tic_6.Visible = False
Pic_tic_7.Visible = False

End Sub

