VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{5E347DF1-C0CC-11D9-84B3-9100B9C7DC45}#6.0#0"; "BZ_Frame2.ocx"
Object = "{C0F23CC2-0EE2-4CD5-975D-EDA99765C6F1}#1.0#0"; "Light_Transparent_Button.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   7170
   ClientLeft      =   3930
   ClientTop       =   2100
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7785
   Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_XP_Frame1 
      Height          =   7215
      Left            =   -45
      TabIndex        =   1
      Top             =   -30
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   12726
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
      Begin Babak_Zawari_Frame.BZ_XP_Frame Nominal_Current 
         Height          =   1170
         Left            =   2632
         TabIndex        =   23
         Top             =   4200
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   2064
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
         HeaderText      =   "Nominal area of conductor (mm2)"
         FooterText      =   ""
         HeaderText      =   "Nominal area of conductor (mm2)"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
         Begin VB.ComboBox Combo 
            DataMember      =   "qw"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Form2.frx":0000
            Left            =   555
            List            =   "Form2.frx":0028
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   420
            Width           =   1395
         End
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame Core 
         Height          =   1650
         Left            =   5145
         TabIndex        =   18
         Top             =   2175
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2910
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
         HeaderText      =   "No. of Cores"
         FooterText      =   ""
         HeaderText      =   "No. of Cores"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
         Begin VB.OptionButton O_triplecore 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   20
            Top             =   1020
            Width           =   195
         End
         Begin VB.OptionButton O_singlecore 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   19
            Top             =   525
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.Label L_TC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Triple Core"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   555
            TabIndex        =   22
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label L_SC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Single Core"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   555
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame Insulation 
         Height          =   1650
         Left            =   2865
         TabIndex        =   13
         Top             =   2175
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2910
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
         HeaderText      =   "Insulation"
         FooterText      =   ""
         HeaderText      =   "Insulation"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
         Begin VB.OptionButton O_EPR 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1020
            Width           =   180
         End
         Begin VB.OptionButton O_XLPE 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   510
            Value           =   -1  'True
            Width           =   180
         End
         Begin VB.Label L_EPR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EPR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   675
            TabIndex        =   17
            Top             =   990
            Width           =   420
         End
         Begin VB.Label L_XLPE 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "XLPE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   675
            TabIndex        =   16
            Top             =   480
            Width           =   525
         End
      End
      Begin Babak_Zawari_Frame.BZ_XP_Frame BZ_XP_Frame2 
         Height          =   1650
         Left            =   585
         TabIndex        =   8
         Top             =   2175
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2910
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
         HeaderText      =   "Conductor"
         FooterText      =   ""
         HeaderText      =   "Conductor"
         FooterText      =   ""
         ZOrderOnFocus   =   0   'False
         Begin VB.OptionButton O_Alu 
            Caption         =   "Alu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   285
            TabIndex        =   10
            Top             =   1035
            Width           =   210
         End
         Begin VB.OptionButton O_Cu 
            Caption         =   "Cu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   285
            TabIndex        =   9
            Top             =   525
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.Label L_Alu 
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Alu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   705
            TabIndex        =   12
            Top             =   1005
            Width           =   360
         End
         Begin VB.Label L_Cu 
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   705
            TabIndex        =   11
            Top             =   480
            Width           =   285
         End
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5152
         TabIndex        =   5
         Top             =   1425
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1297
         TabIndex        =   4
         Top             =   1425
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3442
         TabIndex        =   3
         Top             =   1425
         Width           =   1215
      End
      Begin Light_ButtonTP.LSButtonT LSButton_Exit 
         Height          =   495
         Left            =   997
         TabIndex        =   2
         Top             =   5850
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
         Caption         =   "Exit"
      End
      Begin Light_ButtonTP.LSButtonT LSButton_Next 
         Height          =   495
         Left            =   4425
         TabIndex        =   0
         Top             =   5850
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
         Caption         =   "Next"
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   0
         Left            =   300
         Picture         =   "Form2.frx":0062
         Stretch         =   -1  'True
         Top             =   1335
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   6765
         Picture         =   "Form2.frx":036C
         Stretch         =   -1  'True
         Top             =   1282
         Width           =   495
      End
      Begin VB.Label L_Title 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   22.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3045
         TabIndex        =   7
         Top             =   510
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   0
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   855
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1508
      _cy             =   873
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "                Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Setting"
      Begin VB.Menu Password 
         Caption         =   "              Change Password"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Audio_Click()
'If ws_pass.Cells(2, 1) = 1 Then
'ws_pass.Cells(2, 1) = 2
'Audio.Caption = "                   Audio Off"
'wb.Save
'Else
'ws_pass.Cells(2, 1) = 1
'Audio.Caption = "                   Audio On"
'wb.Save
'End If

'End Sub

Public Sub Combo_Change()

End Sub


Private Sub exit_Click()
Z = MsgBox("  Are You Sure?", vbYesNo, "Exit")
If Z = 6 Then
    wb.Close
    xlApp.Quit
    Unload Me
    Unload Music
End If
End Sub

Private Sub Form_Activate()
'Set ws_pass = wb.Worksheets("sheet1")
'If ws_pass.Cells(2, 1) = 1 Then
'WindowsMediaPlayer1.URL = App.Path & "\music\1.mp3"
'Audio.Caption = "                 Audio Off"
'Else
'Audio.Caption = "                 Audio On"
'End If


End Sub

Private Sub Form_Load()
dd = 1
Dim m As Date
Dim pYear As Integer, pMonth  As Integer, pDay As Integer, DayName As String
GetJalaliDate Year(Date), Month(Date), Day(Date), pYear, pMonth, pDay, DayName
Text2.Text = " " & pYear & "/" & pMonth & "/" & pDay
Text3.Text = "   " & DayName
m = Text2.Text
If m = #6/1/1390# Then
MsgBox (" »—‰«„Â ‰Ì«“ »Â ¬ÅœÌ  œ«—œ ·ÿ›« »« ‘„«—Â  ·›‰ 0911111111  „«” Õ«’· ›—„«ÌÌœ")
End
End If

Combo.Text = 16

 
End Sub

Public Sub LSButton_Next_Click()


    xx = Combo.Text
    Select Case xx
      Case 16
        a = 3
      Case 25
        a = 4
      Case 35
        a = 5
      Case 50
        a = 6
      Case 70
        a = 7
      Case 95
        a = 8
      Case 120
        a = 9
      Case 150
        a = 10
      Case 185
        a = 11
      Case 240
        a = 12
      Case 300
        a = 13
      Case 400
        a = 14
    End Select







If O_singlecore.Value = True Then

  If O_XLPE.Value = True Then
    yy = "XLPE"
  
    If O_Cu.Value = True Then
      zz = "Cu"
      aa = 1
    Else
      zz = "Alu"
      aa = 2
    End If
        
  Else
    yy = "EPR"
    If O_Cu.Value = True Then
      zz = "Cu"
      aa = 3
    Else
      zz = "Alu"
      aa = 4
    
    End If
    
  End If
  
singlecore.Show
Unload Me

Else

  If O_XLPE.Value = True Then
    yy = "XLPE"
  
    If O_Cu.Value = True Then
    zz = "Cu"
    aa = 5
    
    Else
    zz = "Alu"
    aa = 6
    
    End If
    
  Else
  yy = "EPR"
    If O_Cu.Value = True Then
    zz = "Cu"
    aa = 7
        
    Else
    zz = "Alu"
    aa = 8
    End If
        
  End If

triplecore.Show
Unload Me

End If




End Sub

Private Sub LSButton_Exit_Click()
Z = MsgBox("  Are You Sure?", vbYesNo, "Exit")
If Z = 6 Then
    wb.Close
    xlApp.Quit
    Unload Me
    Unload Music
End If
End Sub

Private Sub Password_Click()
Form3.Show

End Sub

Private Sub Timer1_Timer()
Text1.Text = Time
End Sub
