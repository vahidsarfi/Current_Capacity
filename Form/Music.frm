VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Music 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   630
   ClientLeft      =   615
   ClientTop       =   405
   ClientWidth     =   3990
   LinkTopic       =   "Form4"
   ScaleHeight     =   630
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer3 
      Height          =   645
      Left            =   -30
      TabIndex        =   0
      Top             =   -15
      Width           =   3990
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
      _cx             =   7038
      _cy             =   1138
   End
End
Attribute VB_Name = "Music"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
WindowsMediaPlayer3.URL = App.Path & "\music\1.mp3"
End Sub

'\B/--------------------------------------------------------------------------------
'/E\--------------------------------------------------------------------------------
Private Sub WindowsMediaPlayer3_OpenStateChange(ByVal NewState As Long)

End Sub
