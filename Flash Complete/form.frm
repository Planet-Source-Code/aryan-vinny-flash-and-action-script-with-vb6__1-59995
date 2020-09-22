VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Form1 
   Caption         =   "How To Use Flash  In VB Example!"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin Project1.isButton isButton1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Icon            =   "form.frx":0000
      Style           =   0
      Caption         =   "Walk"
      IconAlign       =   1
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   33023
      FontColor       =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _cx             =   10821
      _cy             =   8916
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Call ShockwaveFlash1.LoadMovie(0, CurDir + "/t.swf")
End Sub


Private Sub isButton1_Click()
''Sets The Variable "Move" In The Fla File As 1...
Call ShockwaveFlash1.SetVariable("Move", Str(1))
End Sub

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args_ As String)
''Retrives The Commands From The Within The Flash File...
If command = "fin" And args_ = "true" Then
    MsgBox "Yeah!...kinda cool huh lol?"
    
    
    Form1.Refresh
    
End If
'this catches the action sent by Flash, and uses it in the TextBox.
End Sub

