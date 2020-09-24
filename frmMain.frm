VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "BeMUD"
   ClientHeight    =   3000
   ClientLeft      =   8340
   ClientTop       =   1455
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5820
   Begin VB.Timer tmrDelayedComsMob 
      Enabled         =   0   'False
      Index           =   0
      Left            =   4320
      Top             =   1320
   End
   Begin VB.Timer tmrBleedingMob 
      Enabled         =   0   'False
      Index           =   0
      Left            =   4320
      Top             =   840
   End
   Begin VB.Timer tmrDelayedComs 
      Enabled         =   0   'False
      Index           =   0
      Left            =   5280
      Top             =   1260
   End
   Begin VB.Timer tmrBleeding 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   5280
      Top             =   840
   End
   Begin VB.TextBox txtOutput 
      Height          =   2175
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "OnTop"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1440
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock wskAccept 
      Index           =   0
      Left            =   5430
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   5445
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblPSC 
      AutoSize        =   -1  'True
      Caption         =   "Please click here to vote for BeMud on PlanetSourceCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      MouseIcon       =   "frmMain.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2760
      Width           =   4140
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "Users Management"
      Visible         =   0   'False
      Begin VB.Menu mnuUserIP 
         Caption         =   "IP"
      End
      Begin VB.Menu mnuKickUser 
         Caption         =   "Kick"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'*  Original BeMud copyright (C)1999-2000 by Vitaly Belman                  *
'*                                                                          *
'*  In order to use any part of BeMud, you must comply with                 *
'*  both the original BeMud license in 'license.doc'.  As stated in the     *
'*  'License.doc', you may not remove any orignal copyright.                *
'*                                                                          *
'*   BeMud is copyright 1999 - Vitaly Belman                                *
'*   BeMud is:                                                              *
'*       Vitaly Belman (vitali@actcom.co.il)                                *
'*       ICQ: 1912453                                                       *
'*                                                                          *
'*   By using this code, you have agreed to follow the terms of the         *
'*   BeMud license, in the file license.doc                                 *
'****************************************************************************
Option Explicit
Option Base 1
Private Sub Form_Load()
    Dim I As Integer
    Me.Show
    Set DB = OpenDatabase(App.Path & "\BeMud.mdb")
    LoadFromDatabaseToMemory 'Loads areas and emotes
    frmMain.Caption = "BeMUD - BeMud.selfhost.com"
    ReDim Char(1) 'Sets the UBound of Char to 1 to make ReDims possible via GetFreeWinsockIndex
    Randomize
'\B/-------------------------------Start to listen----------------------------------------
    wskListen.LocalPort = 23
    wskListen.Listen
'/E\-------------------------------Start to listen----------------------------------------
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim I%, Users As Variant
    Users = StringToArray(AllUsers)
    For I = LBound(Users) To UBound(Users)
        CloseConnection (Users(I))
    Next I
End Sub
'Logging the data on the output server screen
Sub Log(Data As String, Sign As String)
    frmMain.txtOutput.Text = frmMain.txtOutput.Text & Sign & ": " & Data & vbCrLf
    frmMain.txtOutput.SelStart = Len(frmMain.txtOutput.Text)
End Sub
Private Sub chkOnTop_Click()
    '\B/---------------------------Sets the window topmost------------------------------------
    Select Case chkOnTop.Value
    Case 1
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, _
      Me.Left \ Screen.TwipsPerPixelX, Me.Top \ Screen.TwipsPerPixelY, _
      Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY, 0)
    Case 0
    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, _
    Me.Left \ Screen.TwipsPerPixelX, Me.Top \ Screen.TwipsPerPixelY, _
    Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY, 0)
    End Select
   '/E\---------------------------Sets the window topmost------------------------------------
End Sub

Private Sub Form_DblClick()
    Clipboard.SetText "Host: bemud.hn.org" & "   Port: " & wskListen.LocalPort & vbNewLine & _
        "Host: " & wskListen.LocalIP & "   Port: " & wskListen.LocalPort
End Sub
'This sub keeps the objects on the form in the right size
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        chkOnTop.Top = ScaleHeight - chkOnTop.Height - 250
        chkOnTop.Left = ScaleWidth - chkOnTop.Width
        txtOutput.Width = ScaleWidth
        txtOutput.Height = ScaleHeight - txtInput.Height - 300
        txtInput.Width = ScaleWidth - chkOnTop.Width
        txtInput.Top = ScaleHeight - txtInput.Height - 250
        lblPSC.Top = ScaleHeight - lblPSC.Height - 20
        lblPSC.Left = (ScaleWidth / 2) - (lblPSC.Width / 2)
    End If
End Sub

Private Sub lblPSC_Click()
    Shell "Explorer http://planet-source-code.com/xq/ASP/txtCodeId.14000/lngWId.1/qx/vb/scripts/ShowCode.htm"
End Sub

Private Sub Label1_Click()

End Sub

Private Sub tmrBleedingMob_Timer(Index As Integer)
    
    Call BleedingTimerMob(Index)

End Sub
Private Sub tmrDelayedComsMob_Timer(Index As Integer)
    
    Call DelayedComsMob(Index)

End Sub

'Sends the text from the server side to selected clients
Private Sub txtInput_KeyPress(KeyAscii As Integer)
Dim I As Integer
Dim SysMessage As String
    If KeyAscii = 13 Then
        SysMessage = bRED & "System message - " & Date & ": " & RET & WHITE & "  " & txtInput.Text
        Call TransmitList(AllUsers, SysMessage)
        txtInput.SelStart = 0
        txtInput.SelLength = Len(txtInput.Text)
        KeyAscii = 0
    End If
End Sub
'Keeps the output scrollbar down-most
Private Sub txtOutput_Change()
    txtOutput.SelStart = Len(txtOutput.Text)
End Sub

Private Sub wskAccept_Close(Index As Integer)
    Call CloseConnection(Index)
End Sub
Private Sub wskAccept_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call WinsockError(Index, Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub
'Letting users to connect the mud
Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
    Call ConnectionRequest(requestID)
End Sub
Private Sub tmrBleeding_Timer(Index As Integer)
    Call BleedingTimer(Index)
End Sub
Private Sub tmrDelayedComs_Timer(Index As Integer)
    Call DelayedCommands(Index)
End Sub
Private Sub wskAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Call DataArrival(Index, bytesTotal)
End Sub

