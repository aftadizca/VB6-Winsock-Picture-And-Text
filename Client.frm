VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Client 
   BackColor       =   &H00404040&
   Caption         =   "CLIENT"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14910
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Tag             =   "Test"
   Begin VB.FileListBox RcvImg 
      Appearance      =   0  'Flat
      Height          =   4590
      Left            =   11400
      Pattern         =   "*.jpg;*.png"
      TabIndex        =   13
      Top             =   360
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   4590
      Left            =   7080
      ScaleHeight     =   135.111
      ScaleMode       =   0  'User
      ScaleWidth      =   4200
      TabIndex        =   12
      Top             =   360
      Width           =   4230
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "BROWSE...."
      Height          =   720
      Left            =   7080
      TabIndex        =   11
      Top             =   5040
      Width           =   4215
   End
   Begin VB.TextBox Username 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Text            =   "Client"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox IpText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox PortText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Start 
      Appearance      =   0  'Flat
      Caption         =   "CONNECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3960
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton CloseBtn 
      Appearance      =   0  'Flat
      Caption         =   "CLOSE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox ChatDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   6735
   End
   Begin VB.TextBox Message 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   4815
   End
   Begin VB.CommandButton SendMsg 
      Appearance      =   0  'Flat
      Caption         =   "SEND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   5040
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bSendingFile As Boolean
Private lTotal As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim bFileArriving As Boolean
Dim sFile As String
Dim sArriving As String

Dim FileName As String
Dim FileTitle As String

Dim pb As New PropertyBag
Private Sub CloseBtn_Click()
    Winsock1.Close
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Closed :" & Winsock1.RemoteHost & vbCrLf
    Start.Enabled = True
    CloseBtn.Enabled = False
End Sub

Private Sub Command1_Click()
    CommonDialog1.Filter = "IMAGE|*.jpg"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Picture1.Picture = LoadPicture(CommonDialog1.FileName)
        FileName = CommonDialog1.FileName
        FileTitle = CommonDialog1.FileTitle
    End If
End Sub

Private Sub Form_Load()
    IpText.Text = Winsock1.LocalIP
    PortText.Text = "11111"
    
    Dim imgPath As String
    imgPath = App.Path & "\tempClient\"
    
    Dim oDir As New Scripting.FileSystemObject
    If Not oDir.FolderExists(imgPath) Then
        MkDir (imgPath)
    End If
    
    RcvImg.Path = App.Path & "\tempClient\"
End Sub

Private Sub RcvImg_Click()
    Picture1.Picture = LoadPicture(RcvImg.Path & "\" & RcvImg.FileName)
    FileName = RcvImg.Path & "\" & RcvImg.FileName
    FileTitle = RcvImg.FileName
End Sub

Private Sub SendMsg_Click()
    'Set pb = New PropertyBag
    'pb.WriteProperty "username", Username.Text, "Client"
    'pb.WriteProperty "message", Message.Text, 0
    'pb.WriteProperty "pic", Picture1.Picture, 0
    
    If Message.Text <> "" Or FileName <> "" Then
            SendData FileName, FileTitle, Message.Text, Winsock1
            'Winsock1(Index).SendData "Server : " & Message.Text
        If Message.Text <> "" Then
            ChatDisplay.Text = ChatDisplay.Text & "ME : " & Message.Text & vbCrLf
        End If
    End If
    Message.Text = ""
    Picture1.Picture = LoadPicture()
    FileName = ""
    FileTitle = ""
    
End Sub

Private Sub Start_Click()
    Winsock1.Connect Winsock1.LocalIP, "11111"
    Start.Enabled = False
    CloseBtn.Enabled = True
End Sub

Private Sub Winsock1_Close()
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Closed :" & Winsock1.RemoteHost & vbCrLf
    Start.Enabled = True
    CloseBtn.Enabled = False
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Started :" & Winsock1.RemoteHost & vbCrLf
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    'Dim MsgData As String
    'Winsock1.GetData MsgData, vbString
    'ChatDisplay.Text = ChatDisplay.Text & MsgData & vbCrLf
    
    
    
    Dim strData As String
    Dim ifreefile
    
    DoEvents
    Winsock1.GetData strData
    
    If Right$(strData, 7) = "FILEEND" Then
        bFileArriving = False
        'lblProgress = "Saving File to " & App.Path & "\tempClient\" & sFile
        sArriving = sArriving & Left$(strData, Len(strData) - 7)
        
        ifreefile = FreeFile
        
        If Dir(App.Path & "\tempClient\" & sFile) <> "" Then
            MsgBox "File Already Exists"
        Else
            Open App.Path & "\tempClient\" & sFile For Binary Access Write As #ifreefile
            Put #ifreefile, 1, sArriving
            Close #ifreefile
            'ShellExecute 0, vbNullString, sFile, vbNullString, vbNullString, vbNormalFocus
            RcvImg.Refresh
            ChatDisplay.Text = ChatDisplay.Text & sFile & " received from " & Winsock1.RemoteHostIP & vbCrLf
        End If
        sArriving = ""
        'lblProgress = "Complete"
    ElseIf Left$(strData, 4) = "FILE" Then
        bFileArriving = True
        sFile = Right$(strData, Len(strData) - 4)
    ElseIf Left$(strData, 5) = "MSSGG" Then
        If Right$(strData, Len(strData) - 5) <> "" Then
            ChatDisplay.Text = ChatDisplay.Text & Winsock1.RemoteHostIP & " : " & Right$(strData, Len(strData) - 5) & vbCrLf
        End If
        bFileArriving = False
    ElseIf bFileArriving Then
        'lblProgress = "Receiving " & bytesTotal & " bytes for " & sFile & " from " & tcpClient.RemoteHostIP
        sArriving = sArriving & strData
    End If
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Failed :" & Winsock1.RemoteHost & vbCrLf
    Winsock1.Close
    Start.Enabled = True
    CloseBtn.Enabled = False
End Sub


'Function send file
Public Sub SendData(ByVal sFile As String, ByVal sSaveAs As String, ByVal msg As String, ByVal tcpCtl As Winsock)
'On Error GoTo ErrHandler
    Dim sSend As String, sBuf As String
    Dim ifreefile As Integer
    Dim lRead As Long, lLen As Long, lThisRead As Long, lLastRead As Long
    
    ifreefile = FreeFile
    
    If sFile <> "" Then
        tcpCtl.SendData "MSSGG" & msg
        DoEvents
        ' Open file for binary access:
        Open sFile For Binary Access Read As #ifreefile
        lLen = LOF(ifreefile)
        
        ' Loop through the file, loading it up in chunks of 64k:
        Do While lRead < lLen
            lThisRead = 65536
            If lThisRead + lRead > lLen Then
                lThisRead = lLen - lRead
            End If
            If Not lThisRead = lLastRead Then
                sBuf = Space$(lThisRead)
            End If
            Get #ifreefile, , sBuf
            lRead = lRead + lThisRead
            sSend = sSend & sBuf
        Loop
        lTotal = lLen
        Close ifreefile
        bSendingFile = True
        '// Send the file notification
        tcpCtl.SendData "FILE" & sSaveAs
        DoEvents
        '// Send the file
        tcpCtl.SendData sSend
        DoEvents
        '// Finished
        tcpCtl.SendData "FILEEND"
        bSendingFile = False
        Exit Sub
    Else
        tcpCtl.SendData "MSSGG" & msg
        DoEvents
    End If
'ErrHandler:
    'MsgBox "Errorssss " & Err & " : " & Error
End Sub
