VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Server 
   BackColor       =   &H00404040&
   Caption         =   "SERVER"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14775
   DrawStyle       =   5  'Transparent
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
   ScaleHeight     =   5925
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox RcvImg 
      Appearance      =   0  'Flat
      Height          =   4590
      Left            =   11280
      Pattern         =   "*.jpg;*.bmp"
      TabIndex        =   11
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BROWSE...."
      Height          =   720
      Left            =   6960
      TabIndex        =   10
      Top             =   5040
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   6960
      ScaleHeight     =   4515
      ScaleWidth      =   4155
      TabIndex        =   9
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton SendMsg 
      Caption         =   "SEND"
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Message 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox ChatDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1560
      Width           =   6735
   End
   Begin VB.CommandButton CloseBtn 
      Appearance      =   0  'Flat
      Caption         =   "CLOSE"
      Enabled         =   0   'False
      Height          =   840
      Left            =   5400
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Start 
      Appearance      =   0  'Flat
      Caption         =   "LISTEN"
      Height          =   840
      Left            =   3840
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   6600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Left            =   120
      TabIndex        =   3
      Top             =   960
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
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
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bSendingFile As Boolean
Private lTotal As Long
Public NumSockets As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim bFileArriving As Boolean
Dim sFile As String
Dim sArriving As String

Dim FileName As String
Dim FileTitle As String

Private Sub CloseBtn_Click()
    Dim Index As Integer
        For Index = 0 To NumSockets
            Winsock1(Index).Close
        Next Index
    ChatDisplay.Text = ChatDisplay.Text & "----Server shutdown" & vbCrLf
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
    IpText.Text = Winsock1(0).LocalIP
    PortText.Text = "11111"
    Dim Client1 As New Client
    Dim Client2 As New Client
    Dim Client3 As New Client
    
    Dim imgPath As String
    imgPath = App.Path & "\tempServer\"
    
    Dim oDir As New Scripting.FileSystemObject
    If Not oDir.FolderExists(imgPath) Then
        MkDir (imgPath)
    End If
    
    RcvImg.Path = imgPath
    
    Client1.Show
    'Client2.Show
    'Client3.Show
End Sub


Private Sub RcvImg_Click()
    Picture1.Picture = LoadPicture(RcvImg.Path & "\" & RcvImg.FileName)
    FileName = RcvImg.Path & "\" & RcvImg.FileName
    FileTitle = RcvImg.FileName
End Sub

Private Sub SendMsg_Click()
    Dim Index As Integer
    Debug.Print CommonDialog1.FileName
    If Message.Text <> "" Or FileName <> "" Then
        For Index = 1 To NumSockets
            SendData FileName, FileTitle, Message.Text, Winsock1(Index)
            'Winsock1(Index).SendData "Server : " & Message.Text
        Next Index
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
    Winsock1(0).LocalPort = PortText.Text
    ChatDisplay.Text = ChatDisplay.Text & "----Starting server in port " & Winsock1(0).LocalPort & vbCrLf
    Winsock1(0).Listen
    Start.Enabled = False
    CloseBtn.Enabled = True
End Sub

Private Sub Winsock1_Close(Index As Integer)
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Closed :" & Winsock1(Index).RemoteHostIP & vbCrLf
    Winsock1(Index).Close
    'Unload Winsock1(Index)
End Sub

Private Sub Winsock1_Connect(Index As Integer)
    Debug.Print "Winsock Connect"
    ChatDisplay.Text = ChatDisplay.Text & "----Connected : " & Winsock1(Index).RemoteHostIP & vbCrLf
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    NumSockets = NumSockets + 1
    Load Winsock1(NumSockets)
    Winsock1(NumSockets).Accept requestID
    ChatDisplay.Text = ChatDisplay.Text & "----Connected : " & Winsock1(Index).RemoteHostIP & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    Dim ifreefile
    
    Debug.Print Index
    
    DoEvents
    Winsock1(Index).GetData strData
    
    If Right$(strData, 7) = "FILEEND" Then
        bFileArriving = False
        sArriving = sArriving & Left$(strData, Len(strData) - 7)
        
        ifreefile = FreeFile
        
        If Dir(App.Path & "\tempServer\" & sFile) <> "" Then
            MsgBox "File Already Exists"
        Else
            Open App.Path & "\tempServer\" & sFile For Binary Access Write As #ifreefile
            Put #ifreefile, 1, sArriving
            Close #ifreefile
            'ShellExecute 0, vbNullString, sFile, vbNullString, vbNullString, vbNormalFocus
            RcvImg.Refresh
            ChatDisplay.Text = ChatDisplay.Text & sFile & " received from " & Winsock1(Index).RemoteHostIP & vbCrLf
        End If
        sArriving = ""
    ElseIf Left$(strData, 4) = "FILE" Then
        bFileArriving = True
        sFile = Right$(strData, Len(strData) - 4)
    ElseIf Left$(strData, 5) = "MSSGG" Then
        If Right$(strData, Len(strData) - 5) <> "" Then
            ChatDisplay.Text = ChatDisplay.Text & Winsock1(Index).RemoteHostIP & " : " & Right$(strData, Len(strData) - 5) & vbCrLf
        End If
        bFileArriving = False
    ElseIf bFileArriving Then
        sArriving = sArriving & strData
    End If
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








