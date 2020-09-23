VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Simple http Server Beta 1"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txterrcnt 
      Height          =   285
      Left            =   5880
      TabIndex        =   26
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txthits 
      Height          =   285
      Left            =   5880
      TabIndex        =   24
      Text            =   "0"
      Top             =   255
      Width           =   495
   End
   Begin VB.ListBox lststatus 
      Height          =   1620
      Left            =   45
      TabIndex        =   22
      Top             =   2790
      Width           =   5250
   End
   Begin VB.PictureBox BottomBar 
      Height          =   375
      Left            =   15
      ScaleHeight     =   315
      ScaleWidth      =   7200
      TabIndex        =   20
      Top             =   4470
      Width           =   7260
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   495
         TabIndex        =   21
         Top             =   60
         Width           =   45
      End
      Begin VB.Image imgser 
         Height          =   480
         Index           =   0
         Left            =   -60
         Picture         =   "Form1.frx":0000
         Top             =   -60
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   345
      Index           =   4
      Left            =   5475
      TabIndex        =   18
      Top             =   3060
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Settings"
      Enabled         =   0   'False
      Height          =   345
      Index           =   3
      Left            =   5475
      TabIndex        =   17
      Top             =   1740
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   1425
      TabIndex        =   11
      Top             =   5505
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   45
      TabIndex        =   3
      Top             =   120
      Width           =   5265
      Begin VB.TextBox txtCGIBase 
         Height          =   285
         Left            =   1155
         TabIndex        =   27
         Top             =   1695
         Width           =   2430
      End
      Begin VB.CheckBox chkcgi 
         Caption         =   "Enable Scriping"
         Height          =   195
         Left            =   3705
         TabIndex        =   19
         Top             =   1785
         Width           =   1455
      End
      Begin VB.TextBox txtport 
         Height          =   285
         Left            =   1170
         TabIndex        =   15
         Text            =   "80"
         Top             =   2130
         Width           =   510
      End
      Begin VB.TextBox txtPage 
         Height          =   285
         Left            =   1170
         TabIndex        =   13
         Top             =   1350
         Width           =   1410
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "...."
         Height          =   300
         Left            =   4125
         TabIndex        =   10
         Top             =   1005
         Width           =   390
      End
      Begin VB.TextBox txtRoot 
         Height          =   285
         Left            =   1170
         TabIndex        =   9
         Text            =   "\"
         Top             =   1005
         Width           =   2925
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   645
         Width           =   2910
      End
      Begin VB.TextBox txthostName 
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2925
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Script Path"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Default Port"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   2190
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Default Page"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   1395
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Home Path"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "IP Address"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hostname"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   285
         Width           =   720
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   6480
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   345
      Index           =   2
      Left            =   5475
      TabIndex        =   2
      Top             =   2610
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Stop Server"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   5475
      TabIndex        =   1
      Top             =   2175
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Server"
      Height          =   345
      Index           =   0
      Left            =   5475
      TabIndex        =   0
      Top             =   1305
      Width           =   1590
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Errors"
      Height          =   195
      Left            =   5385
      TabIndex        =   25
      Top             =   645
      Width           =   405
   End
   Begin VB.Label lblhits 
      AutoSize        =   -1  'True
      Caption         =   "Hits"
      Height          =   195
      Left            =   5385
      TabIndex        =   23
      Top             =   300
      Width           =   270
   End
   Begin VB.Image imgser 
      Height          =   480
      Index           =   2
      Left            =   315
      Picture         =   "Form1.frx":0442
      Top             =   6090
      Width           =   480
   End
   Begin VB.Image imgser 
      Height          =   480
      Index           =   1
      Left            =   345
      Picture         =   "Form1.frx":0884
      Top             =   6075
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Users As Integer
Dim ErrorsCnt As Integer
Dim Main_Path As String
Dim ErrPagesPath As String
Dim Parm As String

Sub RunCgi(lzCgiPath As String, Index)
Dim TFile As Long
Dim Data As String
    TFile = FreeFile
        Open lzCgiPath For Binary Access Read As #TFile
            Data = Space(LOF(TFile))
            Get #TFile, , Data
        Close #TFile
        ' send header
    header = "HTTP/1.0 200 OK" & vbCrLf
    header = header & "Server: DMServer Beta 1" & vbCrLf
    header = header & "Content-Type: " & "text/html" & vbCrLf
    header = header & "Accept-Ranges: bytes" & vbCrLf
    header = header & "Content-Length: " & LTrim(Str(Len(Data))) & vbCrLf
    header = header & vbCrLf
    Winsock1(Index).SendData header + Data
   
    
End Sub

Sub BadPage(lzPagePath As String, mIndex)
Dim TFile As Long
Dim StrBuff() As Byte
    TFile = FreeFile
        Open lzPagePath For Binary As #TFile
            ReDim StrBuff(0 To LOF(TFile))
            Get #TFile, , StrBuff()
        Close #TFile
    Winsock1(mIndex).SendData StrBuff
    ErrorsCnt = ErrorsCnt + 1
    txterrcnt = ErrorsCnt
End Sub

Sub SendData(page As String, Index)
Dim databyte() As Byte
Dim r As String
Dim Res As String

On Error Resume Next
    page = ConvertWebSlash(page)
    If page = "" Then page = txtPage.Text
    If FSO.FileExists(txtRoot.Text & page) Then
    
        Open txtRoot.Text & page For Binary Shared As #1
            ReDim databyte(0 To LOF(1))
            Get #1, , databyte()
            Close #1
        Winsock1(Index).SendData databyte()
        Exit Sub
    Else
        BadPage addSlash(App.Path) & "errors\err404.htm", Index
    End If
    
End Sub
Private Sub cmdPath_Click()
Dim mPath As String
    mPath = GetFolder(Form1.hwnd, "Please select your servers home folder")
    If mPath = "" Then mPath = "\"
    txtRoot.Text = addSlash(mPath)
    
    
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            On Error GoTo Terr
            Winsock1(0).LocalPort = Val(txtport.Text)
            Winsock1(0).Listen
            Command1(0).Enabled = False
            Command1(1).Enabled = True
            imgser(0).Picture = imgser(2).Picture
            lblStatus.Caption = "Status : Server is running."
            If Err Then
Terr:
                MsgBox "The port you requested this server to run on is in use please check you have not other programs useing this port", vbInformation, "Address in use"
            End If
        Case 1
            Winsock1(0).Close
            Command1(0).Enabled = True
            Command1(1).Enabled = False
            imgser(0).Picture = imgser(1).Picture
            lblStatus.Caption = "Status : Server has been stoped."
    Case 2
        If Winsock1(0).State = sckListening Then
            MsgBox "The server is still running please shut down before you exit this program", vbInformation, "Exit...."
        Else
            End
        End If
        
    Case 3
        If IsNumeric(txtport.Text) = False Then MsgBox "Inviald port number entered settings will not be saved", vbCritical, "Inviald Port Number": Exit Sub
        WritePrivateProfileString "DmServer", "HomeRoot", txtRoot.Text, Conf_Filename
        WritePrivateProfileString "DmServer", "Default", txtPage.Text, Conf_Filename
        WritePrivateProfileString "DmServer", "ScriptPath", txtCGIBase.Text, Conf_Filename
        WritePrivateProfileString "DmServer", "Port", txtport.Text, Conf_Filename
        WritePrivateProfileString "DmServer", "CGIEnabled", CStr(chkcgi.Value), Conf_Filename
        
        MsgBox "new settings have been updated", vbInformation, "Update settings"
End Select

End Sub


Private Sub Form_Load()
    Main_Path = addSlash(App.Path)
    ErrPagesPath = addSlash(App.Path) & "errors\"
    Conf_Filename = Main_Path & "config.ini"
    Set FSO = New FileSystemObject
    If FSO.FileExists(Conf_Filename) = False Then
        MsgBox "You need to config your setting before this server can be run", vbInformation, "First time use"
    Else
        txtRoot.Text = ReadConfig("DmServer", "HomeRoot")
        txtPage.Text = ReadConfig("DmServer", "Default")
        txtCGIBase.Text = ReadConfig("DmServer", "ScriptPath")
        txtport.Text = ReadConfig("DmServer", "Port")
        chkcgi.Value = ReadConfig("DmServer", "CGIEnabled")
    End If
    
    txthostName = "http://" & Winsock1(0).LocalHostName
    txtIP.Text = Winsock1(0).LocalIP
    lblStatus.Caption = "Status : ide"
    
    FlatBorder txthostName.hwnd
    FlatBorder txtIP.hwnd
    FlatBorder txtRoot.hwnd
    FlatBorder cmdPath.hwnd
    FlatBorder txtPage.hwnd
    FlatBorder txtport.hwnd
    FlatBorder txthits.hwnd
    FlatBorder txterrcnt.hwnd
    FlatBorder txtCGIBase.hwnd
    FlatBorder BottomBar.hwnd
    For I = 0 To Command1.Count - 1
        FlatBorder Command1(I).hwnd
    Next
        
End Sub

Private Sub txtCGIBase_Change()
    Command1(3).Enabled = True
    
End Sub

Private Sub txtPage_Change()
    Command1(3).Enabled = True
    
End Sub

Private Sub txtport_Change()
    Command1(3).Enabled = True
    
End Sub

Private Sub txtRoot_Change()
    Command1(3).Enabled = True
    
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Users = Users + 1
If Index = 0 Then
    Load Winsock1(Users)
    Winsock1(Users).Accept requestID
    Connections = Connections + 1
    txthits.Text = Users - 1
End If
    If Err Then Err.Clear
    
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim StrData As String
Dim page As String, ExecuteCmd As String, strPost As String
Dim ipart As Integer, lpart As Integer
On Error Resume Next
    Winsock1(Index).GetData StrData, , bytesTotal
    
    If Left(StrData, 4) = "GET " Then
        page = Trim(Mid(StrData, 5, InStr(5, StrData, " ") - 5))
    End If
    
    If Left(page, 1) = "/" Then page = Right(page, Len(page) - 1)
    lststatus.AddItem Winsock1(Index).RemoteHostIP & " " & " GET" & page & " HTTP/1.1"
    If InStr(page, "cgi-bin") Then
        If chkcgi Then
            cgi = ConvertWebSlash(Mid(page, 1, InStr(2, page, "/")))
            cgi = Left(cgi, Len(cgi) - 1)
            l = "\" & cgi
            If page = "cgi-bin/" Then BadPage addSlash(App.Path) & "errors\err413.htm", Index: Exit Sub
            If cgi = "" Then BadPage addSlash(App.Path) & "errors\err413.htm", Index: Exit Sub
            If Not l = txtCGIBase Then BadPage addSlash(App.Path) & "errors\err501.htm", Index: Exit Sub
            ExecuteCmd = ConvertWebSlash(txtRoot.Text & page)
            If Not FSO.FileExists(ExecuteCmd) Then
                BadPage ErrPagesPath & "err404.htm", Index
                Exit Sub
            Else
                Select Case LCase(Right(ExecuteCmd, 3))
                    Case ".pl"
                        p = "c:\perl\bin\perl.exe " & ExecuteCmd & ">" & txtRoot.Text & "cgi-bin\temp.htm"
                        Shell "cmd.exe /c" & p, vbHide
                        RunCgi txtRoot.Text & "cgi-bin\temp.htm", Index
                    Case "exe"
                        p = "c:\perl\bin\perl.exe " & ExecuteCmd & ">" & txtRoot.Text & "cgi-bin\temp.htm"
                        Shell "cmd.exe /c" & p, vbHide
                        RunCgi txtRoot.Text & "cgi-bin\temp.htm", Index
                    Case "cgi"
                        p = "c:\perl\bin\perl.exe " & ExecuteCmd & ">" & txtRoot.Text & "cgi-bin\temp.htm"
                        Shell "cmd.exe /c" & p, vbHide
                        RunCgi txtRoot.Text & "cgi-bin\temp.htm", Index
                    Case Else
                        BadPage ErrPagesPath & "err503.htm", Index
                End Select
            End If
        Else
            BadPage addSlash(App.Path) & "errors\err502.htm", Index: Exit Sub
        End If
        
        Exit Sub
    Else
        SendData page, Index
    End If
    page = ""
    Data = ""
    cgi = ""
    l = ""
    Parm = ""
    
    


    
    
    
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock1(Index).Close
    Unload Winsock1(Index)
    
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    Winsock1(Index).Close
    Unload Winsock1(Index)
    
End Sub
