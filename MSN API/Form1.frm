VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Windows Messenger"
   ClientHeight    =   7140
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3360
      Top             =   5880
   End
   Begin MSComctlLib.ImageList BigStatIco 
      Left            =   1680
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   42
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2724
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3B92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2880
      Top             =   5880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   5880
   End
   Begin PicClip.PictureClip PC 
      Left            =   240
      Top             =   8520
      _ExtentX        =   21696
      _ExtentY        =   1085
      _Version        =   393216
      Cols            =   20
      Picture         =   "Form1.frx":5000
   End
   Begin MSComctlLib.ImageList img 
      Left            =   3840
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   17
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DA4E
            Key             =   "online"
            Object.Tag             =   "Pic_On"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DDE6
            Key             =   "Pic_Off"
            Object.Tag             =   "Pic_Off"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E17E
            Key             =   "Pic_time"
            Object.Tag             =   "Pic_time"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E516
            Key             =   "Pic_away"
            Object.Tag             =   "Pic_away"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E8AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EC74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F07E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F488
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9128
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   1
      ImageList       =   "img"
      Appearance      =   0
   End
   Begin VB.Label ConStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refreshing..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   1200
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Stat 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   330
   End
   Begin VB.Image BigStat 
      Height          =   495
      Left            =   50
      Top             =   120
      Width           =   495
   End
   Begin VB.Label UserName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Status:"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   120
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu RightMnu 
      Caption         =   "RightMnu"
      Visible         =   0   'False
      Begin VB.Menu SendIMMnu 
         Caption         =   "Send an Instant Message"
      End
      Begin VB.Menu StartVoiceMnu 
         Caption         =   "Start a Voice Conversation"
      End
      Begin VB.Menu StartVideoMnu 
         Caption         =   "Start a Video Conversation"
      End
      Begin VB.Menu SendFIleMnu2 
         Caption         =   "Send a File or Photo"
         Index           =   0
      End
      Begin VB.Menu SentEmailToMnu 
         Caption         =   "Send E-mail ()"
      End
      Begin VB.Menu bar6 
         Caption         =   "-"
      End
      Begin VB.Menu RemoteAssistanceMnu 
         Caption         =   "Ask for Remote Assistance"
      End
      Begin VB.Menu StartAppMnu 
         Caption         =   "Start Application Sharing"
      End
      Begin VB.Menu StartWhiteboardMnu 
         Caption         =   "Start Whiteboard"
      End
      Begin VB.Menu bar7 
         Caption         =   "-"
      End
      Begin VB.Menu BlockMnu 
         Caption         =   "Block"
      End
      Begin VB.Menu DeleteContactMnu 
         Caption         =   "Delete Contact"
      End
      Begin VB.Menu ViewProfileMnu 
         Caption         =   "View Profile"
      End
      Begin VB.Menu PropertiesMnu 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu AutoSignInMnu 
         Caption         =   "Auto Sign In"
      End
      Begin VB.Menu SigninMnu 
         Caption         =   "Sign In..."
      End
      Begin VB.Menu SignOutMnu 
         Caption         =   "Sign Out"
      End
      Begin VB.Menu MyStatMnu 
         Caption         =   "My status"
         WindowList      =   -1  'True
         Begin VB.Menu OnlineMnu 
            Caption         =   "Online"
            Checked         =   -1  'True
         End
         Begin VB.Menu BusyMnu 
            Caption         =   "Busy"
            Checked         =   -1  'True
         End
         Begin VB.Menu BRBMnu 
            Caption         =   "Be Right Back"
            Checked         =   -1  'True
         End
         Begin VB.Menu AwayMnu 
            Caption         =   "Away"
            Checked         =   -1  'True
         End
         Begin VB.Menu ondafonemnu 
            Caption         =   "On the Phone"
            Checked         =   -1  'True
         End
         Begin VB.Menu gone2lunchmnu 
            Caption         =   "Out to Lunch"
            Checked         =   -1  'True
         End
         Begin VB.Menu Apprearoffmnu 
            Caption         =   "Appear Offline"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu checkemailmnu 
         Caption         =   "My E-mail Inbox"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu savecontactlistmnu 
         Caption         =   "Save Contact List..."
      End
      Begin VB.Menu importcontactsmnu 
         Caption         =   "Import Contacts from a Saved File..."
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu SendFIleMnu 
         Caption         =   "Send File or Photos"
      End
      Begin VB.Menu openreceivedmnu 
         Caption         =   "Open Received Files"
      End
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu closemnu 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu StatusMnu 
      Caption         =   "StatusMnu"
      Visible         =   0   'False
      Begin VB.Menu OnlineMnu2 
         Caption         =   "Online"
         Checked         =   -1  'True
      End
      Begin VB.Menu BusyMnu2 
         Caption         =   "Busy"
         Checked         =   -1  'True
      End
      Begin VB.Menu BRBMnu2 
         Caption         =   "Be Right Back"
         Checked         =   -1  'True
      End
      Begin VB.Menu AwayMnu2 
         Caption         =   "Away"
         Checked         =   -1  'True
      End
      Begin VB.Menu OTPMnu2 
         Caption         =   "On the Phone"
         Checked         =   -1  'True
      End
      Begin VB.Menu OTLMnu2 
         Caption         =   "Out to Lunch"
         Checked         =   -1  'True
      End
      Begin VB.Menu AppearOff2 
         Caption         =   "Appear Offline"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================'
'| Created By: §e7eN                                                  |'
'| Description: This is an Example of MSN API.                        |'
'|              It is still buggy any incomplete but I will           |'
'|              Release updates until I have fully maped the MSN API  |'
'|                                                                    |'
'| Contact: hate_114@hotmail.com                                      |'
'|                                                                    |'
'| *If you wish to use this in one of your Programs please E-mail me* |'
'======================================================================



Public mWindow2 As MessengerAPI.IMessengerConversationWnd
Public MsgrContacts As MessengerAPI.IMessengerContacts
Public MsgrContact As MessengerAPI.IMessengerContact
Public mWindow As MessengerAPI.IMessengerWindow
Public WithEvents MSNAPI As MessengerAPI.Messenger
Attribute MSNAPI.VB_VarHelpID = -1
Dim strServiceID As String
Dim APICheck1 As Boolean
Dim W As Integer

Private Sub AppearOff2_Click()
SetStatus MISTATUS_INVISIBLE
End Sub

Private Sub Apprearoffmnu_Click()
SetStatus MISTATUS_OFFLINE
End Sub

Private Sub autosigninmnu_Click()
AutoSignIn

Timer2.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub AwayMnu_Click()
SetStatus MISTATUS_AWAY
End Sub

Private Sub AwayMnu2_Click()
SetStatus MISTATUS_AWAY
End Sub

Private Sub BigStat_Click()
Me.PopupMenu Me.StatusMnu
End Sub

Private Sub BlockMnu_Click()
Block TV.SelectedItem.Tag
End Sub

Private Sub BRBMnu_Click()
SetStatus MISTATUS_BE_RIGHT_BACK
End Sub

Private Sub BRBMnu2_Click()
SetStatus MISTATUS_BE_RIGHT_BACK
End Sub

Private Sub BusyMnu_Click()
SetStatus MISTATUS_BUSY
End Sub

Private Sub BusyMnu2_Click()
SetStatus MISTATUS_BUSY
End Sub

Private Sub checkemailmnu_Click()
ReadMail
End Sub

Private Sub closemnu_Click()
End
End Sub

Private Sub ConnectBox_Click()

End Sub

Private Sub Form_Load()
Set MSNAPI = New MessengerAPI.Messenger
strServiceID = MSNAPI.MyServiceId
ConStat = "Please Sign in"
GetList
EnableDisableControls
APICheck1 = False

End Sub

Sub EnableDisableControls()
If GetStatus = MISTATUS_OFFLINE Then
SigninMnu.Enabled = True
Me.AutoSignInMnu.Enabled = True
TV.Visible = False
SignOutMnu.Enabled = False
MyStatMnu.Enabled = False
checkemailmnu.Enabled = False
savecontactlistmnu.Enabled = False
importcontactsmnu.Enabled = False
SendFIleMnu.Enabled = False
End If

If Not GetStatus = MISTATUS_OFFLINE Then
SigninMnu.Enabled = False
Me.AutoSignInMnu.Enabled = False
TV.Visible = True
SignOutMnu.Enabled = True
MyStatMnu.Enabled = True
checkemailmnu.Enabled = True
savecontactlistmnu.Enabled = True
importcontactsmnu.Enabled = True
SendFIleMnu.Enabled = True
End If
End Sub

'------------------------------Signing In----------------------------------------
Public Sub AutoSignIn()
If Not MSNAPI.MyStatus = MISTATUS_OFFLINE Then Exit Sub
MSNAPI.AutoSignIn
End Sub

Public Sub Signin()
If MSNAPI.MyStatus = MISTATUS_OFFLINE Then MSNAPI.Signin 0&, "", ""
End Sub

Public Sub signout()
If MSNAPI.MyStatus <> MISTATUS_OFFLINE Then MSNAPI.signout
End Sub

'---------------------------------Status-----------------------------------------
Public Sub SetStatus(Status As MISTATUS)
MSNAPI.MyStatus = Status

End Sub

Public Function GetStatus() As MISTATUS
GetStatus = MSNAPI.MyStatus
End Function

'---------------------------------IM Stuff---------------------------------------
Public Sub InstantMessage(strContactName As String)
        Set MsgrContact = Nothing
        Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)
        Set mWindow = MSNAPI.InstantMessage(MsgrContact)
        'Set mWindow2 = MSNAPI.InstantMessage(strContactName)
        mWindow.Show
'MsgBox mWindow2.History
End Sub


Public Sub SendFile(strContactName As String)

        Set MsgrContact = Nothing
        Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)

        Set mWindow = MSNAPI.SendFile(MsgrContact, strSendFileName)
        
End Sub
Public Sub StartVideo(strContactName)

        Set MsgrContact = Nothing
        Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)
        Set mWindow = MSNAPI.StartVideo(MsgrContact)
        
End Sub

Public Sub RemoteAssistance(strContactName)

        Set MsgrContact = Nothing
        Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)
        Set mWindow = MSNAPI.InviteApp(MsgrContact, "{B90D0115-3AEA-45D3-801E-93913008D49E}")
        '"{44BBA842-CC51-11CF-AAFA-00AA00B6015C}")
        End Sub



Public Sub StartVoice(strContactName)

        Set MsgrContact = Nothing
        Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)
        Set mWindow = MSNAPI.StartVoice(MsgrContact)
        
End Sub

'-------------------------------Email Stuff--------------------------------------
Function GetInboxNum() As Integer
GetInboxNum = CStr(MSNAPI.UnreadEmailCount(MUAFOLDER_INBOX))
End Function

Function GetOtherFolderNum() As Integer
GetOtherFolderNum = CStr(MSNAPI.UnreadEmailCount(MUAFOLDER_ALL_OTHER_FOLDERS))
End Function

Public Sub SendEmail(strContactName As String)
        Set MsgrContact = Nothing
        Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)
        MSNAPI.SendMail MsgrContact
End Sub

Public Sub ReadMail()
MSNAPI.OpenInbox
End Sub

'---------------------------------Profiles-----------------------------------
Public Sub ViewProfile(strContactName)
        Set MsgrContact = Nothing
        Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)
        MSNAPI.ViewProfile MsgrContact
End Sub

Public Sub Block(strContactName)
    Set MsgrContact = Nothing
    Set MsgrContact = MSNAPI.GetContact(strContactName, strServiceID)
    
    If MsgrContact.Blocked = True Then
        MsgrContact.Blocked = False
        MsgBox ("Contact: " & CStr(MsgrContact.SignInName) & " is now Unblocked")
    Else
        MsgrContact.Blocked = True
        MsgBox ("Contact: " & CStr(MsgrContact.SignInName) & " is now Blocked")
    End If
    GetList
End Sub

Sub GetList()
Dim MsgrContacts As IMessengerContacts
Dim Stat As MISTATUS
Dim OnlineUsers, OfflineUsers As Integer
TV.Visible = False
Timer3.Enabled = True
If MSNAPI.MyStatus = MISTATUS_OFFLINE Then Exit Sub
If MSNAPI.MyStatus >= 66 Then Exit Sub
Set MsgrContacts = MSNAPI.MyContacts
For x = 1 To MsgrContacts.Count - 1
Stat = MsgrContacts.Item(x).Status
If Stat = MISTATUS_OFFLINE Then OfflineUsers = OfflineUsers + 1 Else OnlineUsers = OnlineUsers + 1
Next

TV.Nodes.Clear
If GetInboxNum = 0 Then
TV.Nodes.Add , , "Email", "No new e-mail messages", 7
Else
TV.Nodes.Add , , "Email", GetInboxNum & " Unread Messages)", 7
End If

TV.Nodes.Add , , "online", "Online(" & OnlineUsers & ")", 6
TV.Nodes.Add , , "offline", "Offline (" & OfflineUsers & ")", 6
TV.Nodes.Item(2).Expanded = True
TV.Nodes.Item(3).Expanded = True
TV.Nodes.Item(2).Bold = True
TV.Nodes.Item(3).Bold = True
TV.Nodes.Item(2).ForeColor = &H8000000D
TV.Nodes.Item(3).ForeColor = &H8000000D
TV.Nodes.Item(2).Sorted = True
TV.Nodes.Item(3).Sorted = True


For x = 1 To MsgrContacts.Count - 1
AddNode MsgrContacts.Item(x).FriendlyName, MsgrContacts.Item(x).SignInName, MsgrContacts(x).Status, MsgrContacts.Item(x).Blocked
Next
UserName = MSNAPI.MyFriendlyName
TV.Visible = True
Timer3.Enabled = True
End Sub


Public Sub DeleteNode(ByVal name As String)
Dim temp As String
For x = 1 To TV.Nodes.Count - 1
    If TV.Nodes.Item(x).Text = name Then TV.Nodes.Remove (x)
Next x
End Sub

Private Sub AddNode(ByVal FriendlyName As String, SignInName As String, ByVal stats As MISTATUS, Blocked As Boolean)
Dim Node As Node

If Blocked = True And stats = MISTATUS_OFFLINE Then
Set Node = TV.Nodes.Add("offline", tvwChild, , FriendlyName, 8)
Node.Tag = SignInName
Exit Sub
End If
If Blocked = True And Not stats = MISTATUS_OFFLINE Then
Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName & " (Blocked)", 8)
Node.Tag = SignInName
Exit Sub
End If

If stats = MISTATUS_OFFLINE Then Set Node = TV.Nodes.Add("offline", tvwChild, , FriendlyName, 2)
If stats = 18 Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName & " (Away)", 3)
If stats = MISTATUS_BE_RIGHT_BACK Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName & " (Be Right Back)", 3)
If stats = MISTATUS_BUSY Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName & " (Busy)", 4)
If stats = 34 Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName & " (Idle)", 3)
If stats = MISTATUS_INVISIBLE Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName, 9)
If stats = MISTATUS_ONLINE Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName, 1)
If stats = MISTATUS_ON_THE_PHONE Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName & "(On the Phone)", 3)
If stats = MISTATUS_OUT_TO_LUNCH Then Set Node = TV.Nodes.Add("online", tvwChild, , FriendlyName & " (Out to Lunch)", 4)
'If Node = Nothing Then AddNode FriendlyName, SignInName, stats
Node.Tag = SignInName
End Sub

Public Function NodeIndex(ByVal name As String) As Integer
Dim temp As String
For x = 1 To TV.Nodes.Count - 1
    If TV.Nodes.Item(x).Text = name Then NodeIndex = x
Next x
End Function

Private Sub Form_Resize()
UserName.Width = Me.Width - UserName.Left - Stat.Width - 50

Image1.Left = (Me.Width / 2) - (Image1.Width / 2)
ConStat.Left = (Me.Width / 2) - (ConStat.Width / 2)

TV.Width = Me.Width - 250
TV.Height = Me.Height - TV.Top
End Sub

Private Sub gone2lunchmnu_Click()
SetStatus MISTATUS_OUT_TO_LUNCH
End Sub

Private Sub importcontactsmnu_Click()
MsgBox "Currently Unavaliable", , Me.Caption
Exit Sub
End Sub

Private Sub MSNAPI_OnContactBlockChange(ByVal hr As Long, ByVal pContact As Object, ByVal pBoolBlock As Boolean)
GetList
End Sub

Private Sub MSNAPI_OnContactFriendlyNameChange(ByVal hr As Long, ByVal pMContact As Object, ByVal bstrPrevFriendlyName As String)
GetList
End Sub

Private Sub MSNAPI_OnContactStatusChange(ByVal pMContact As Object, ByVal mStatus As MessengerAPI.MISTATUS)
GetList
End Sub

Private Sub MSNAPI_OnMyFriendlyNameChange(ByVal hr As Long, ByVal bstrPrevFriendlyName As String)
GetList
End Sub

Private Sub MSNAPI_OnMyStatusChange(ByVal hr As Long, ByVal mMyStatus As MessengerAPI.MISTATUS)
EnableDisableControls
GetList
End Sub

Private Sub MSNAPI_OnUnreadEmailChange(ByVal mFolder As MessengerAPI.MUAFOLDER, ByVal cUnreadEmail As Long, pBoolfEnableDefault As Boolean)
GetList
End Sub

Private Sub ondafonemnu_Click()
SetStatus MISTATUS_ON_THE_PHONE
End Sub

Private Sub OnlineMnu_Click()
SetStatus MISTATUS_ONLINE
End Sub

Private Sub OnlineMnu2_Click()
SetStatus MISTATUS_ONLINE
End Sub

Private Sub openreceivedmnu_Click()
MsgBox "Currently Unavaliable", , Me.Caption
Exit Sub
End Sub

Private Sub OTLMnu2_Click()
SetStatus MISTATUS_OUT_TO_LUNCH
End Sub

Private Sub OTPMnu2_Click()
SetStatus MISTATUS_ON_THE_PHONE
End Sub

Private Sub RemoteAssistanceMnu_Click()
MsgBox "Currently Unavaliable", , Me.Caption
Exit Sub
If TV.SelectedItem.Tag = "" Then Exit Sub
RemoteAssistance TV.SelectedItem.Tag
End Sub

Private Sub savecontactlistmnu_Click()
MsgBox "Currently Unavaliable", , Me.Caption
Exit Sub
End Sub

Private Sub SendFIleMnu_Click()
MsgBox "Currently Unavaliable", , Me.Caption
Exit Sub
End Sub

Private Sub SendFIleMnu2_Click(Index As Integer)
If TV.SelectedItem.Tag = "" Then Exit Sub
Me.SendFile TV.SelectedItem.Tag
End Sub
 
Private Sub SendIMMnu_Click()
If TV.SelectedItem.Tag = "" Then Exit Sub
InstantMessage TV.SelectedItem.Tag
End Sub



Private Sub SigninMnu_Click()
Signin
End Sub

Private Sub signoutmnu_Click()
signout
End Sub

Private Sub StartAppMnu_Click()
MsgBox "Currently Unavaliable", , Me.Caption
Exit Sub
End Sub

Private Sub StartVideoMnu_Click()
If TV.SelectedItem.Tag = "" Then Exit Sub
Me.StartVideo TV.SelectedItem.Tag
End Sub

Private Sub StartVoiceMnu_Click()
If TV.SelectedItem.Tag = "" Then Exit Sub
Me.StartVoice TV.SelectedItem.Tag
End Sub

Private Sub StartWhiteboardMnu_Click()
MsgBox "Currently Unavaliable", , Me.Caption
Exit Sub
End Sub

Private Sub Timer1_Timer()
Image1.Picture = PC.GraphicCell(W)
W = W + 1
If W = 20 Then W = 0
End Sub

Private Sub Timer2_Timer()

Select Case GetStatus
Case MISTATUS_LOCAL_CONNECTING_TO_SERVER
Timer1.Enabled = True
Image1.Visible = True
ConStat.Caption = "Connecting.."
Case MISTATUS_LOCAL_FINDING_SERVER
ConStat.Caption = "Locateing Server.."
Case MISTATUS_LOCAL_SYNCHRONIZING_WITH_SERVER
ConStat.Caption = "Synchronizeing" & vbCrLf & "With Server.."
Case MISTATUS_ONLINE
Timer1.Enabled = False
Image1.Visible = False
Timer2.Enabled = False
End Select
End Sub

Private Sub Timer3_Timer()
Select Case GetStatus
Case MISTATUS_ONLINE
BigStat.Picture = BigStatIco.ListImages(1).Picture
Case MISTATUS_AWAY, MISTATUS_BE_RIGHT_BACK, MISTATUS_IDLE, MISTATUS_OUT_TO_LUNCH
BigStat.Picture = BigStatIco.ListImages(3).Picture
Case MISTATUS_BUSY, MISTATUS_ON_THE_PHONE
BigStat.Picture = BigStatIco.ListImages(2).Picture
Case MISTATUS_INVISIBLE
BigStat.Picture = BigStatIco.ListImages(4).Picture
End Select

Me.OnlineMnu.Checked = False
Me.OnlineMnu2.Checked = False
Me.BRBMnu.Checked = False
Me.BRBMnu2.Checked = False
Me.BusyMnu.Checked = False
Me.BusyMnu2.Checked = False
Me.AwayMnu.Checked = False
Me.AwayMnu2.Checked = False
Me.AppearOff2.Checked = False
Me.Apprearoffmnu.Checked = False
Me.OTLMnu2.Checked = False
Me.gone2lunchmnu.Checked = False
Me.OTPMnu2.Checked = False
Me.ondafonemnu.Checked = False

If GetStatus = MISTATUS_AWAY Then
Me.AwayMnu.Checked = True
Me.AwayMnu2.Checked = True
End If

If GetStatus = MISTATUS_BE_RIGHT_BACK Then
Me.BRBMnu.Checked = True
Me.BRBMnu2.Checked = True
End If

If GetStatus = MISTATUS_BUSY Then
Me.BusyMnu.Checked = True
Me.BusyMnu2.Checked = True
End If

If GetStatus = MISTATUS_INVISIBLE Then
Me.AppearOff2.Checked = True
Me.Apprearoffmnu.Checked = False
End If

If GetStatus = MISTATUS_ON_THE_PHONE Then
Me.OTPMnu2.Checked = True
Me.ondafonemnu.Checked = True
End If

If GetStatus = MISTATUS_OUT_TO_LUNCH Then
Me.OTLMnu2.Checked = True
Me.gone2lunchmnu.Checked = True
End If

If GetStatus = MISTATUS_ONLINE Then
Me.OnlineMnu.Checked = True
Me.OnlineMnu2.Checked = True
End If
End Sub


Private Sub TV_Click()
If Mid(TV.SelectedItem.Text, 1, 6) = "Online" Then
If TV.Nodes.Item(2).Expanded = True Then
TV.Nodes.Item(2).Image = 5
TV.Nodes.Item(2).Expanded = False
Else
TV.Nodes.Item(2).Image = 6
TV.Nodes.Item(2).Expanded = True
End If
End If

If Mid(TV.SelectedItem.Text, 1, 7) = "Offline" Then
If TV.Nodes.Item(3).Expanded = True Then
TV.Nodes.Item(3).Image = 5
TV.Nodes.Item(3).Expanded = False
Else
TV.Nodes.Item(3).Image = 6
TV.Nodes.Item(3).Expanded = True
End If
End If
End Sub

Private Sub TV_DblClick()

If Mid(TV.SelectedItem.Text, 1, 7) = "Offline" Then Exit Sub
If Mid(TV.SelectedItem.Text, 1, 6) = "Online" Then Exit Sub
InstantMessage TV.SelectedItem.Tag
        
End Sub

Private Sub TV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Mouse.LeftClick
Me.SentEmailToMnu.Caption = "Send E-mail (" & TV.SelectedItem.Tag & ")"
    Set MsgrContact = Nothing
    Set MsgrContact = MSNAPI.GetContact(TV.SelectedItem.Tag, strServiceID)
If MsgrContact.Blocked = True Then Me.BlockMnu.Caption = "Unblock" Else Me.BlockMnu.Caption = "Block"
Me.PopupMenu Me.RightMnu
End If

End Sub


