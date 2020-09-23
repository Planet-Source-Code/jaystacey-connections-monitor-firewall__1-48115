VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Connections Monitor"
   ClientHeight    =   4500
   ClientLeft      =   450
   ClientTop       =   1140
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   11445
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4245
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5760
      Top             =   3360
   End
   Begin VB.PictureBox pic16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2820
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   2280
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   3480
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   4080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "iml32"
      SmallIcons      =   "iml16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Direction"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote Host"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "File Path"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Refreshlist 
         Caption         =   "Refresh"
      End
      Begin VB.Menu ViewCon 
         Caption         =   "View Connections"
      End
      Begin VB.Menu seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu ShowPop 
         Caption         =   "Show Popup"
         Checked         =   -1  'True
      End
      Begin VB.Menu AutoFresh 
         Caption         =   "Automatic Refresh"
         Checked         =   -1  'True
      End
      Begin VB.Menu seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu MinSystray 
         Caption         =   "Minimise to System tray"
      End
      Begin VB.Menu ExitProg 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO
Public Tablenum As Long
Private pTcpTable As MIB_TCPTABLE

Public Sub RefreshView()
  Dim i As Integer, o As Integer
  Dim fileNum As String
  Dim Item As ListItem
  
On Error Resume Next
  ListView1.ListItems.Clear
   
    ListView1.Icons = Nothing
    ListView1.SmallIcons = Nothing
    iml32.ListImages.Clear
    iml16.ListImages.Clear
    
    DoEvents
  LoadProcesses
  
  StatusBar1.Panels(1).Text = Connection(1).LocalHost
  StatusBar1.Panels(2).Text = "Last Refresh - " & Time
  
  For i = 0 To StatsLen - 1
  
  If Connection(i).FileName <> "" Then Set Item = ListView1.ListItems.Add(, , Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\"))) Else Set Item = ListView1.ListItems.Add(, , "Unknown")
    
    If Connection(i).LocalPort = Connection(i).RemotePort And Connection(i).LocalPort <> "" Then Item.SubItems(1) = "Incomming" Else Item.SubItems(1) = "Outgoing"
    Item.SubItems(2) = Connection(i).LocalPort
    Item.SubItems(3) = Connection(i).RemoteHost
    Item.SubItems(4) = Connection(i).RemotePort
    Item.SubItems(5) = Connection(i).State
    Item.SubItems(6) = Connection(i).FileName
    
    'Item.EnsureVisible
    
  Next
  
  DoEvents
  GetAllIcons
  ShowIcons

  DoEvents
    Me.MousePointer = vbNormal
    'Label2.Caption = "Netstat status as of: " & Date & " " & Time
End Sub

Private Sub AutoFresh_Click()
If AutoFresh.Checked = True Then
Timer1.Enabled = False
AutoFresh.Checked = False
Else
AutoFresh.Checked = True
Timer1.Enabled = True
End If
End Sub

Private Sub ExitProg_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Visible = False
pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY

ShowPop.Checked = False
DoEvents
RefreshView

End Sub

Private Sub Form_Resize()

On Error Resume Next
ListView1.Width = Me.Width - 150
ListView1.Left = 0
ListView1.Height = Me.Height - 1050

ListView1.ColumnHeaders(1).Width = 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders(2).Width = 1100
ListView1.ColumnHeaders(3).Width = 1100
ListView1.ColumnHeaders(4).Width = ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders(5).Width = 1100
ListView1.ColumnHeaders(6).Width = 1300
ListView1.ColumnHeaders(7).Width = ListView1.Width \ 2 + 1000

End Sub

Private Sub MinSystray_Click()
ShowIcon Me, "Connection Monitor (Monitoring)"
End Sub

Private Sub Refreshlist_Click()
RefreshView
End Sub

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView1
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

On Local Error Resume Next
For Each Item In ListView1.ListItems
  FileName = Item.SubItems(Item.ListSubItems.Count) ' & Item.Text
  GetIcon FileName, Item.Index
Next

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long

'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hdc, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function

Private Sub ShowPop_Click()
If ShowPop.Checked = True Then
ShowPopup = False
ShowPop.Checked = False
Else
ShowPopup = True
ShowPop.Checked = True
End If
End Sub

Private Sub Timer1_Timer()
'RefreshView
 
Dim pdwSize As Long
Dim bOrder As Long
Dim nRet As Long
Dim TableLen As Long

nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)

TableLen = pTcpTable.dwNumEntries
If Tablenum <> TableLen Then RefreshView
Tablenum = TableLen

End Sub

Private Sub ViewCon_Click()

ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add 1, , "File", 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders.Add 2, , "Direction", 1000, lvwColumnCenter
ListView1.ColumnHeaders.Add 3, , "Local Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 4, , "Remote Host", ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders.Add 5, , "Remote Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 6, , "Status", 1300, lvwColumnCenter
ListView1.ColumnHeaders.Add 7, , "File Path", ListView1.Width \ 2 + 1000

RefreshView
End Sub

