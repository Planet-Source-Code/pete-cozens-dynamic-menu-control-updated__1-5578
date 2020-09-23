VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{28D177B5-F05E-11D3-AEFF-08005AD29D41}#1.0#0"; "DynaMenu.ocx"
Begin VB.Form Form1 
   Caption         =   "Dynamic Menu Test"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin DynaMenus.DynaMenu DynaMenu1 
      Left            =   2640
      Top             =   2400
      _ExtentX        =   1058
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdClearMenu 
      Caption         =   "Clear Menu"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveMenu 
      Caption         =   "Save Menu"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoadMenu 
      Caption         =   "Load Menu"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "FOLDER1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0452
            Key             =   "FOLDER2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08A4
            Key             =   "ITEM1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CF6
            Key             =   "ITEM2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7011
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDM 
         Caption         =   "DynaMenu"
         Begin VB.Menu mnuDMC 
            Caption         =   "DynaMenuChild"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "-Popup Menu-"
      Visible         =   0   'False
      Begin VB.Menu mnuAddMenuItem 
         Caption         =   "Add Menu Item"
      End
      Begin VB.Menu mnuAddPopupMenu 
         Caption         =   "Add Popup Menu"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_CHECKMARKS = &H8
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Const MIIM_ALL = MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU Or MIIM_CHECKMARKS Or MIIM_TYPE Or MIIM_DATA
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuAPI Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long



Private Const INSERT_AT_INDEX = 0   ' Append menus
'Private Const INSERT_AT_INDEX = 1   ' Insert new menu items @ start

Private m_lngNextID As Long

'*******************************************************************************
' Initialise the Dynamic Menus and the TreeView
'-------------------------------------------------------------------------------

Private Sub Form_Load()
    
    ' Important - All three of these menu properties must be set for the
    ' control to work correctly
    Set DynaMenu1.ParentMenu = mnuView
    Set DynaMenu1.PopupMenu = mnuDM
    Set DynaMenu1.ChildMenuArray = mnuDMC     ' Must Be an Array
    
    ' Initialise the TreeView (the menus will already be empty)
    Clear
    
'    Dim hMenu As Long
'    hMenu = GetMenuAPI(Me.hwnd)
'    Dim miView As MENUITEMINFO
'    miView.cbSize = Len(miView)
'    miView.fMask = MIIM_ALL
'    Debug.Assert GetMenuItemInfo(hMenu, 1, 1, miView)
'    Debug.Assert IsMenu(miView.hSubMenu)
'    MsgBox "View Menu Handle = " & miView.hSubMenu

    cmdLoadMenu_Click
End Sub

'*******************************************************************************
' Menu event handler
'-------------------------------------------------------------------------------

Private Sub mnuDMC_Click(Index As Integer)

    ' Determine the CMenu item that was clicked on
    ' NOTE - the index of the CMenu does not necessarily correspond with
    ' the index of the VB menu item used by the dynamic menu, and therefore
    ' can not be used to access DynaMenu's Menu collection.
    
    Dim mnu As CMenu
    Set mnu = DynaMenu1.ItemByMenuIndex(Index)
    
    If Not mnu Is Nothing Then
    
        ' Having determined which menu item was clicked on, do something
        ' useful depending on the key of the menu item.
    
        MsgBox "You clicked on " & mnu.Caption & vbCrLf & _
                            "Internal Key = " & mnu.Key
        Set mnu = Nothing
    Else
        MsgBox "Error - unrecognised menu item !"
    End If
End Sub

'*******************************************************************************
' Basic Infrastructure of the Form - irrelevant to the operation of the
' dynamic menus
'-------------------------------------------------------------------------------

Private Function Max(ByVal a As Long, ByVal b As Long) As Long
    Max = IIf(a >= b, a, b)
End Function

Private Sub mnuAbout_Click()
    MsgBox "About Box"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuViewOptions_Click()
    MsgBox "Options Dialog"
End Sub

Private Sub Form_Resize()
    Dim lng As Long
    lng = Max(ScaleHeight - cmdLoadMenu.Height - TreeView1.Left, 0)
    cmdLoadMenu.Top = lng
    cmdSaveMenu.Top = lng
    cmdClearMenu.Top = lng
    TreeView1.Width = Max(ScaleWidth - 2 * TreeView1.Left, 0)
    TreeView1.Height = Max(cmdLoadMenu.Top - TreeView1.Left, 0)
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
    If TreeView1.SelectedItem.Key = "!0" Then Cancel = 1
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    DynaMenu1.Menu(TreeView1.SelectedItem.Key).Caption = NewString
End Sub

'*******************************************************************************
' Trap the mouse down event to display the context menu
'-------------------------------------------------------------------------------

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button And vbRightButton) Then
        Dim n As Node
        Set n = TreeView1.HitTest(x, y)
        If Not n Is Nothing Then
            Set TreeView1.SelectedItem = n
            If n.Key = "!0" Then
                mnuAddMenuItem.Enabled = True
                mnuAddPopupMenu.Enabled = True
                mnuDelete = False
                mnuRename = 0
            Else
                Dim m As CMenu
                Set m = DynaMenu1.Menu(n.Key)
                mnuAddMenuItem.Enabled = m.IsPopup
                mnuAddPopupMenu.Enabled = m.IsPopup
                mnuDelete = True
                mnuRename = True
            End If
            Me.PopupMenu mnuPopup
        End If
    End If
End Sub

'*******************************************************************************
' Context menu event handlers
'-------------------------------------------------------------------------------

Private Sub mnuAddMenuItem_Click()
    ' Add a menu item to the tree and menu
    ' Make sure that you use a meaningful key!
    m_lngNextID = m_lngNextID + 1
    AddMenuItem "!" & m_lngNextID, TreeView1.SelectedItem.Key, "New Menu Item"
    ' Edit the default caption of the menu item
    Set TreeView1.SelectedItem = TreeView1.Nodes("!" & m_lngNextID)
    TreeView1.StartLabelEdit
End Sub

Private Sub mnuAddPopupMenu_Click()
    ' Add a popup menu to the tree and menu
    ' Make sure that you use a meaningful key!
    m_lngNextID = m_lngNextID + 1
    AddPopupMenu "!" & m_lngNextID, TreeView1.SelectedItem.Key, "New Popup Menu"
    ' Edit the default caption of the popup
    Set TreeView1.SelectedItem = TreeView1.Nodes("!" & m_lngNextID)
    TreeView1.StartLabelEdit
End Sub

Private Sub mnuDelete_Click()
    ' Delete the selected menu item / popup menu
    DeleteMenu TreeView1.SelectedItem.Key
End Sub

Private Sub mnuRename_Click()
    ' Rename the menu item / popup menu
    TreeView1.StartLabelEdit
End Sub

'*******************************************************************************
' Add/Remove items to/from both the TreeView and the Dynamic Menu
'-------------------------------------------------------------------------------

Private Sub AddMenuItem(Key As String, ParentKey As String, Caption As String)

    'Add a node to the tree
    
    TreeView1.Nodes.Add ParentKey, tvwChild, Key, Caption, "ITEM1", "ITEM2"
    TreeView1.Nodes(ParentKey).Expanded = True
    
    'Add an item to the menu
    
    If ParentKey <> "!0" Then
        DynaMenu1.Menu.Add Caption, Key, INSERT_AT_INDEX, ParentKey, False
    Else
        DynaMenu1.Menu.Add Caption, Key, INSERT_AT_INDEX, , False
    End If
    
End Sub

Private Sub AddPopupMenu(Key As String, ParentKey As String, Caption As String)

    'Add a node to the tree
    
    TreeView1.Nodes.Add ParentKey, tvwChild, Key, Caption, "FOLDER1", "FOLDER2"
    TreeView1.Nodes(ParentKey).Expanded = True
    
    'Add an item to the menu
    
    If ParentKey <> "!0" Then
        DynaMenu1.Menu.Add Caption, Key, INSERT_AT_INDEX, ParentKey, True
    Else
        DynaMenu1.Menu.Add Caption, Key, INSERT_AT_INDEX, , True
    End If
    
End Sub

Private Sub DeleteMenu(Key As String)
    TreeView1.Nodes.Remove Key
    DynaMenu1.Menu.Remove Key
End Sub

'*******************************************************************************
' File I/O for saving the menus
'-------------------------------------------------------------------------------

Private Sub cmdLoadMenu_Click()

    Clear

    Dim File As Integer
    File = FreeFile
    Open App.Path & "\MenuFile.txt" For Input As #File
    
    Dim strKey As String, strParentKey As String
    Dim strCaption As String, blnPopup As Boolean
    
    Do While EOF(File) = False
        Input #File, strKey, strParentKey, strCaption, blnPopup
        If (Len(strKey) > 0) Then
            If strParentKey = "" Then strParentKey = "!0"
            If blnPopup Then
                AddPopupMenu strKey, strParentKey, strCaption
            Else
                AddMenuItem strKey, strParentKey, strCaption
            End If
            
            m_lngNextID = Max(m_lngNextID, CLng(Mid$(strKey, 2)))
        End If
    Loop
    
    Close #File

End Sub

Private Sub cmdSaveMenu_Click()
    
    Dim File As Integer
    File = FreeFile
    Open App.Path & "\MenuFile.txt" For Output As #File
    
    Dim m As CMenu
    For Each m In DynaMenu1.Menu
        Write #File, m.Key, m.ParentKey, m.Caption, m.IsPopup
    Next
    
    Close #File

End Sub

'*******************************************************************************
' Clear all items from both the TreeView and the Dynamic Menu
'-------------------------------------------------------------------------------

Private Sub cmdClearMenu_Click()
    Clear
End Sub

Private Sub Clear()
    ' Clear the Menus
    TreeView1.Nodes.Clear
    DynaMenu1.Menu.Clear
    TreeView1.Nodes.Add , , "!0", "Root", "FOLDER1", "FOLDER2"
    m_lngNextID = 0
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------

