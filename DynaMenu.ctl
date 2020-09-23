VERSION 5.00
Begin VB.UserControl DynaMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   570
   ScaleWidth      =   1560
   ToolboxBitmap   =   "DynaMenu.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "DynaMenu.ctx":00FA
      ScaleHeight     =   450
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "DynaMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_objParentMenu As Menu
Private m_objPopupMenu As Menu
Private m_varChildMenuArray As Variant
Private m_colMenus As CMenus

Private Const DEF_CAPTION As String = ""

#Const DEBUG_MODE = False

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    Set m_colMenus = New CMenus
    Set m_colMenus.Parent = Me
End Sub

Private Sub UserControl_Terminate()
    Set m_colMenus.Parent = Nothing
    Set m_colMenus = Nothing
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Private Sub UserControl_Resize()
    Picture1.Move 60, 60
    UserControl.Width = Picture1.Width + 120
    UserControl.Height = Picture1.Height + 120
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Private Sub UserControl_InitProperties()
    'TODO...
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'TODO...
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'TODO...
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Public Property Get ParentMenu() As Variant
    Set ParentMenu = m_objParentMenu
End Property
Public Property Set ParentMenu(vNewValue As Variant)
    If TypeOf vNewValue Is Menu Then
        Set m_objParentMenu = vNewValue
    Else
        Err.Raise vbObjectError, "DynaMenu::ParentMenu() [Set]", _
            "DynaMenu::ParentMenu() [Set] - ParentMenu can only be set " & _
            "to a VB Menu Object"
    End If
    
    If IsMenu(GethMenu(m_objParentMenu)) = 0 Then
        Timer1.Interval = 10
        Timer1.Enabled = True
    End If
        
End Property

Public Property Get PopupMenu() As Variant
    Set PopupMenu = m_objPopupMenu
End Property
Public Property Set PopupMenu(vNewValue As Variant)
    Set m_objPopupMenu = vNewValue
    If TypeOf vNewValue Is Menu Then
        Set m_objPopupMenu = vNewValue
        On Error Resume Next
        m_objPopupMenu.Visible = False
    Else
        Err.Raise vbObjectError, "DynaMenu::PopupMenu() [Set]", _
            "DynaMenu::PopupMenu() [Set] - PopupMenu can only be set " & _
            "to a VB Menu Object"
    End If

End Property

Public Property Get ChildMenuArray() As Variant
    Set ChildMenuArray = m_varChildMenuArray
End Property
Public Property Set ChildMenuArray(vNewValue As Variant)
    Set m_varChildMenuArray = vNewValue
End Property

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Public Property Get Menu() As CMenus
    Set Menu = m_colMenus
End Property

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Public Function ItemByMenuIndex(ByVal Index As Long) As CMenu

    On Error GoTo ErrorTrap
    Dim obj As CMenu
'    Set obj = mCol(m_varMenuItem(Index).Tag)
    For Each obj In m_colMenus
        If (obj.MenuItem.Index = Index) Then
            Set ItemByMenuIndex = obj
            Exit For
        End If
    Next
    Set obj = Nothing
    
    If (ItemByMenuIndex Is Nothing) Then
        MsgBox "No Menu Item found with MenuItemIndex=" & Index
    End If
    
ErrorTrap:
End Function

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Public Sub Refresh()
    DrawMenuBar hParentWnd
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Friend Property Get hParentWnd() As Long
    hParentWnd = UserControl.Parent.hwnd
End Property

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Private Sub Timer1_Timer()
    
    If IsMenu(GethMenu(m_objParentMenu)) = 0 Then Exit Sub
    If IsMenu(GethMenu(m_objPopupMenu)) = 0 Then Exit Sub
    
    Timer1.Enabled = False

    Dim mnu As CMenu
    For Each mnu In m_colMenus
        If Len(mnu.ParentKey) = 0 Then
            AddMenuObject mnu
        End If
    Next
End Sub

Friend Sub AddMenuObject(Menu As CMenu)

    Dim hParentMenu As Long
    If Len(Menu.ParentKey) > 0 Then
        ' Belongs to SubMenu
        Dim mnu As CMenu
        Set mnu = m_colMenus(Menu.ParentKey)
        hParentMenu = mnu.hPopupMenu
        
        'Increment  the parent's ChildCount property
        mnu.ChildCount = mnu.ChildCount + 1
        Set mnu = Nothing
    
    Else
        ' Belongs to Parent Menu
        If ParentMenu Is Nothing Then
            hParentMenu = GetMenu(hParentWnd)
        Else
            hParentMenu = GethMenu(ParentMenu)
        End If
    End If
    
#If DEBUG_MODE Then
    If IsMenu(hParentMenu) = 0 Then
        MsgBox "DynaMenu_AddMenuObject() - Invalid hParentMenu"
    End If
#End If

    Menu.hMenu = hParentMenu
    
    If Menu.IsPopup Then
        AddSubMenu Menu, hParentMenu
    Else
        AddMenuItem Menu, hParentMenu
    End If
    Menu.Caption = Menu.Caption     ' Update physical menu
    Menu.Checked = Menu.Checked
    Menu.Enabled = Menu.Enabled
    
    ' If Index=0 then set index to its true value
    
    If Menu.Index = 0 Then
        Dim mnuTmp As CMenu
        For Each mnuTmp In m_colMenus
            If mnuTmp.ParentKey = Menu.ParentKey Then
                Menu.Index = Menu.Index + 1
            End If
        Next
        Set mnuTmp = Nothing
    End If
End Sub

Private Sub AddMenuItem(Menu As CMenu, hParentMenu As Long)
    
    Dim mnu As Menu
    Set mnu = Menu.MenuItem
    
    Menu.ItemID = GetCommand(mnu)
    Dim lngRet As Long
    If (Menu.Index < 1) Then
        lngRet = AppendMenu(hParentMenu, MF_STRING Or MF_BYCOMMAND, _
                                                Menu.ItemID, DEF_CAPTION)
    Else
        lngRet = InsertMenu(hParentMenu, Menu.Index - 1, _
                        MF_STRING Or MF_BYPOSITION, Menu.ItemID, DEF_CAPTION)
    End If
    Debug.Assert lngRet

#If DEBUG_MODE Then
    If (lngRet = 0) Then
        MsgBox "DynaMenu_AddMenuItem() - Failed to Insert/Append Menu Item"
    End If
#End If
End Sub

Private Sub AddSubMenu(Menu As CMenu, hParentMenu As Long)
    
    Dim mnu As Menu
    Set mnu = Menu.MenuItem
    
    Menu.hPopupMenu = CreatePopupMenu()

    Dim lngRet As Long
    If (Menu.Index = 0) Then
        lngRet = AppendMenu(hParentMenu, MF_POPUP, Menu.hPopupMenu, DEF_CAPTION)
    Else
        lngRet = InsertMenu(hParentMenu, Menu.Index - 1, _
                 MF_POPUP Or MF_BYPOSITION, Menu.hPopupMenu, DEF_CAPTION)
    End If
    
#If DEBUG_MODE Then
    If (lngRet = 0) Then
        MsgBox "DynaMenu_AddSubMenu() - Failed to Insert/Append Menu Item"
    End If
#End If
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------

