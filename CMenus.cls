VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CMenu"
Attribute VB_Ext_KEY = "Member0" ,"CMenu"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

#Const DEBUG_MODE = False

Private mCol As Collection
Private m_objParent As DynaMenu

Friend Property Get Parent() As DynaMenu
    Set Parent = m_objParent
End Property
Friend Property Set Parent(vNewValue As DynaMenu)
    Set m_objParent = vNewValue
End Property

Public Function Add(ByVal Caption As String, _
                    ByVal Key As String, _
                    Optional ByVal Index As Long, _
                    Optional ByVal ParentKey As String, _
                    Optional ByVal IsSubMenu As Boolean = False) As CMenu
    
    If m_objParent.PopupMenu Is Nothing Then
        Err.Raise vbObjectError, "CMenus::Add()", _
            "CMenus::Add() Menu Items can not be added until the " & _
            "DynaMenu's PopupMenu property has been set"
        Exit Function
    End If
    
    If IsNull(m_objParent.ChildMenuArray) Then
        Err.Raise vbObjectError, "CMenus::Add()", _
            "CMenus::Add() Menu Items can not be added until the " & _
            "DynaMenu's ChildMenuArray property has been set"
        Exit Function
    End If
    
    If Len(Key) = 0 Then
        Err.Raise vbObjectError, "CMenus::Add()", _
            "CMenus::Add() Menu Item must have a Key"
        Exit Function
    End If
    
    ' Add a CMenu object to the collection
    
    Dim objNewMember As CMenu
    Set objNewMember = New CMenu
    objNewMember.Key = Key
    objNewMember.ParentKey = ParentKey
    objNewMember.IsPopupMenu = IsSubMenu
    objNewMember.Caption = Caption
    objNewMember.Index = Index
    Set objNewMember.MenuItem = NewMenuItem
    mCol.Add objNewMember, Key
    
    ' Add an item the the visible menu

    m_objParent.AddMenuObject objNewMember
    
    ' Update the indexes of other menu objects
    
    If (Index > 0) Then
        Dim mnu As CMenu
        For Each mnu In m_objParent.Menu
            If (mnu.ParentKey = ParentKey) Then
                If (mnu.Index > Index) Then
                    mnu.Index = mnu.Index + 1
                End If
            End If
        Next
        Set mnu = Nothing
    End If
    
    ' Return the newly added object
    
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As CMenu
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)

    Dim mnu As CMenu
    Set mnu = mCol(vntIndexKey)
    
    ' Remove all child menu items
    
    If mnu.IsPopup Then
        Dim tmp As CMenu
        For Each tmp In mCol
            If tmp.ParentKey = mnu.Key Then
                Remove tmp.Key
            End If
        Next
        Set tmp = Nothing
    End If
    
    ' Decrement the parent menu item's ChildCount property
    
    If (Len(mnu.ParentKey) <> 0) Then
        Dim mnuParent As CMenu
        Set mnuParent = mCol(mnu.ParentKey)
        mnuParent.ChildCount = mnuParent.ChildCount - 1
    End If
    
    ' Delete the Win32 Menu item
    
    Dim lngRet As Long
    If mnu.IsPopup Then
        lngRet = DestroyMenu(mnu.hPopupMenu)
        lngRet = DeleteMenu(mnu.hMenu, mnu.hPopupMenu, MF_BYCOMMAND)
    Else
        lngRet = DeleteMenu(mnu.hMenu, mnu.ItemID, MF_BYCOMMAND)
    End If
#If DEBUG_MODE Then
    If (lngRet = 0) Then
        MsgBox "CMenus_Remove() - Failed to DeleteMenu"
    End If
#End If

    ' Delete the associated VB menu item
    
    Unload mnu.MenuItem
    
    ' Update the indexes of other menu objects
    
    If (mnu.Index > 0) Then
        Dim mnuTmp As CMenu
        For Each mnuTmp In m_objParent.Menu
            If (mnuTmp.ParentKey = mnu.ParentKey) Then
                If (mnuTmp.Index >= mnu.Index) Then
                    mnuTmp.Index = mnuTmp.Index - 1
                End If
            End If
        Next
        Set mnuTmp = Nothing
    End If
    
    ' Tidy Up
    
    Set mnu = Nothing

    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Clear()
    Dim m As CMenu
    For Each m In mCol
        If m.ParentKey = "" Then
            Remove m.Key
        End If
    Next
End Sub

Public Property Get RootMenuItems() As Long
    Dim l As Long
    Dim m As CMenu
    For Each m In mCol
        If (Len(m.Key) = 0) Then
            l = l + 1
        End If
    Next
    Set m = Nothing
    RootMenuItems = l
End Property

Private Function NewMenuItem() As Menu
    Dim mnu As Variant
    Set mnu = m_objParent.ChildMenuArray

    Debug.Assert Not IsNull(mnu)

    Dim lng As Long
    lng = mnu.UBound + 1
    Load mnu(lng)
    Set NewMenuItem = mnu(lng)
    mnu = Null
End Function

