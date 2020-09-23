VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

#Const DEBUG_MODE = False

Private m_strCaption As String
Private m_strKey As String
Private m_strParentKey As String
Private m_blnEnabled As Boolean
Private m_strTag As String
Private m_hMenu As Long
Private m_hPopupMenu As Long
Private m_lngItemID As Long
Private m_lngIndex As Long
Private m_blnChecked As Boolean
Private m_lngChildCount As Long
Private m_blnPopupMenu As Boolean
Private m_objMenuItem As Menu

Private Sub Class_Initialize()
    m_strCaption = ""
    m_strKey = ""
    m_strParentKey = ""
    m_strTag = ""
    m_hMenu = 0
    m_hPopupMenu = 0
    m_lngItemID = 0
    m_lngIndex = 0
    m_blnEnabled = True
    m_blnChecked = False
    m_blnPopupMenu = False
End Sub
Private Sub Class_Terminate()
    Set m_objMenuItem = Nothing
End Sub

Public Property Get Caption() As String
    Caption = m_strCaption
End Property
Public Property Let Caption(ByVal vData As String)
    m_strCaption = vData
    
    If IsMenu(m_hMenu) Then
        Dim lRet As Long
        If m_hPopupMenu Then
            lRet = ModifyMenu(m_hMenu, m_hPopupMenu, MF_POPUP, m_hPopupMenu, _
                                                        vData)
        Else
            If vData = "-" Then
                lRet = ModifyMenu(m_hMenu, m_lngItemID, _
                        MF_SEPARATOR Or MF_BYCOMMAND, 0&, 0&)
            Else
                lRet = ModifyMenu(m_hMenu, m_lngItemID, MF_BYCOMMAND, _
                                                        m_lngItemID, vData)
            End If
        End If
#If DEBUG_MODE Then
        If (lRet = 0) Then
            MsgBox "CMenu_Caption(Let) - Failed to ModifyMenu"
        End If
#End If
    End If
End Property

Public Property Get Key() As String
    Key = m_strKey
End Property
Friend Property Let Key(ByVal vData As String)
    m_strKey = vData
End Property

Public Property Get ParentKey() As String
    ParentKey = m_strParentKey
End Property
Friend Property Let ParentKey(ByVal vData As String)
    m_strParentKey = vData
End Property

Public Property Get Tag() As String
    Tag = m_strTag
End Property
Public Property Let Tag(ByVal vData As String)
    m_strTag = vData
End Property

Public Property Get IsPopup() As Boolean
    IsPopup = m_blnPopupMenu
End Property
Friend Property Let IsPopupMenu(ByVal vNewValue As Boolean)
    m_blnPopupMenu = vNewValue
End Property

Friend Property Get hMenu() As Long
    hMenu = m_hMenu
End Property
Friend Property Let hMenu(ByVal vNewValue As Long)
    m_hMenu = vNewValue
End Property

Friend Property Get hPopupMenu() As Long
    hPopupMenu = m_hPopupMenu
End Property
Friend Property Let hPopupMenu(ByVal vNewValue As Long)
    m_hPopupMenu = vNewValue
End Property

Friend Property Get ItemID() As Long
    ItemID = m_lngItemID
End Property
Friend Property Let ItemID(ByVal vNewValue As Long)
    m_lngItemID = vNewValue
End Property

Friend Property Get Index() As Long
    Index = m_lngIndex
End Property
Friend Property Let Index(ByVal vNewValue As Long)
    m_lngIndex = vNewValue
End Property

Public Property Get Checked() As Boolean
    Checked = m_blnChecked
End Property
Public Property Let Checked(ByVal vData As Boolean)
        
    m_blnChecked = False
    If m_blnPopupMenu Then Exit Property
    If (IsMenu(m_hMenu) = 0) Then Exit Property
    If Caption = "-" Then Exit Property
    m_blnChecked = vData
    
    Dim lngRet As Long
    If vData Then
        lngRet = CheckMenuItem(m_hMenu, m_lngItemID, _
                                MF_BYCOMMAND Or MF_CHECKED)
    Else
        lngRet = CheckMenuItem(m_hMenu, m_lngItemID, _
                                MF_BYCOMMAND Or MF_UNCHECKED)
    End If
#If DEBUG_MODE Then
    If (lngRet = -1) Then
        MsgBox "CMenu_Checked(Let) - Failed to CheckMenuItem"
    End If
#End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_blnEnabled
End Property
Public Property Let Enabled(ByVal vData As Boolean)
    m_blnEnabled = vData
    If (IsMenu(m_hMenu) = 0) Then Exit Property
    If Caption = "-" Then Exit Property
    
    Dim lngRet As Long
    If m_hPopupMenu Then
        If vData Then
            lngRet = EnableMenuItem(m_hMenu, m_hPopupMenu, _
                                    MF_ENABLED)
        Else
            lngRet = EnableMenuItem(m_hMenu, m_hPopupMenu, _
                                    MF_GRAYED)
        End If
    Else
        If vData Then
            lngRet = EnableMenuItem(m_hMenu, m_lngItemID, _
                                    MF_BYCOMMAND Or MF_ENABLED)
        Else
            lngRet = EnableMenuItem(m_hMenu, m_lngItemID, _
                                    MF_BYCOMMAND Or MF_GRAYED)
        End If
    End If
    
#If DEBUG_MODE Then
    If (lngRet = -1) Then
        MsgBox "CMenu_Enabled(Let) - Failed to EnableMenuItem"
    End If
#End If
End Property

Public Property Get ChildCount() As Long
    If (hPopupMenu <> 0) Then
        ChildCount = m_lngChildCount
    Else
        ' Err.Raise ???
    End If
End Property

Friend Property Let ChildCount(ByVal vNewValue As Long)
    m_lngChildCount = vNewValue
End Property

Friend Property Get MenuItem() As Menu
    Set MenuItem = m_objMenuItem
End Property
Friend Property Set MenuItem(NewValue As Menu)
    Set m_objMenuItem = NewValue
End Property

