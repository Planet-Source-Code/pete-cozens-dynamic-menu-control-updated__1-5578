Attribute VB_Name = "modMenuAPI"
Option Explicit

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_SEPARATOR = &H800&

Public Const MF_GRAYED = &H1&
Public Const MF_ENABLED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_UNCHECKED = &H0&

Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (A As Any, B As Any, ByVal C As Long)
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal uIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal uFlag As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

'*********************************************************************************************
' Menu flags
'*********************************************************************************************
Enum MenuFlags
    Checked = &H1
    Hidden = &H2
    Grayed = &H4
    PopUp = &H8
    WindowList = &H20
    LastItem = &H100
End Enum


'*********************************************************************************************
' Internal menu struct
'*********************************************************************************************
Type MenuStruct
'    Reserved(0 To 48) As Long
'    '              ^
'    '              |
'    '              +---- For VB6 replace the 48 by 54
'    '
    Reserved(0 To 54) As Long
    '              ^
    '              |
    '              +---- For VB6 replace the 48 by 54
    '
    dwFlags As MenuFlags
    lpNextMenu As Long
    lpFirstItem As Long
    lpszName As Long
    hMenu As Long
    wID As Integer
    wShortcut As Integer
End Type

'*********************************************************************************************
' Menu shortcuts
'*********************************************************************************************
Enum MenuShortcuts
    vbNoShortcut
    vbCtrlA
    vbCtrlB
    vbCtrlC
    vbCtrlD
    vbCtrlE
    vbCtrlF
    vbCtrlG
    vbCtrlH
    vbCtrlI
    vbCtrlJ
    vbCtrlK
    vbCtrlL
    vbCtrlM
    vbCtrlN
    vbCtrlO
    vbCtrlP
    vbCtrlQ
    vbCtrlR
    vbCtrlS
    vbCtrlT
    vbCtrlU
    vbCtrlV
    vbCtrlW
    vbCtrlX
    vbCtrlY
    vbCtrlZ
    vbF1
    vbF2
    vbF3
    vbF4
    vbF5
    vbF6
    vbF7
    vbF8
    vbF9
    vbF10
    vbF11
    vbF12
    vbCtrlF1
    vbCtrlF2
    vbCtrlF3
    vbCtrlF4
    vbCtrlF5
    vbCtrlF6
    vbCtrlF7
    vbCtrlF8
    vbCtrlF9
    vbCtrlF10
    vbCtrlF11
    vbCtrlF12
    vbShiftF1
    vbShiftF2
    vbShiftF3
    vbShiftF4
    vbShiftF5
    vbShiftF6
    vbShiftF7
    vbShiftF8
    vbShiftF9
    vbShiftF10
    vbShiftF11
    vbShiftF12
    vbShiftCtrlF1
    vbShiftCtrlF2
    vbShiftCtrlF3
    vbShiftCtrlF4
    vbShiftCtrlF5
    vbShiftCtrlF6
    vbShiftCtrlF7
    vbShiftCtrlF8
    vbShiftCtrlF9
    vbShiftCtrlF10
    vbShiftCtrlF11
    vbShiftCtrlF12
    vbCtrlInsert
    vbShiftInset
    vbDelete
    vbShiftDel
    vbAltBackspace
End Enum

'*********************************************************************************************
' GetFirstChildMenu
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object of which the
'                   menu handle is wanted.
'
' Returns:
'
'   The menu handle if it's a popup menu.
'
'*********************************************************************************************
Public Function GethMenu(ByVal MenuObject As VB.Menu) As Long
Dim mnu As MenuStruct
    
    ' Get the menu struct
    
    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)
    
    ' Get the hMenu only if the menu
    ' is a popup menu. A popup menu is a
    ' menu with child items.

    If mnu.lpFirstItem Then
        GethMenu = mnu.hMenu
    Else
        GethMenu = 0
    End If
    
    If IsMenu(GethMenu) = 0 Then
        MoveMemory GethMenu, ByVal (ObjPtr(MenuObject) + 224), Len(GethMenu)
    End If
    
    Debug.Assert (IsMenu(GethMenu) <> 0)
    
End Function

'*********************************************************************************************
' GetParenthMenu
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object of which
'                   parent menu handle is wanted.
'
' Returns:
'
'   The parent menu handle.
'
'*********************************************************************************************
Public Function GetParenthMenu(ByVal MenuObject As VB.Menu) As Long
Dim mnu As MenuStruct
    
    ' Get the menu struct
    
    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)
    
    ' Get the hMenu only if the menu
    ' isn't a popup menu.

    If mnu.lpFirstItem = 0 Then
        GetParenthMenu = mnu.hMenu
    Else
        GetParenthMenu = 0
    End If
    
End Function

'*********************************************************************************************
' GetShortcut
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object of which
'                   the shortcut is wanted.
'
' Returns:
'
'   The shortcut.
'
'*********************************************************************************************
Public Function GetShortcut(ByVal MenuObject As VB.Menu) As MenuShortcuts
Dim mnu As MenuStruct

    ' Get the menu struct
    
    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)
    
    ' Only non popup items can have a shortcut
    
    If mnu.lpFirstItem = 0 Then
        GetShortcut = mnu.wShortcut
    Else
        GetShortcut = vbNoShortcut
    End If
    
End Function



'*********************************************************************************************
' GetNextMenu
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object of which
'                   next menu object is wanted.
'
' Returns:
'
'   The next menu object or Nothing if this is the
'   last menu.
'
'*********************************************************************************************
Public Function GetNextMenu(ByVal MenuObject As VB.Menu) As VB.Menu
Dim mnu As MenuStruct, Nxt As VB.Menu

    ' Get the menu struct
    
    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)
            
    ' Get the next menu only if there's one
    ' and this is not the last item
    
    If mnu.lpNextMenu <> 0 And (mnu.dwFlags And LastItem) = 0 Then
        
        ' Get a copy without AddRef
        MoveMemory Nxt, mnu.lpNextMenu, 4
        
        ' Get the object with AddRef
        Set GetNextMenu = Nxt
        
        ' Release the copy
        MoveMemory Nxt, 0&, 4
        
    End If

End Function

'*********************************************************************************************
' GetParentMenu
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object of which
'                   parent menu object is wanted.
'
' Returns:
'
'   The parent menu object.
'
'*********************************************************************************************
Public Function GetParentMenu(ByVal MenuObject As VB.Menu) As VB.Menu
Dim Nxt As VB.Menu, mnu As MenuStruct

    ' Get the next menu until we found the last item
    ' in the menu. In the last menu object the next
    ' menu points to the parent.

    Set Nxt = GetNextMenu(MenuObject)
    
    Do While Not Nxt Is Nothing
    
        MoveMemory mnu, ByVal ObjPtr(Nxt), Len(mnu)
        
        If (mnu.dwFlags And LastItem) = LastItem Then
            
            Dim Parent As VB.Menu
            
            MoveMemory Parent, mnu.lpNextMenu, 4
        
            ' Get the object with AddRef
            Set GetParentMenu = Parent
        
            MoveMemory Parent, 0&, 4
            
            Exit Do
            
        End If
        
        Set Nxt = GetNextMenu(Nxt)
        
    Loop

End Function


'*********************************************************************************************
' GetFirstChildMenu
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object of which the
'                   first child is wanted.
'
' Returns:
'
'   The first child menu object.
'
'*********************************************************************************************
Public Function GetFirstChildMenu(ByVal MenuObject As VB.Menu) As VB.Menu
Dim mnu As MenuStruct, Itm As Menu

    ' Get menu struct from object
    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)

    ' Check the pointer. If it's null there's
    ' no child item.
    If mnu.lpFirstItem <> 0 Then
        
        ' Get the object reference. Since
        ' IUnknown::AddRef is not called
        ' DO NOT set this object to Nothing.
        MoveMemory Itm, mnu.lpFirstItem, 4
        
        ' Get a copy with AddRef.
        Set GetFirstChildMenu = Itm
        
        MoveMemory Itm, 0&, 4
        
    End If
    
End Function



'*********************************************************************************************
' GetShortcut
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object.
'
' Returns:
'
'   True if the menu item has children, otherwise False.
'
'*********************************************************************************************
Public Function IsPopupMenu(MenuObject As VB.Menu) As Boolean
Dim mnu As MenuStruct

    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)
    
    IsPopupMenu = mnu.lpFirstItem

End Function

'*********************************************************************************************
' SetShortcut
'*********************************************************************************************
'
' Changes a menu shortcut
'
' Parameters:
'
'   MenuObject:     The menu object.
'
'*********************************************************************************************
Public Sub SetShortcut(ByVal MenuObject As VB.Menu, ByVal NewShortcut As MenuShortcuts)
Dim mnu As MenuStruct

    ' Get the menu struct
    
    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)
    
    ' Only non popup items can have a shortcut
    
    If mnu.lpFirstItem = 0 Then
    
        If NewShortcut > vbAltBackspace Then
            NewShortcut = vbAltBackspace
        ElseIf NewShortcut < 0 Then
            NewShortcut = vbNoShortcut
        End If
        
        mnu.wShortcut = NewShortcut
        
        ' Change only that value
    
        MoveMemory ByVal ObjPtr(MenuObject) + 218, mnu.wShortcut, 2
        
        ' Set the caption to update
        ' the shortcut text
        MenuObject.Caption = MenuObject.Caption
        
    End If

End Sub

'*********************************************************************************************
' GetCommand
'*********************************************************************************************
'
' Parameters:
'
'   MenuObject:     The menu object of which the
'                   command is wanted. Only non
'                   popup items have a command.
'
' Returns:
'
'   The command.
'
'*********************************************************************************************
Public Function GetCommand(ByVal MenuObject As VB.Menu) As Long
Dim mnu As MenuStruct
    
    ' Get the menu struct from object
    
    MoveMemory mnu, ByVal ObjPtr(MenuObject), Len(mnu)
    
    ' Get the command only if the menu
    ' isn't a popup menu.

    If mnu.lpFirstItem = 0 Then
        GetCommand = mnu.wID
    Else
        GetCommand = 0
    End If
    
    If GetCommand = 0 Then
        Dim l As Integer
        MoveMemory l, ByVal (ObjPtr(MenuObject) + 228), Len(l)
        GetCommand = l
    End If
    
End Function


