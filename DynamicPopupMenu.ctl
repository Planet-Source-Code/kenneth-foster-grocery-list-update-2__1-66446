VERSION 5.00
Begin VB.UserControl DynamicPopupMenu 
   BackColor       =   &H00000000&
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   InvisibleAtRuntime=   -1  'True
   Picture         =   "DynamicPopupMenu.ctx":0000
   ScaleHeight     =   435
   ScaleWidth      =   1050
   ToolboxBitmap   =   "DynamicPopupMenu.ctx":068A
End
Attribute VB_Name = "DynamicPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetLastError Lib "kernel32.dll" () As Long

' Exposed Enumeration
Public Enum mceItemStates
    mceDisabled = 1
    mceGrayed = 2
End Enum

' Property variables
Private psCaption As String             ' Caption of menu item (with the arrow >) if this is submenu
Private piHwnd As Long                  ' Handle to Menu
Private m_ItemCaption(999) As String

' Supporting API code
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_TYPE = &H10
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_RIGHTBUTTON = &H2

Private Type POINT
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
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
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, lpNewItem As String) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal uFlags As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal Hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal Hwnd As Long, ByVal lptpm As Any) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Property Let Caption(ByVal sCaption As String)

    psCaption = sCaption
    
End Property


Private Property Get Caption() As String

    Caption = psCaption
    
End Property



Private Sub Remove(ByVal iMenuPosition As Long)
    
    DeleteMenu piHwnd, iMenuPosition, MF_BYPOSITION
    
End Sub

Public Property Get Hwnd() As Long
    
    Hwnd = piHwnd

End Property


Private Sub Add(ByVal iMenuID As Long, vMenuItem As Variant, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional eItemState As mceItemStates, Optional ByVal imgUnchecked As Long = 0, Optional ByVal imgChecked As Long = 0)
    
    If vMenuItem = "-" Then ' Make a seperator
        AppendMenu piHwnd, MF_STRING Or MF_SEPARATOR, iMenuID, ByVal vbNullString
    Else
        AppendMenu piHwnd, MF_STRING Or -bChecked * MF_CHECKED, iMenuID, ByVal vMenuItem
    End If

    ' Menu Icons
    If imgChecked = 0 Then imgChecked = imgChecked ' Need a value for both
    SetMenuItemBitmaps piHwnd, iMenuID, MF_BYCOMMAND, imgUnchecked, imgChecked
    
    ' Default item
    If bDefault Then SetMenuDefaultItem piHwnd, iMenuID, 0
    
    ' Disabled (Regular color text)
    If eItemState = mceDisabled Then EnableMenuItem piHwnd, iMenuID, MF_BYCOMMAND Or MF_DISABLED
    ' Disabled (disabled color text)
    If eItemState = mceGrayed Then EnableMenuItem piHwnd, iMenuID, MF_BYCOMMAND Or MF_GRAYED

    ' ADDED
    m_ItemCaption(iMenuID) = CStr(vMenuItem)

End Sub


Private Function Show(Optional ByVal iFormHwnd As Long = -1, _
                     Optional ByVal X As Long = -1, _
                     Optional ByVal Y As Long = -1, _
                     Optional ByVal iControlHwnd As Long = -1) As Long
                     
    Dim iHwnd As Long, iX As Long, iY As Long
    
    ' If no form is passed, use the current window
    iFormHwnd = Screen.ActiveForm.Hwnd
    iHwnd = iFormHwnd
    
    ' If passed a control handle, left-bottom orient to the control.
    If iControlHwnd <> -1 Then
        Dim rt As RECT
        GetWindowRect iControlHwnd, rt
        iX = rt.Left
        iY = rt.Bottom
    Else
        Dim pt As POINT
        GetCursorPos pt
        If X = -1 Then iX = pt.X Else: iX = X
        If Y = -1 Then iY = pt.Y Else: iY = Y
    End If
    '
    Show = TrackPopupMenuEx(piHwnd, TPM_RETURNCMD Or TPM_RIGHTBUTTON, iX, iY, iHwnd, ByVal 0&)
    
    
End Function

Private Function GetItemCaption(ByVal itm As Long) As String

    ' ADDED SO WE CAN READ THE VALUE OF THE SELECTED ITEM
    '
    On Error GoTo Err
    GetItemCaption = m_ItemCaption(itm)
    Exit Function
    
Err:
    GetItemCaption = "UNKNOWN"
    
End Function


Public Function Popup(ByVal mItems As String) As String

    If Len(Trim$(mItems)) = 0 Then Exit Function
    '
    Dim ItemToAdd() As String
    ItemToAdd = Split(mItems, ",")
    '
    ' ALTHOUGH THE CODE CAN DO MORE
    ' THIS IS ONLY ADDING LIST ITEMS, NO BITMAPS, NO
    ' CHECKMARKS OR DEFAULS OR SUBMENUS
    '
    Dim i As Long
    For i = 0 To UBound(ItemToAdd)
        If Len(Trim$(ItemToAdd(i))) > 0 Then
            Add i + 1, Trim$(ItemToAdd(i))
        End If
    Next
    '
    ' THIS SHOWS THE MENU, AND RETURNS THE SELECTED CHOICE AS A STRING
    Popup = GetItemCaption(Show())
    '
    ' GET RID OF MENU AND ITEMS
    DestroyMenu piHwnd
    '
    ' RECREATE FOR NEXT TIME
    '
    piHwnd = CreatePopupMenu()
    
End Function

Private Sub UserControl_Initialize()

    piHwnd = CreatePopupMenu()

End Sub

Private Sub UserControl_Terminate()
    
    DestroyMenu piHwnd

End Sub
