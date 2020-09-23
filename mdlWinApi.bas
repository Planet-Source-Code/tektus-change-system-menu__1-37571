Attribute VB_Name = "mdlWinApi"
Option Explicit

'System Menu Exemple
'Ozan Yasin Dogan, Istanbul / Turkey
'-----------------------------------
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Sub RemoveSysMenu(FormName As Form, ByVal MenuItemNumber As Long)
    Dim hSysMenu As Long, nCnt As Long
    hSysMenu = GetSystemMenu(FormName.hwnd, False)
    If hSysMenu Then
       nCnt = GetMenuItemCount(hSysMenu)
       If nCnt Then
          RemoveMenu hSysMenu, MenuItemNumber, MF_BYPOSITION Or MF_REMOVE
          DrawMenuBar FormName.hwnd
       End If
    End If
End Sub
