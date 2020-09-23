VERSION 5.00
Begin VB.Form frmAimPw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aim: De/Encrypt Saved PW"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   1890
      Width           =   2955
   End
   Begin VB.CommandButton cmdDelPW 
      Caption         =   "Delete PW"
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1530
      Width           =   1455
   End
   Begin VB.CommandButton cmdChangePW 
      Caption         =   "Change PW"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   1530
      Width           =   1455
   End
   Begin VB.ListBox lstPW 
      Height          =   1425
      Left            =   1560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   1455
   End
   Begin VB.ListBox lstSN 
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1455
   End
End
Attribute VB_Name = "frmAimPw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const AimUsers = "Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Users"
Private Const HKEY_CURRENT_USER = -2147483647
Private Const REG_SZ = 1
Private Declare Function RegCloseKey Lib "advapi32" (ByVal Hkey As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal Hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Function DeCrypt(PW As String) As String
    Dim A As Long, B As Byte, C As Byte, D As Byte
    DeCrypt = Space$(Len(PW) / 2)
    For A = 1 To Len(PW) Step 2
        C = A \ 2 Mod 16 + 1
        B = Asc(Mid$(PW, A, 1)) - 65
        If C = 4 Or C = 9 Or C = 11 Or C = 14 Then B = B + IIf(B Mod 2 < 1, 1, -1)
        If C = 5 Or C = 10 Or C = 15 Or C = 16 Then B = B + IIf(B Mod 4 < 2, 2, -2)
        If C = 1 Or C = 6 Or C = 11 Or C = 12 Then B = B + IIf(B Mod 8 < 4, 4, -4)
        If C = 2 Or C = 7 Or C = 13 Or C = 16 Then B = B + IIf(B < 8, 8, -8)
        If C = 12 Or C = 13 Or C = 14 Or C = 15 Then B = 15 - B
        D = B * 16
        B = Asc(Mid$(PW, A + 1, 1)) - 65
        If C = 5 Or C = 7 Or C = 9 Or C = 10 Or C = 16 Then B = B + IIf(B Mod 2 < 1, 1, -1)
        If C = 1 Or C = 6 Or C = 9 Or C = 11 Or C = 12 Then B = B + IIf(B Mod 4 < 2, 2, -2)
        If C = 2 Or C = 7 Or C = 8 Or C = 9 Or C = 13 Then B = B + IIf(B Mod 8 < 4, 4, -4)
        If C = 3 Or C = 12 Or C = 14 Then B = B + IIf(B < 8, 8, -8)
        If C = 8 Or C = 10 Or C = 11 Then B = 15 - B
        Mid$(DeCrypt, A \ 2 + 1, 1) = Chr$(D + B)
    Next
End Function
Private Function EnCrypt(PW As String) As String
    Dim A As Long, B As Byte, C As Byte, D As Byte
    EnCrypt = Space$(2 * Len(PW))
    For A = 1 To Len(PW)
        B = Asc(Mid$(PW, A, 1)) \ 16
        C = A Mod 16
        If C = 4 Or C = 9 Or C = 11 Or C = 14 Then B = B + IIf(B Mod 2 < 1, 1, -1)
        If C = 5 Or C = 10 Or C = 15 Or C = 0 Then B = B + IIf(B Mod 4 < 2, 2, -2)
        If C = 1 Or C = 6 Or C = 11 Or C = 12 Then B = B + IIf(B Mod 8 < 4, 4, -4)
        If C = 2 Or C = 7 Or C = 13 Or C = 0 Then B = B + IIf(B < 8, 8, -8)
        If C = 12 Or C = 13 Or C = 14 Or C = 15 Then B = 15 - B
        D = B
        B = Asc(Mid$(PW, A, 1)) Mod 16
        If C = 5 Or C = 7 Or C = 9 Or C = 10 Or C = 0 Then B = B + IIf(B Mod 2 < 1, 1, -1)
        If C = 1 Or C = 6 Or C = 9 Or C = 11 Or C = 12 Then B = B + IIf(B Mod 4 < 2, 2, -2)
        If C = 2 Or C = 7 Or C = 8 Or C = 9 Or C = 13 Then B = B + IIf(B Mod 8 < 4, 4, -4)
        If C = 3 Or C = 12 Or C = 14 Then B = B + IIf(B < 8, 8, -8)
        If C = 8 Or C = 10 Or C = 11 Then B = 15 - B
        Mid$(EnCrypt, 2 * A - 1, 2) = Chr$(D + 65) & Chr$(B + 65)
    Next
End Function
Private Sub cmdChangePW_Click()
    Dim Hkey As Long, NewPW As String
    If RegOpenKeyEx(HKEY_CURRENT_USER, AimUsers & "\" & lstSN.List(lstSN.ListIndex) & "\Login", 0, 0, Hkey) Then
        RegCloseKey Hkey
        MsgBox "Can't Find AIM", 0, "Error"
        Exit Sub
    End If
    NewPW = "ÿÿ" & EnCrypt(InputBox("Enter New PassWord:", "New PW"))
    If Len(NewPW) > 2 Then RegSetValueEx Hkey, "Password", 0, REG_SZ, ByVal NewPW, Len(NewPW)
    RegCloseKey Hkey
    cmdRefresh_Click
End Sub
Private Sub cmdDelPW_Click()
    Dim Hkey As Long, NewPW As String
    If RegOpenKeyEx(HKEY_CURRENT_USER, AimUsers & "\" & lstSN.List(lstSN.ListIndex) & "\Login", 0, 0, Hkey) Then
        RegCloseKey Hkey
        MsgBox "Can't Find AIM", 0, "Error"
        Exit Sub
    End If
    RegSetValueEx Hkey, "Password", 0, REG_SZ, ByVal "", 0
    RegCloseKey Hkey
    cmdRefresh_Click
End Sub
Private Sub cmdRefresh_Click()
    lstSN.Clear
    lstPW.Clear
    Form_Load
End Sub
Private Sub Form_Load()
    Dim Key As Long, Hkey As Long, hKey2 As Long, SN As String, PW As String
    If RegOpenKeyEx(HKEY_CURRENT_USER, AimUsers, 0, 0, Hkey) Then
        RegCloseKey Hkey
        MsgBox "Can't Find AIM", 0, "Error"
        Exit Sub
    End If
    Key = 0
    Do
        SN = String(256, 0)
        PW = String(256, 0)
        If RegEnumKey(Hkey, Key, SN, 256) Then
            RegCloseKey Hkey
            Exit Do
        Else
            SN = Left(SN, InStr(SN, Chr(0)) - 1)
            lstSN.AddItem SN
            RegOpenKeyEx HKEY_CURRENT_USER, AimUsers & "\" & SN & "\Login", 0, 0, hKey2
            If RegQueryValueExStr(hKey2, "Password", 0, REG_SZ, PW, 256) Then
                lstPW.AddItem "Not Found"
            Else
                PW = Left(PW, InStr(PW, Chr(0)) - 1)
                If PW = "" Or Left(PW, 2) <> "ÿÿ" Then
                lstPW.AddItem "Not Found"
                Else
                lstPW.AddItem DeCrypt(Right(PW, Len(PW) - 2))
                End If
                RegCloseKey hKey2
            End If
        End If
        Key = Key + 1
    Loop
    If Key = 0 Then Exit Sub
    lstSN.ListIndex = 0
    lstPW.ListIndex = 0
End Sub
Private Sub lstSN_KeyPress(KeyAscii As Integer)
    lstSN_Scroll
End Sub
Private Sub lstSN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstSN_Scroll
End Sub
Private Sub lstSN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstSN_Scroll
End Sub
Private Sub lstSN_Scroll()
    lstPW.TopIndex = lstSN.TopIndex
    lstPW.ListIndex = lstSN.ListIndex
End Sub
Private Sub lstPW_KeyPress(KeyAscii As Integer)
    lstPW_Scroll
End Sub
Private Sub lstPW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstPW_Scroll
End Sub
Private Sub lstPW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstPW_Scroll
End Sub
Private Sub lstPW_Scroll()
    lstSN.TopIndex = lstPW.TopIndex
    lstSN.ListIndex = lstPW.ListIndex
End Sub
