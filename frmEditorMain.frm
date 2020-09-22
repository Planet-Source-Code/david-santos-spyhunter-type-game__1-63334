VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditorMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level Editor"
   ClientHeight    =   8700
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drawing Tools"
      Height          =   2535
      Left            =   5520
      TabIndex        =   4
      Top             =   6120
      Width           =   2055
      Begin VB.CheckBox Check1 
         Caption         =   "Causes Jump"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Slows Down"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Passable"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw"
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Road Object #0 "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2460
      Left            =   0
      Picture         =   "frmEditorMain.frx":0000
      ScaleHeight     =   2400
      ScaleWidth      =   5280
      TabIndex        =   2
      Top             =   6120
      Width           =   5340
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      Left            =   7320
      Max             =   0
      Min             =   288
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picTarget 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   7275
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.PictureBox picObjects 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2460
      Left            =   0
      Picture         =   "frmEditorMain.frx":29444
      ScaleHeight     =   2400
      ScaleWidth      =   5280
      TabIndex        =   8
      Top             =   6120
      Width           =   5340
   End
   Begin VB.PictureBox picObjMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2460
      Left            =   0
      Picture         =   "frmEditorMain.frx":52888
      ScaleHeight     =   2400
      ScaleWidth      =   5280
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuToolsPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuToolsCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuToolsDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "frmEditorMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim startX As Long
Dim startY As Long
Dim areaX As Long
Dim areaY As Long

Dim selX As Long
Dim selY As Long
Dim selW As Long
Dim selH As Long

Dim lastW As Long
Dim lastH As Long
Dim carx As Long

Dim mChanged As Boolean

Dim RoadMap(0 To 15, 0 To 299) As Byte
Dim ObjectMap(0 To 15, 0 To 299) As Byte
Dim RoadInfo(0 To 54) As Byte

Dim selblock As Byte

Dim TempCopy(192) As Byte

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest

Dim mode As Integer


Private Sub cboType_Click()
 If cboType.ListIndex = 0 Then
  picSource.Visible = True
  picObjects.Visible = False
 End If
 
 If cboType.ListIndex = 1 Then
  picSource.Visible = False
  picObjects.Visible = True
 End If
 
 Render

End Sub


Private Sub Check1_Click(Index As Integer)
 If Check1(Index).Value = 0 Then
  RoadInfo(selblock) = RoadInfo(selblock) And Not 2 ^ Index
 Else
  RoadInfo(selblock) = RoadInfo(selblock) Or 2 ^ Index
 End If
End Sub

Private Sub cmdDraw_Click()
 cmdDraw.BackColor = RGB(255, 255, 0)
 cmdSelect.BackColor = RGB(192, 192, 192)
 
 
 BitBlt picTarget.hDC, 0, 0, picTarget.ScaleWidth, picTarget.ScaleHeight, picTemp.hDC, 0, 0, SRCCOPY
 mode = 1
End Sub

Private Sub cmdSelect_Click()
 cmdDraw.BackColor = RGB(192, 192, 192)
 cmdSelect.BackColor = RGB(255, 255, 0)
 
 mode = 0
End Sub

Private Sub Form_Load()
 picTarget.Width = 512 * Screen.TwipsPerPixelX
 picTarget.Height = 384 * Screen.TwipsPerPixelY
 
 picTemp.Width = 512 * Screen.TwipsPerPixelX
 picTemp.Height = 384 * Screen.TwipsPerPixelY
 picTemp.ScaleMode = vbPixels
 
 picTarget.ScaleMode = vbPixels
 picSource.ScaleMode = vbPixels
 picObjects.ScaleMode = vbPixels

 picSource.Picture = LoadPicture(App.Path & "\road.bmp")
 picObjects.Picture = LoadPicture(App.Path & "\objects.bmp")
 picObjMask.Picture = LoadPicture(App.Path & "\objects mask.bmp")
 
 If Dir("roadinfo.dat") <> "" Then
  Open "roadinfo.dat" For Binary As 1
   Get #1, , RoadInfo
  Close 1
 End If

 cboType.AddItem "Road"
 cboType.AddItem "Objects"
 cboType.ListIndex = 0
 carlong = -1
 
End Sub

Private Sub Form_Resize()
 VScroll1.Left = picTarget.Left + picTarget.Width + 10
 VScroll1.Height = picTarget.Height
 picSource.Top = picTarget.Top + picTarget.Height + 100
 picObjects.Top = picTarget.Top + picTarget.Height + 100
 Frame1.Top = picTarget.Top + picTarget.Height + 100

End Sub


Private Sub Form_Unload(Cancel As Integer)
 If mChanged Then
  ret = MsgBox("Would you like to save changes to " & CommonDialog1.FileTitle & "?", vbYesNoCancel + vbQuestion, App.Title)
  If ret = vbYes Then mnuFileSave_Click
  If ret = vbCancel Then Cancel = 1
 End If
 Open "roadinfo.dat" For Binary As 1
  Put #1, , RoadInfo
 Close 1
End Sub

Private Sub mnuFileExit_Click()
 Unload Me
End Sub

Private Sub mnuFileLoad_Click()
On Error GoTo errhandler
 CommonDialog1.CancelError = True
 CommonDialog1.Filter = "Map Files|*.map|"
 CommonDialog1.ShowOpen
 
 Open CommonDialog1.FileName For Binary As 1
  Get #1, , RoadMap
  Get #1, , ObjectMap
  Get #1, , carx
 Close 1
 
 VScroll1.Value = 1
 VScroll1.Value = 0
 cmdSelect_Click
 
 Exit Sub
errhandler:
 If Err.Number <> cdlCancel Then
  Err.Raise Err.Number
  Stop
 End If
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo errhandler
 CommonDialog1.CancelError = True
 CommonDialog1.Filter = "Map Files|*.map|"
 CommonDialog1.ShowSave
 Dim fso As New FileSystemObject
 
 If carx = -1 Then
  MsgBox "You need to put a red car object in the very first row of the map!", vbExclamation, App.Title
  Exit Sub
 End If
 
 If fso.FileExists(CommonDialog1.FileName) Then
  ret = MsgBox("Are you aure you want to overwrite " & CommonDialog1.FileTitle & "?", vbYesNo + vbDefaultButton2 + vbQuestion, App.Title)
  If ret = vbNo Then Exit Sub
 End If
 
 Open CommonDialog1.FileName For Binary As 1
  Put #1, , RoadMap
  Put #1, , ObjectMap
  Put #1, , carx
 Close 1
 
 mChanged = False
 
 Exit Sub
errhandler:
 If Err.Number <> cdlCancel Then
  Err.Raise Err.Number
  Stop
 End If
End Sub

Private Sub mnuToolsCopy_Click()
 For i = 0 To (selH \ 32) - 1
  For j = 0 To (selW \ 32) - 1
   If cboType.ListIndex = 0 Then TempCopy(j + i * (selW \ 32)) = RoadMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value)
   If cboType.ListIndex = 1 Then TempCopy(j + i * (selW \ 32)) = ObjectMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value)
  Next
 Next
 lastW = selW
 lastH = selH
 'BitBlt picTarget.hDC, 0, 0, picTarget.ScaleWidth, picTarget.ScaleHeight, picTemp.hDC, 0, 0, SRCCOPY
 'picTarget.Refresh
End Sub

Private Sub mnuToolsCut_Click()
 For i = 0 To (selH \ 32) - 1
  For j = 0 To (selW \ 32) - 1
   
   If cboType.ListIndex = 0 Then
    TempCopy(j + i * (selW \ 32)) = RoadMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value)
    RoadMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value) = 0
   End If
   
   If cboType.ListIndex = 1 Then
    TempCopy(j + i * (selW \ 32)) = ObjectMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value)
    ObjectMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value) = 0
   End If
    
  Next
 Next
 mChanged = True
 
 lastW = selW
 lastH = selH
 Render
End Sub

Private Sub mnuToolsDelete_Click()
 For i = 0 To (selH \ 32) - 1
  For j = 0 To (selW \ 32) - 1
   If ((selX \ 32) + j < 16) And (11 - (selY \ 32) - i + VScroll1.Value >= 0) Then
    If cboType.ListIndex = 0 Then RoadMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value) = 0
    If cboType.ListIndex = 1 Then ObjectMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value) = 0
   End If
  Next
 Next
 mChanged = True
 
 Render
End Sub

Private Sub mnuToolsPaste_Click()
 For i = 0 To (lastH \ 32) - 1
  For j = 0 To (lastW \ 32) - 1
   If ((selX \ 32) + j < 16) And (11 - (selY \ 32) - i + VScroll1.Value >= 0) Then
    If cboType.ListIndex = 0 Then RoadMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value) = TempCopy(j + i * (lastW \ 32))
    If cboType.ListIndex = 1 Then ObjectMap((selX \ 32) + j, 11 - (selY \ 32) - i + VScroll1.Value) = TempCopy(j + i * (lastW \ 32))
   End If
  Next
 Next
 mChanged = True


 BitBlt picTarget.hDC, 0, 0, picTarget.ScaleWidth, picTarget.ScaleHeight, picTemp.hDC, 0, 0, SRCCOPY
 picTarget.Refresh
 Render
End Sub

Private Sub picSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  startX = x
  startY = y
  areaX = x + 32
  areaY = y + 32
  picSource.Line (startX, startY)-(x + 32, y + 32), RGB(255, 255, 0), B
  selblock = x \ 32 + (y \ 32) * 11
  For i = 0 To Check1.UBound
   Check1(i).Value = 1 And (RoadInfo(selblock) \ (2 ^ i))
  Next
  
 End If
End Sub

Private Sub picSource_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  picSource.Cls
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  picSource.Line (startX, startY)-(x + 32, y + 32), RGB(255, 255, 0), B
 End If
End Sub

Private Sub picSource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  areaX = x - startX + 32
  areaY = y - startY + 32
  picSource.Cls
  cmdDraw_Click
  Label1.Caption = "Road Object #" & (x \ 32 + (y \ 32) * 11)
 End If
End Sub

Private Sub picTarget_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
 End Select
End Sub

Private Sub picTarget_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  If mode = 0 Then
   x = Int(x / 32) * 32
   y = Int(y / 32) * 32
   selX = x
   selY = y
  End If
  picTarget_MouseMove Button, Shift, x, y
 Else
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  selX = x
  selY = y
  PopupMenu mnuTools
 End If
End Sub

Private Sub picTarget_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  If mode = 1 Then
   If (y \ 32 <= 11) And (x \ 32 >= 0) Then
    
    mChanged = True
    
    For i = 0 To (areaY \ 32) - 1
     For j = 0 To (areaX \ 32) - 1
      If (j + (x \ 32) <= 15) And ((11 - (i + y \ 32) + VScroll1.Value) >= 0) Then
       If cboType.ListIndex = 0 Then RoadMap(j + (x \ 32), 11 - (i + y \ 32) + VScroll1.Value) = startX \ 32 + (startY \ 32) * 11 + j + (i * 11)
       If cboType.ListIndex = 1 Then
        ObjectMap(j + (x \ 32), 11 - (i + y \ 32) + VScroll1.Value) = startX \ 32 + (startY \ 32) * 11 + j + (i * 11)
        If ObjectMap(j + (x \ 32), 11 - (i + y \ 32) + VScroll1.Value) = 1 Then
         If 11 - (i + y \ 32) + VScroll1.Value > 0 Then
          ObjectMap(j + (x \ 32), 11 - (i + y \ 32) + VScroll1.Value) = 0
         Else
          For B = 0 To 15
           If ObjectMap(B, 11 - (i + y \ 32) + VScroll1.Value) = 1 Then ObjectMap(B, 11 - (i + y \ 32) + VScroll1.Value) = 0
          Next
          ObjectMap(j + (x \ 32), 11 - (i + y \ 32) + VScroll1.Value) = 1
          carx = j + (x \ 32)
         End If
        End If
       End If
      End If
     Next
    Next
   
   Render
   
   End If
  Else
   selW = x + 32 - selX
   selH = y + 32 - selY
   Me.Caption = RoadMap(x \ 32, y \ 32)
  
   BitBlt picTarget.hDC, 0, 0, picTarget.ScaleWidth, picTarget.ScaleHeight, picTemp.hDC, 0, 0, SRCCOPY
   picTarget.Line (selX, selY)-(selX + selW, selY + selH), RGB(255, 255, 0), B
   picTarget.Refresh
  End If
 End If
End Sub

Private Sub VScroll1_Change()
 Render
End Sub

Private Sub VScroll1_Scroll()
 Render
End Sub

Sub Render()
Dim tile As Long
 For i = 11 To 0 Step -1
  For j = 0 To 15
   tile = RoadMap(j, VScroll1.Value + i)
   BitBlt picTemp.hDC, j * 32, (11 - i) * 32, 32, 32, picSource.hDC, (tile Mod 11) * 32, (tile \ 11) * 32, SRCCOPY
   
    tile = ObjectMap(j, VScroll1.Value + i)
    BitBlt picTemp.hDC, j * 32, (11 - i) * 32, 32, 32, picObjMask.hDC, (tile Mod 11) * 32, (tile \ 11) * 32, SRCAND
    BitBlt picTemp.hDC, j * 32, (11 - i) * 32, 32, 32, picObjects.hDC, (tile Mod 11) * 32, (tile \ 11) * 32, SRCPAINT
   
  Next
 Next
 BitBlt picTarget.hDC, 0, 0, picTarget.ScaleWidth, picTarget.ScaleHeight, picTemp.hDC, 0, 0, SRCCOPY
 picTarget.Refresh
End Sub

Private Sub picObjects_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  startX = x
  startY = y
  areaX = x + 32
  areaY = y + 32
  picObjects.Line (startX, startY)-(x + 32, y + 32), RGB(255, 255, 0), B
 End If
End Sub

Private Sub picObjects_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  picObjects.Cls
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  picObjects.Line (startX, startY)-(x + 32, y + 32), RGB(255, 255, 0), B
 End If
End Sub

Private Sub picObjects_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 1 Then
  x = Int(x / 32) * 32
  y = Int(y / 32) * 32
  areaX = x - startX + 32
  areaY = y - startY + 32
  picObjects.Cls
  cmdDraw_Click
 End If
End Sub

