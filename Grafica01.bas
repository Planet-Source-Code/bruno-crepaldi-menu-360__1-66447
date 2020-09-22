Attribute VB_Name = "Grafica01"
Option Explicit
'
'                 Dichiarazioni Mouse
'   Ritorna la posizione Assoluta del mouse in PIXEL  X e Y
'
Private Declare Function M_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINT) As Long
Private Type POINT
      X       As Long
      Y       As Long
End Type
Public pt      As POINT
'
'            Dichiarazioni per Form Trasparente
'
Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const BM_SETSTATE = &HF3

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'
Private Secondi As Long
Private TimeOld As Variant
'
Public MouseX As Long
Public MouseY As Long
Public OldMouseX As Long
Public OldMouseY As Long
Public MouseDown As Boolean
' Screen Saver
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
               (ByVal uAction As Long, ByVal uParam As Long, _
               ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long

Public Sub ScreenSaverActive(Active As Boolean)
Dim Enabled As Long
Dim ret As Long

Enabled = IIf(Active, 1, 0)
ret = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, Enabled, 0, 0)
End Sub
'
'                         Sezione Mouse
'
Public Sub GetCursorPos(xX As Long, xY As Long)
  Call M_GetCursorPos(pt)
  xX = pt.X
  xY = pt.Y
End Sub

Public Sub SpostaForm(MyFrm As Form)
   Dim DifX As Long
   Dim DifY As Long
   
     If MouseDown = False Then Exit Sub ' Tasto non Premuto esci
     
     OldMouseX = MyFrm.ScaleX(MouseX, vbPixels, vbTwips)
     OldMouseY = MyFrm.ScaleY(MouseY, vbPixels, vbTwips)
   
     Call GetCursorPos(MouseX, MouseY)
       
     DifX = OldMouseX - MyFrm.ScaleX(MouseX, vbPixels, vbTwips)
     DifY = OldMouseY - MyFrm.ScaleY(MouseY, vbPixels, vbTwips)

     MyFrm.Left = MyFrm.Left - DifX
     MyFrm.Top = MyFrm.Top - DifY

End Sub
'
'
'                       Form Trasparente
'
Public Sub FormTrasparente(MyFrm As Form, Clr As Double)
  Dim ret As Long
  
  ret = GetWindowLong(MyFrm.hWnd, GWL_EXSTYLE)
  ret = ret Or WS_EX_LAYERED
  SetWindowLong MyFrm.hWnd, GWL_EXSTYLE, ret
  SetLayeredWindowAttributes MyFrm.hWnd, Clr, 0, LWA_COLORKEY
End Sub
