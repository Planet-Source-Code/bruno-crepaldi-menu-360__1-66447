VERSION 5.00
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":058A
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ImgMenuBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   20
      Top             =   0
      Width           =   375
   End
   Begin VB.Timer Timer_Shift 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10320
      Top             =   8640
   End
   Begin VB.PictureBox ImgSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   2880
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox ImgSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   2400
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   18
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox ImgSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1920
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   17
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox ImgSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1440
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   16
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox ImgSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   960
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   15
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   12
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   11
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   10
      Top             =   3960
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   9
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   8
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   5
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox ImgTitolo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   2280
      Picture         =   "FrmMain.frx":2405CE
      ScaleHeight     =   1800
      ScaleWidth      =   9750
      TabIndex        =   1
      Top             =   6480
      Width           =   9750
   End
   Begin VB.Timer Timer_Menu 
      Interval        =   20
      Left            =   11520
      Top             =   8640
   End
   Begin VB.PictureBox PicFondo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.Timer Timer_Lmpg 
      Interval        =   400
      Left            =   10920
      Top             =   8640
   End
   Begin VB.Label LblSub 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROGDVB IRDETO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   8400
      Width           =   3615
   End
   Begin VB.Label LblSub 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROGDVB IRDETO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TrspColMsk As Long
Private i As Long
Private i1 As Long

Private WaitValue As Boolean
Private ImgNum As Long
Private Pc_X As Long                  ' posizione centrale asse X
Private Pc_Y As Long                  ' posizione centrale asse Y
Private Raggio As Long

Private Dist As Long
Private ImgStep As Long
Private M_Dir01 As Long
Private M_Dir02 As Long
Private M_Dir03 As Long
Private Lmpg As Single
Private SelSub As Long
'------------------------------------------------------------------------
'  Form Load
'------------------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------------
'  UnREM for have a Traspararent Background
'
'  ClrTrasp = &HFAE4E4
'  Me.BackColor = ClrTrasp
'  FormTrasparente Me, ClrTrasp
'------------------------------------------------------------------
  Timer_Lmpg.Interval = 450
  Timer_Lmpg.Enabled = False
  Timer_Menu.Interval = 10
  Timer_Menu.Enabled = False
  Timer_Shift.Interval = 50
  Timer_Shift.Enabled = False
  Remote_Set

  With Me
    .KeyPreview = True
   ' vbTwips ' vbPixels
    .ScaleMode = vbPixels
    .AutoRedraw = True
  '
    .Top = 0
    .Left = 0
    .Height = Screen.Height
    .Width = Screen.Width
  End With
'
   For i = 0 To 6
     CaricaFondi (i)
   Next i
'
   Label3.Caption = ""
   Label3.Top = Me.ScaleHeight - Label3.Height
   Lmpg = 0.5
 '
   LblSub(0).Visible = False
   LblSub(1).Visible = False
   LblSub(1).Left = LblSub(1).Left + 2
   LblSub(1).Top = LblSub(1).Top + 2
   CaricaTestiSub
 '-------------------------------------------------------------------
 ' Load Menu and Titles Pictures
 '-------------------------------------------------------------------
  For i = 0 To Img_Tot
    CaricaIcone (i)
    CaricaTitoli (i)
  Next i
 '
   ImgTitolo.Visible = False
   ImgTitolo.TabStop = False
   ImgTitolo.Left = (Me.ScaleWidth - ImgTitolo.Width) / 2
   ImgTitolo.Top = Me.ScaleHeight - ImgTitolo.Height - 10
 '-------------------------------------------------------------------
 ' Set Menu Pictures
 '-------------------------------------------------------------------
  ImgNum = 0
  ImgStep = (360 / (ImgQnt + 1))
  Dist = 0
'
  For ImgNum = 0 To ImgQnt
    With ImgMenu(ImgNum)
       .ScaleMode = vbPixels
       .AutoRedraw = True
       .TabStop = False
       .Visible = False
       .Picture = Img(ImgNum, 0)
    End With
    Rotazione(ImgNum) = Dist
    Dist = Dist + ImgStep
    '
    Menu.Submenu(ImgNum) = False
   Next ImgNum
 '-------------------------------------------------------------------
 ' Load Sub Menu Pictures
 '-------------------------------------------------------------------
   For i = 0 To 9  'ImgQnt
     For i1 = 0 To 4 'Img_Tot_Sub
       Call CaricaIconeSub(i, i1)
     Next i1
   Next i
 '-------------------------------------------------------------------
 ' Set ImgMenuBack
 '-------------------------------------------------------------------
  With ImgMenuBack
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .TabStop = False
      .Visible = False
      .Width = ImgMenu(0).Width
      .Height = ImgMenu(0).Height
  End With
 '-------------------------------------------------------------------
 ' Set Sub Menu Pictures
 '-------------------------------------------------------------------
  For i = 0 To ImgSub.Count - 1
    With ImgSub(i)
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .TabStop = False
      .Visible = False
      .Width = 127                ' With of Picture
    End With
 '  ImgSub(i).Left = Me.ScaleWidth / ImgSub.Count * i + (Me.ScaleWidth - ImgSub(i).Width * (ImgSub.Count + 1)) / 2
    ImgSub(i).Left = ((Me.ScaleWidth / (ImgSub.Count + 1)) * (i + 1)) - (ImgSub(i).ScaleWidth / 2)
  Next i
 '
    Menu.Submenu(0) = True
    Menu.Submenu(1) = True
  '
   Imagine.Center_X = ImgMenu(0).ScaleWidth / 2
   Imagine.Center_Y = ImgMenu(0).ScaleHeight / 2
  '
   Pc_X = (Me.ScaleWidth / 2) - Imagine.Center_X
'  Pc_Y = (Me.ScaleHeight / 2) - Imagine.Center_Y
   Pc_Y = Me.ScaleHeight - Imagine.Center_Y
   Raggio = (Me.ScaleWidth / 2) - (ImgMenu(0).ScaleWidth / 2) ' 300
'
   M_Dir01 = 0
   M_Dir02 = 0
   M_Dir03 = 0
   
   Timer_Menu.Enabled = True
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Me.Cls
End Sub
'
'-----------------------------------------------------------------
'                              Trasparenze
'-----------------------------------------------------------------
Private Sub ImgTrasparente(Num)
 
 'Imagine.TrspCol(Num) = &HFFFFFF
  Imagine.TrspCol(Num) = GetPixel(ImgMenu(Num).hdc, 0, 0)
 'TransBltNow(hDestDC As Long, lDestX As Long, lDestY As Long, LWidth As Long, lHeight As Long, hSourceDC As Long, lSourceX As Long, lSourceY As Long, lTransColor As Long)
  TransBltNow Me.hdc, ImgMenu(Num).Left, ImgMenu(Num).Top, ImgMenu(Num).Width, ImgMenu(Num).Height, ImgMenu(Num).hdc, 0, 0, Imagine.TrspCol(Num)

End Sub
'
Private Sub TitoloTrasparente(Num)
  ImgTitolo.Picture = ImgTit(Num)
  ImgTitolo.Left = (Me.ScaleWidth - ImgTitolo.Width) / 2
  Titolo.TrspCol = GetPixel(ImgTitolo.hdc, 0, 0)
  TransBltNow Me.hdc, ImgTitolo.Left, ImgTitolo.Top, ImgTitolo.Width, ImgTitolo.Height, ImgTitolo.hdc, 0, 0, Titolo.TrspCol

End Sub
'
Private Sub ImgSubTrasparente(Num As Long, NumSub As Long)
 Dim ImgSubTrsp As Long
  ImgSub(NumSub).Picture = Img_Sub(Num, NumSub)
  DoEvents
  ImgSubTrsp = GetPixel(ImgSub(NumSub).hdc, 0, 0)
  TransBltNow Me.hdc, ImgSub(NumSub).Left, ImgSub(NumSub).Top, ImgSub(NumSub).Width, ImgSub(NumSub).Height, ImgSub(NumSub).hdc, 0, 0, ImgSubTrsp

End Sub
'
Private Sub PicFondoTrasparente(Num As Long)
 Dim ImgFndTrsp As Long
  PicFondo.Picture = ImgFondo(Num)
  ImgFndTrsp = GetPixel(ImgMenu(0).hdc, 0, 0)
  TransBltNow Me.hdc, 0, 0, PicFondo.Width, PicFondo.Height, PicFondo.hdc, 0, 0, ImgFndTrsp

End Sub
'------------------------------------------------------------------
'  ImgMenuVedi
'------------------------------------------------------------------
Public Sub ImgMenuVedi()
 For ImgNum = 0 To ImgQnt
   ImgTrasparente ImgNum
 Next ImgNum
End Sub
'------------------------------------------------------------------
'                          Timer Lampeggi
'------------------------------------------------------------------
Private Sub Timer_Lmpg_Timer()
     
   Lmpg = Lmpg * -1
   ImgSub(SelSub).Visible = Lmpg + 0.5

End Sub
'------------------------------------------------------------------
'                          Timer Menu
'------------------------------------------------------------------
Private Sub Timer_Menu_Timer()
     
  If RuotaMenu = True Then
       
       If Remote.Pressed = False Then
          TitoloTrasparente (Menu.select)
          Imagine.Corrente = 0
          Timer_Shift.Enabled = True
       End If
       
       WaitKeyMenu
  End If
  
End Sub
'------------------------------------------------------------------
'              Timer Menu OK Ruota Imagine
'------------------------------------------------------------------
Private Sub Timer_Shift_Timer()
  
  If Menu.Ready = False Then
    Timer_Shift.Enabled = False
   Imagine.Corrente = 0
  End If
  
  ImgMenu(Menu.select) = Img(Menu.select, Imagine.Corrente)
   
  BitBltItNow Me.hdc, Imagine.Back_X, Imagine.Back_Y, ImgMenuBack.hdc, ImgMenuBack.Width, ImgMenuBack.Height, 0, 0
  
  ImgTrasparente Menu.select
  Imagine.Corrente = Imagine.Corrente + 1
  If Imagine.Corrente > Imagine.Totale(Menu.select) Then Imagine.Corrente = 0
  Me.Refresh

End Sub
'------------------------------------------------------------------
'                       Ruota Menu
'------------------------------------------------------------------
Public Function RuotaMenu() As Boolean

 Me.Cls
 
 ' Line (Me.ScaleWidth / 2, 0)-(Me.ScaleWidth / 2, Me.ScaleHeight), 0
  '
   BitBltItNow ImgMenuBack.hdc, 0, 0, Me.hdc, ImgMenu(0).Width, ImgMenu(0).Height, Imagine.Back_X, Imagine.Back_Y
'   ImgMenuBack.Refresh
  '
  RuotaMenu = False
  Menu.Ready = False
 '
 For ImgNum = 0 To ImgQnt
   
   ImgMenu(ImgNum).Left = (Pc_X + Raggio * Cos((Rotazione(ImgNum) Mod 360) * 2 * pi / 360 - (0.5 * pi)))
   ImgMenu(ImgNum).Top = (Pc_Y + Raggio * Sin((Rotazione(ImgNum) Mod 360) * 2 * pi / 360 - (0.5 * pi)))
   
   ImgTrasparente ImgNum

   If Abs(Rotazione(ImgNum) - M_Dir01) <> 0 Then ' Piu o meno 1 gradi
     Rotazione(ImgNum) = Rotazione(ImgNum) + (Img_Step * M_Dir03)
   Else
     
     Menu.select = ImgNum
     Rotazione(ImgNum) = Abs(M_Dir02 - Img_Step)
     Imagine.Back_X = ImgMenu(ImgNum).Left
     Imagine.Back_Y = ImgMenu(ImgNum).Top
     RuotaMenu = True
     Menu.Ready = True
   End If
  '
 Next ImgNum

End Function
'------------------------------------------------------------------------
'  Evento Mouse
'------------------------------------------------------------------------
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Button
   Case 1
    Remote.KeyPress = 37
   Case 2
    Remote.KeyPress = 39
   Case 4
    Remote.KeyPress = 13
 End Select
  Remote.Pressed = True
End Sub
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Remote.Pressed = False
  Remote.KeyPress = 0
End Sub
'------------------------------------------------------------------------
' Evento Keyboard
'------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Remote.Pressed = True
 Remote.KeyPress = KeyCode
End Sub
'
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 Remote.Pressed = False
 Remote.KeyPress = 0
End Sub

'--------------------------------------------------------------------
'                       WaitKey Menu
'--------------------------------------------------------------------
Private Function WaitKeyMenu() As Integer
Pippo:
   Do While Remote.KeyPress = 0
      DoEvents                          ' Passa il controllo ad altri processi.
   Loop
 '                                      ' Tasto Premuto
        Select Case Remote.KeyPress
         
         Case Remote.OK                     '  OK
           Remote.KeyPress = 0
           RunSubMenu (Menu.select)
         
         Case Remote.Right, 39              '  Destra
           M_Dir01 = 0
           M_Dir02 = 360
           M_Dir03 = -1
         
         Case Remote.Left, 37               '  Sinistra
           M_Dir01 = 360
           M_Dir02 = 0
           M_Dir03 = 1
         
         Case Remote.Escape, 27
           Unload Me
         
         Case Else
           Remote.KeyPress = 0
           GoTo Pippo
           M_Dir01 = 0
           M_Dir02 = 0
           M_Dir03 = 0
       End Select
       
End Function
'--------------------------------------------------------------------
'                       WaitKey SUB MENU
'--------------------------------------------------------------------
Private Function WaitKeySub() As Integer
  Dim M_Dir As Long
'  Dim SelSub As Long
  Dim SelSub_Max As Long
  
  SelSub_Max = ImgSub.Count - 1
  SelSub = 0
  Timer_Lmpg.Enabled = True
  '
   LblSub(0).Visible = True
   LblSub(1).Visible = True

Pippo:

    ImgSub(SelSub).Visible = False
    SelSub = SelSub + M_Dir
    If SelSub < 0 Then SelSub = SelSub_Max
    If SelSub > SelSub_Max Then SelSub = 0

    ImgSub(SelSub).Visible = True
    
    LblSub(0).Caption = TxtSub(Menu.select, SelSub)
    LblSub(1).Caption = LblSub(0).Caption
    
    LblSub(0).Left = ImgSub(SelSub).Left
    LblSub(1).Left = ImgSub(SelSub).Left + 2

    
'    Me.Refresh
'
   Remote.KeyPress = 0
   Do While Remote.KeyPress = 0
      DoEvents                          ' Passa il controllo ad altri processi.
   Loop
 '                                      ' Tasto Premuto
   M_Dir = 0
        
     Select Case Remote.KeyPress
         
         Case Remote.OK                     '  OK
           Timer_Lmpg.Enabled = False
           Remote.KeyPress = 0
           ImgSub(SelSub).Visible = False
           
           LblSub(0).Visible = False
           LblSub(1).Visible = False
           
           RunShell ((Menu.select + 1) * 100) + SelSub
           Exit Function
         Case Remote.Right, 39              '  Destra
           M_Dir = 1
           Lmpg = 0.5                       ' reset timer
           Timer_Lmpg.Enabled = False       ' reset timer
           Timer_Lmpg.Enabled = True        ' reset timer

         Case Remote.Left, 37               '  Sinistra
           M_Dir = -1
           Lmpg = 0.5                       ' reset timer
           Timer_Lmpg.Enabled = False       ' reset timer
           Timer_Lmpg.Enabled = True        ' reset timer
         
         Case Remote.Escape, 27
           Timer_Lmpg.Enabled = False
           Remote.KeyPress = 0
           ImgSub(SelSub).Visible = False
           
           LblSub(0).Visible = False
           LblSub(1).Visible = False
           
           RunShell -1
           Exit Function
         Case Else
           Remote.KeyPress = 0
    End Select
GoTo Pippo

End Function


'-----------------------------------------------------------------
'                          RunSubMneu
'-----------------------------------------------------------------
Public Sub RunSubMenu(Index As Long)
 
  If Menu.Submenu(Index) = False Then RunShell (Index): Exit Sub ' SubMenu = False
 '
   PicFondoTrasparente (6)

   For i = 0 To ImgSub.Count - 1
    ImgSubTrasparente Menu.select, i
   Next i
  
  Me.Refresh
 '
  WaitKeySub
End Sub
'-----------------------------------------------------------------
'                          RunShell
'-----------------------------------------------------------------
Public Sub RunShell(Index As Long)
 Dim RetVal
 Dim StrCmd As String
 Label3.Caption = "LOADING"
 DoEvents
 On Error Resume Next
 Select Case Index
  Case 1
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
  Case 2
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
  Case 3
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
  Case 4
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
  Case 5
    StrCmd = "C:\windows\Explorer.exe "
    RetVal = Shell(StrCmd, 3)
  Case 6
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
  Case 7
    StrCmd = "C:\Programmi\Internet Explorer\IEXPLORE.EXE"
    RetVal = Shell(StrCmd, 3)
  Case 8
    StrCmd = "C:\Programmi\Outlook Express\msimn.exe"
    RetVal = Shell(StrCmd, 3)
 '
 '   Sub Menu
 '
    Case 100
    StrCmd = "C:\Programmi\Outlook Express\msimn.exe"
    RetVal = Shell(StrCmd, 3)
  Case 101
    StrCmd = "C:\Programmi\Internet Explorer\IEXPLORE.EXE"
    RetVal = Shell(StrCmd, 3)
    Case 200
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
  Case 201
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
   Case 202
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
   Case 203
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
   Case 204
    StrCmd = ""
    RetVal = Shell(StrCmd, 1)
 End Select
 On Error GoTo 0
'
Label3.Caption = ""
Cls
ImgMenuVedi
TitoloTrasparente (Menu.select)
WaitKeyMenu
'Me.Refresh
End Sub


