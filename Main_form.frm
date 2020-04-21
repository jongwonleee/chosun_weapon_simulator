VERSION 5.00
Begin VB.Form Main_form 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "Form1"
   ClientHeight    =   11760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   784
   ScaleMode       =   3  'ÇÈ¼¿
   ScaleWidth      =   1021
   Begin VB.PictureBox iconlist 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Left            =   600
      Picture         =   "Main_form.frx":0000
      ScaleHeight     =   118
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   1050
      TabIndex        =   7
      Top             =   9120
      Width           =   15810
   End
   Begin VB.PictureBox iconlists 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1830
      Left            =   480
      Picture         =   "Main_form.frx":5AD24
      ScaleHeight     =   118
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   1050
      TabIndex        =   6
      Top             =   9120
      Width           =   15810
   End
   Begin VB.PictureBox mids 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8910
      Left            =   12600
      Picture         =   "Main_form.frx":B5A48
      ScaleHeight     =   590
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   800
      TabIndex        =   5
      Top             =   120
      Width           =   12060
   End
   Begin VB.PictureBox mid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8910
      Left            =   12600
      Picture         =   "Main_form.frx":20F5CC
      ScaleHeight     =   590
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   800
      TabIndex        =   4
      Top             =   0
      Width           =   12060
   End
   Begin VB.PictureBox bots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8910
      Left            =   12480
      Picture         =   "Main_form.frx":369150
      ScaleHeight     =   590
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   800
      TabIndex        =   3
      Top             =   0
      Width           =   12060
   End
   Begin VB.PictureBox bot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8910
      Left            =   12360
      Picture         =   "Main_form.frx":4C2CD4
      ScaleHeight     =   590
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   800
      TabIndex        =   2
      Top             =   0
      Width           =   12060
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   9120
   End
   Begin VB.PictureBox mainbg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8910
      Left            =   12240
      Picture         =   "Main_form.frx":61C858
      ScaleHeight     =   590
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   0
      Width           =   12060
   End
   Begin VB.PictureBox bg 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8910
      Left            =   0
      ScaleHeight     =   590
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12060
      Begin VB.Label icname 
         BackStyle       =   0  'Åõ¸í
         BeginProperty Font 
            Name            =   "±Ã¼­Ã¼"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5595
         TabIndex        =   8
         Top             =   7920
         Width           =   900
      End
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   0
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "Main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim speed As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private DX As New DirectX7
Private DDraw As DirectDraw7

Dim GameExit, entering, icmoving As Boolean
Dim boty, botyspd, mbmv, midsn, midx, midt, iconx, iconout As Integer
Dim icn, icno As Integer


Private Sub Form_Load()

Set DDraw = DX.DirectDrawCreate("")
DDraw.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_NOWINDOWCHANGES
DDraw.SetDisplayMode 1024, 768, 16, 0, DDSDM_DEFAULT

sound_init Main_form
GameExit = True

boty = -590
mbmv = 1
botyspd = 20
midsn = 0
midx = -800
iconx = 1050
iconout = 1
icn = -4


End Sub


Private Sub Timer1_Timer()
bg.Cls
BitBlt bg.hDC, 0, 0, 800, 590, mainbg.hDC, 0, 0, SRCPAINT


If iconout < 3 Then
Call bot_move
Call icon_starting
End If


Call key_input
Call bot_stay
Call icon_move


End Sub

Private Sub bot_move()
BitBlt bg.hDC, 0, boty, 800, 600, bots.hDC, 0, 0, SRCPAINT
BitBlt bg.hDC, 0, boty, 800, 600, bot.hDC, 0, 0, SRCAND
BitBlt bg.hDC, midx, 0, 800, 600, mids.hDC, 0, 0, SRCPAINT
BitBlt bg.hDC, midx, 0, 800, 600, mid.hDC, 0, 0, SRCAND


Select Case mbmv
Case 1 ' ³ª¹«º¸µå ¶³¾îÁü
    If boty < -100 Then
        boty = boty + botyspd
        playsound chain, 2
    ElseIf boty < -3 Then
        boty = boty + botyspd
        botyspd = botyspd / 10 * 8
    ElseIf boty > -3 Then
        boty = 0
        mbmv = 2
    End If
Case 2 ' ¹öÆ° ¹× È­¸é ³ª¿È
    If midsn < 10 Then
        midt = midt + 1
        If midt > 1 Then
            midx = 0
            midsn = midsn + 1
            mid.Picture = LoadPicture(App.Path & "\image\Main\main_mid\main_mid" & midsn & ".bmp")
            midt = 0
        End If
    ElseIf midsn >= 10 Then
        playsound bg_main, 1
        iconout = 2
    End If
End Select
End Sub

Private Sub bot_stay()
BitBlt bg.hDC, 0, boty, 800, 600, bots.hDC, 0, 0, SRCPAINT
BitBlt bg.hDC, 0, boty, 800, 600, bot.hDC, 0, 0, SRCAND
BitBlt bg.hDC, midx, 0, 800, 600, mids.hDC, 0, 0, SRCPAINT
BitBlt bg.hDC, midx, 0, 800, 600, mid.hDC, 0, 0, SRCAND

End Sub

Private Sub icon_starting()
BitBlt bg.hDC, 87, 435, 631, 118, iconlists.hDC, iconx, 0, SRCPAINT
BitBlt bg.hDC, 87, 435, 631, 118, iconlist.hDC, iconx, 0, SRCAND
If iconout = 2 Then
    If iconx > 200 Then
        iconx = iconx - 8
    End If
    If iconx <= 200 Then
        If iconx > -280 Then
            iconx = iconx - 5
        End If
    End If
    If iconx <= -280 Then
        iconx = -280
        iconout = 3
        icname.Left = 383
        icname = "Á¤º¸"
    End If
End If
End Sub

Private Sub icon_move()
BitBlt bg.hDC, 87, 435, 631, 118, iconlists.hDC, iconx, 0, SRCPAINT
BitBlt bg.hDC, 87, 435, 631, 118, iconlist.hDC, iconx, 0, SRCAND

If icmoving = True Then
    If icn < icno Then
        If icn * 70 < iconx Then
            iconx = iconx - 5
        ElseIf icn * 70 >= iconx Then
            icmoving = False
        End If
    End If
    If icn > icno Then
        If icn * 70 > iconx Then
            iconx = iconx + 5
        ElseIf icn * 70 >= iconx Then
            icmoving = False
        End If
    End If
    
    Select Case icn
        Case -4
            icname.Left = 383
            icname.Width = 57
            icname = "Á¤º¸"
        Case -3
            icname.Left = 383
            icname.Width = 57
            icname = "Æí°ï"
        Case -2
            icname.Left = 383
            icname.Width = 57
            icname = "Çùµµ"
        Case -1
            icname.Left = 383
            icname.Width = 57
            icname = "È¯µµ"
        Case 0
            icname.Left = 383
            icname.Width = 57
            icname = "ÀåÃ¢"
        Case 1
            icname.Left = 383
            icname.Width = 57
            icname = "´ÉÃ³"
        Case 2
            icname.Left = 383
            icname.Width = 57
            icname = "ÃµÀÚÃÑÅë"
        Case 3
            icname.Left = 383
            icname.Width = 57
            icname = "ÆíÀü"
        Case 4
            icname.Left = 373
            icname.Width = 60
            icname = "¼ö³ë±Ã"
        Case 5
            icname.Left = 383
            icname.Width = 57
            icname = "½ÂÀÚÃÑÅë"
        Case 6
            icname.Left = 373
            icname.Width = 60
            icname = "½Å±âÀü"
        Case 7
            icname.Left = 373
            icname.Width = 60
            icname = "Èæ°¢±Ã"
        Case 8
            icname.Left = 383
            icname = "Á¶ÃÑ"
        Case 9
            icname.Left = 373
            icname.Width = 60
            icname = "ÆÇ¿Á¼±"
        Case 10
            icname.Left = 373
            icname.Width = 60
            icname = "°ÅºÏ¼±"
    End Select
End If
End Sub

Private Sub key_input()
If iconout = 3 Then
    If icmoving = False Then
        If GetKeyState(vbKeyLeft) < 0 Then
            If icn > -4 Then
                icmoving = True
                icno = icn
                icn = icn - 1
            End If
        ElseIf GetKeyState(vbKeyRight) < 0 Then
            If icn < 10 Then
                icmoving = True
                icno = icn
                icn = icn + 1
            End If
        End If
    End If
End If
End Sub
