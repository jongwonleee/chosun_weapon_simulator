VERSION 5.00
Begin VB.Form Start_form 
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   637
   ScaleMode       =   3  'ÇÈ¼¿
   ScaleWidth      =   681
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10080
      Top             =   6120
   End
   Begin VB.PictureBox sbots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2415
      Left            =   6840
      Picture         =   "Start_form.frx":0000
      ScaleHeight     =   157
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   186
      TabIndex        =   5
      Top             =   5760
      Width           =   2850
   End
   Begin VB.PictureBox sbot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2415
      Left            =   6600
      Picture         =   "Start_form.frx":157B4
      ScaleHeight     =   157
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   186
      TabIndex        =   4
      Top             =   5760
      Width           =   2850
   End
   Begin VB.PictureBox titles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5685
      Left            =   240
      Picture         =   "Start_form.frx":2AF68
      ScaleHeight     =   375
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   500
      TabIndex        =   3
      Top             =   5760
      Width           =   7560
   End
   Begin VB.PictureBox title 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5685
      Left            =   120
      Picture         =   "Start_form.frx":B44F0
      ScaleHeight     =   375
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   500
      TabIndex        =   2
      Top             =   5760
      Width           =   7560
   End
   Begin VB.PictureBox mainbg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5685
      Left            =   7680
      Picture         =   "Start_form.frx":13DA78
      ScaleHeight     =   375
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   500
      TabIndex        =   1
      Top             =   0
      Width           =   7560
   End
   Begin VB.PictureBox bg 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5625
      Left            =   0
      ScaleHeight     =   371
      ScaleMode       =   3  'ÇÈ¼¿
      ScaleWidth      =   496
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      Begin VB.Image Image1 
         Height          =   2355
         Left            =   4590
         Top             =   600
         Width           =   2790
      End
   End
End
Attribute VB_Name = "Start_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ttlt, ttll, sbt, sbl, ttlc, ttlud As Integer

Private Sub Form_Load()
sound_init Start_form

ttlt = 0
ttll = -375
ttlc = 0
ttlud = 0
sbt = 600 ' 306
sbl = 40 '186
End Sub

Private Sub Image1_Click()
If ttlc = 7 Then
            stopsound bg_start
            Unload Me
            Main_form.Show
End If
End Sub

Private Sub Timer1_Timer()
bg.Cls
BitBlt bg.hDC, 0, 0, 500, 375, mainbg.hDC, 0, 0, SRCPAINT
Select Case ttlc
    Case 0
        If ttll < 0 Then
                ttll = ttll + 15
        ElseIf ttll >= 0 Then
                ttlc = 1
        End If
    
    Case 1
        playsound eff_boing(1), 2
        If ttll > -200 Then
            If ttlud = 0 Then
                ttll = ttll - 10
            End If
        ElseIf ttll <= -200 Then
            ttlud = 1
        End If
            
            If ttlud = 1 Then
                ttll = ttll + 15
                If ttll >= 0 Then
                    ttlc = 2
                    ttlud = 0
                End If
            End If
    Case 2
        playsound eff_boing(2), 2
        If ttll > -100 Then
            If ttlud = 0 Then
                ttll = ttll - 8
            End If
        ElseIf ttll <= -100 Then
            ttlud = 1
        End If
        
        If ttlud = 1 Then
            ttll = ttll + 15
            If ttll >= 0 Then
                ttlc = 3
                ttlud = 0
            End If
        End If
    Case 3
        playsound eff_boing(3), 3
        If ttll > -50 Then
            If ttlud = 0 Then
                ttll = ttll - 7
            End If
        ElseIf ttll <= -50 Then
            ttlud = 1
        End If
        
        If ttlud = 1 Then
            ttll = ttll + 15
            If ttll >= 0 Then
                ttlc = 4
                ttlud = 0
            End If
        End If
    Case 4
        playsound eff_boing(4), 3
        If ttll > -30 Then
            If ttlud = 0 Then
                ttll = ttll - 6
            End If
        ElseIf ttll <= -30 Then
            ttlud = 1
        End If
        
        If ttlud = 1 Then
            ttll = ttll + 15
            If ttll >= 0 Then
                ttlc = 5
                ttlud = 0
            End If
        End If
    Case 5
        playsound eff_boing(5), 3
        If ttll > -10 Then
            If ttlud = 0 Then
                ttll = ttll - 5
            End If
        ElseIf ttll <= -10 Then
            ttlud = 1
        End If
        
        If ttlud = 1 Then
            ttll = ttll + 15
            If ttll >= 0 Then
                ttlc = 6
                ttlud = 0
            End If
        End If
    Case 5
        playsound eff_boing(6), 2
        If ttll > -5 Then
            If ttlud = 0 Then
                ttll = ttll - 4
            End If
        ElseIf ttll <= -5 Then
            ttlud = 1
        End If
        
        If ttlud = 1 Then
            ttll = ttll + 15
            If ttll >= 0 Then
                ttlc = 6
                ttlud = 0
            End If
        End If
    Case 6
    playsound bg_start, 5
        If sbt > 306 Then
            sbt = sbt - 3
        ElseIf sbt <= 306 Then
            ttlc = 7
        End If
End Select

BitBlt bg.hDC, ttlt, ttll, 500, 375, titles.hDC, 0, 0, SRCPAINT
BitBlt bg.hDC, ttlt, ttll, 500, 375, title.hDC, 0, 0, SRCAND
BitBlt bg.hDC, sbt, sbl, 186, 157, sbots.hDC, 0, 0, SRCPAINT
BitBlt bg.hDC, sbt, sbl, 186, 157, sbot.hDC, 0, 0, SRCAND
End Sub
