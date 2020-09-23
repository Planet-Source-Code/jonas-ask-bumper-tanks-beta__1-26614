VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bumper Tanks"
   ClientHeight    =   4665
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8700
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicKingM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2790
      Left            =   11280
      Picture         =   "Main.frx":0BC2
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   45
      Top             =   9840
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox PicKing 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2790
      Left            =   8940
      Picture         =   "Main.frx":14E9C
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   44
      Top             =   9840
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   7680
      TabIndex        =   43
      Top             =   540
      Width           =   975
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   8700
      Picture         =   "Main.frx":29176
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   42
      Top             =   3120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox PicCrate 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8700
      Picture         =   "Main.frx":29488
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   40
      Top             =   2880
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   10260
      Picture         =   "Main.frx":2979A
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   38
      Top             =   540
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox PicBackM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   10560
      Picture         =   "Main.frx":8F4C4
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   39
      Top             =   600
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox PicCloudM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   4
      Left            =   4140
      Picture         =   "Main.frx":F51EE
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   37
      Top             =   7020
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox PicCloudM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   4140
      Picture         =   "Main.frx":F7860
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   36
      Top             =   6660
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox PicCloudM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   2
      Left            =   5040
      Picture         =   "Main.frx":F81EA
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   35
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PicCloudM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4140
      Picture         =   "Main.frx":FD7B8
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   34
      Top             =   5880
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.PictureBox PicCloudM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Index           =   0
      Left            =   4140
      Picture         =   "Main.frx":FF9AA
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   33
      Top             =   4860
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.PictureBox PicCloud 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   4
      Left            =   1740
      Picture         =   "Main.frx":1075AC
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   32
      Top             =   7020
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox PicCloud 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1740
      Picture         =   "Main.frx":109C1E
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   31
      Top             =   6660
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox PicCloud 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   2
      Left            =   2520
      Picture         =   "Main.frx":10A5A8
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   30
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PicCloud 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1680
      Picture         =   "Main.frx":10FB76
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.PictureBox PicCloud 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Index           =   0
      Left            =   1680
      Picture         =   "Main.frx":111D68
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   28
      Top             =   4860
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   7680
      Picture         =   "Main.frx":11996A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   27
      Top             =   2700
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   7680
      Picture         =   "Main.frx":119DE4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   7680
      Picture         =   "Main.frx":11A25E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   25
      Top             =   3300
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   7680
      Picture         =   "Main.frx":11A6D8
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   8040
      Picture         =   "Main.frx":11AB52
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   23
      Top             =   2700
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   8040
      Picture         =   "Main.frx":11AFCC
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   8040
      Picture         =   "Main.frx":11B446
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   21
      Top             =   3300
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   8040
      Picture         =   "Main.frx":11B8C0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicShellM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   8940
      Picture         =   "Main.frx":11BD3A
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox PicShell 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   8820
      Picture         =   "Main.frx":11BDAC
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.CommandButton cmdMax 
      Caption         =   "Max"
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox PicMuzM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   0
      Left            =   7860
      Picture         =   "Main.frx":11BE1E
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   16
      Top             =   1980
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicMuzM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   1
      Left            =   7860
      Picture         =   "Main.frx":11BEF0
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   15
      Top             =   2100
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicMuzM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   2
      Left            =   7860
      Picture         =   "Main.frx":11BFC2
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   14
      Top             =   2220
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicMuzM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   3
      Left            =   7860
      Picture         =   "Main.frx":11C094
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   13
      Top             =   2340
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicMuz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   3
      Left            =   7620
      Picture         =   "Main.frx":11C166
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicMuz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   2
      Left            =   7620
      Picture         =   "Main.frx":11C238
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   11
      Top             =   2220
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicMuz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   1
      Left            =   7620
      Picture         =   "Main.frx":11C30A
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   10
      Top             =   2100
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicMuz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   0
      Left            =   7620
      Picture         =   "Main.frx":11C3DC
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   9
      Top             =   1980
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicPlayerM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   9000
      Picture         =   "Main.frx":11C4AE
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox PicPlayerM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   0
      Left            =   9000
      Picture         =   "Main.frx":11C720
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   7
      Top             =   1860
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox PicPlayer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   8700
      Picture         =   "Main.frx":11C992
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox PicPlayer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   0
      Left            =   8700
      Picture         =   "Main.frx":11CC04
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4170
      Index           =   2
      Left            =   -120
      Picture         =   "Main.frx":11CE76
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   4
      Top             =   8220
      Width           =   7500
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4170
      Index           =   1
      Left            =   7920
      Picture         =   "Main.frx":182BA0
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   3
      Top             =   5460
      Width           =   7500
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   0
      Left            =   3780
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   2
      Top             =   5100
      Width           =   435
   End
   Begin VB.PictureBox PicMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4230
      Left            =   60
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   1
      Top             =   60
      Width           =   7560
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "New Game"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      Height          =   1695
      Left            =   7680
      TabIndex        =   46
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2040
      TabIndex        =   41
      Top             =   4320
      Width           =   3585
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SeenKing As Boolean
Dim LastTick As String
Dim MaxSpeed As Boolean

Private Sub cmdPause_Click()
    If TheKing > 0 Then Exit Sub 'Cannot halt such a bussy man ;)
    MainPause = Not MainPause
    cmdPause.FontBold = MainPause
    cmdStart.Enabled = Not MainPause
End Sub

Sub MainLoop()
    Do
        Starttick = GetTickCount
        DoEvents
        
        'THIS IS A FRAME LIMITER --------
        NowTime = GetTickCount
        Do Until NowTime - LastTick > 30 Or MaxSpeed
            DoEvents
            NowTime = GetTickCount
        Loop
        LastTick = NowTime
        '----------------------
        ' Pause
        Do Until MainPause = False
            DoEvents
        Loop
        DoKeys 'Porcess the key inputs
        DoCrates 'Tick the crates
        DoPhysics 'Move the tanks around
        MoveShots 'Tick the projectiles
        DoClouds 'Tick the clouds
        DoMessage 'Tick the messages
        DoExplo '... The explosions
        TankThink '..the fireing rate of the tank etc..
        CheckForBlueSuedeShoes 'hehe
        DoGraphics 'Finaly paint the graphics
        
        'Print Player info:
        Dim Txt As String
        Txt = ""
        For A = 1 To UBound(P)
            Txt = Txt & "       Player " & A & ": " & P(A).Points & " "
            For B = 0 To P(A).Life
                Txt = Txt & "|"
            Next B
            Txt = Txt & vbNewLine
            
            If P(A).PUp.SPWeap > 0 Then 'Print super weapon icon
                BitBlt PicMain.hDC, 8, ((A - 1) * 17) + 2, 15, 15, PicIcon(P(A).PUp.SPWeap - 1).hDC, 0, 0, SRCCOPY
            End If
        Next A
        Txt = Mid(Txt, 1, Len(Txt) - 2) 'Remove last VbNewLine
        PicMain.Print Txt
        
        'Frame Rate
        'PicMain.Print "   FPS: " & Int(1000 / (GetTickCount - Starttick + 1))
    Loop
End Sub

Private Sub cmdStart_Click()
    For A = 1 To UBound(P) 'Spawn all the tanks
        SpawnTank A
    Next A
    CreateMessage "Game started", 0
    MainLoop 'Start the loop
End Sub

Private Sub cmdMax_Click()
    MaxSpeed = Not MaxSpeed
End Sub

Private Sub Form_Load()
    Randomize
    LoadPictures 'Set buffers and pictures
    LastTick = GetTickCount 'give the framelimiter
    lblCredits.Caption = "Game and graphics by Jonas Ask"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
Sub CheckForBlueSuedeShoes()
Dim i As Integer
    If SeenKing Then Exit Sub
    For A = 1 To UBound(P)
        If P(A).Life = 0 Then i = i + 1
    Next A
    If i = UBound(P) And GetAsyncKeyState(vbKeyT) And GetAsyncKeyState(vbKeyK) Then
        SeenKing = True
        PlaySound "elvis"
        TheKing = BoardW + 148
    End If
End Sub
