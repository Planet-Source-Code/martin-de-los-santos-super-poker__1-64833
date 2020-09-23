VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frm_main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00078118&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M@rtsoft->Juegos->Poker"
   ClientHeight    =   5475
   ClientLeft      =   1245
   ClientTop       =   1020
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPp_Ju_Poker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   400
      Left            =   3360
      Top             =   120
   End
   Begin VB.PictureBox Picdob 
      Appearance      =   0  'Flat
      BackColor       =   &H0007A618&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   360
      ScaleHeight     =   1575
      ScaleWidth      =   5295
      TabIndex        =   37
      Top             =   1320
      Width           =   5290
      Begin VB.CommandButton Doblar 
         BackColor       =   &H0007A618&
         Caption         =   "Plantarse"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1440
         Picture         =   "FrmPp_Ju_Poker.frx":0442
         TabIndex        =   40
         ToolTipText     =   "Presione este botón si desea parar de doblar la apuesta"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Cmayor 
         BackColor       =   &H0007A618&
         Caption         =   "Mayor "
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         ToolTipText     =   "Presione este botón si cree que es mayor"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Cmenor 
         BackColor       =   &H0007A618&
         Caption         =   "Menor"
         Height          =   255
         Left            =   2760
         TabIndex        =   38
         ToolTipText     =   "Presione este botón si cree que es manor"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label EtDoSig 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   2640
         TabIndex        =   43
         Top             =   240
         Width           =   330
      End
      Begin VB.Image carta 
         Height          =   1155
         Index           =   6
         Left            =   4150
         Stretch         =   -1  'True
         Top             =   240
         Width           =   945
      End
      Begin VB.Image carta 
         Height          =   1155
         Index           =   5
         Left            =   210
         Stretch         =   -1  'True
         Top             =   240
         Width           =   945
      End
      Begin VB.Label EtImApDo 
         BackStyle       =   0  'Transparent
         Caption         =   "    "
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   3050
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.Label EtApDo 
         BackStyle       =   0  'Transparent
         Caption         =   "Apuesta doblada:"
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   1320
         TabIndex        =   41
         Top             =   270
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   1380
         Index           =   5
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1125
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   1380
         Index           =   6
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1125
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H0007A618&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   855
         Left            =   1320
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   2655
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H0007A618&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         Height          =   1575
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   5300
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4560
      Top             =   120
   End
   Begin VB.CommandButton cmd_jugar 
      BackColor       =   &H0007A618&
      Caption         =   "Repartir"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Picture         =   "FrmPp_Ju_Poker.frx":09CC
      TabIndex        =   21
      Top             =   3720
      Width           =   1215
   End
   Begin VB.OptionButton opt_1 
      BackColor       =   &H0007A618&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      MaskColor       =   &H008080FF&
      TabIndex        =   15
      Top             =   675
      Value           =   -1  'True
      Width           =   200
   End
   Begin VB.OptionButton opt_5 
      BackColor       =   &H0007A618&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   675
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   4080
      Top             =   120
   End
   Begin VB.OptionButton opt_10 
      BackColor       =   &H0007A618&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   675
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   7200
      X2              =   9480
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Image ImaCor 
      Height          =   195
      Left            =   5280
      Picture         =   "FrmPp_Ju_Poker.frx":0F56
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Impic 
      Height          =   195
      Left            =   5640
      Picture         =   "FrmPp_Ju_Poker.frx":1398
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image ImTre 
      Height          =   195
      Left            =   6000
      Picture         =   "FrmPp_Ju_Poker.frx":17DA
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image ImRom 
      Height          =   195
      Left            =   6360
      Picture         =   "FrmPp_Ju_Poker.frx":1C1C
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label EtEjPo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                                            Tabla de puntuaciones"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   585
      Left            =   7200
      TabIndex        =   53
      Top             =   360
      Width           =   2340
   End
   Begin VB.Label EtEjP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   9090
      TabIndex        =   52
      Top             =   600
      Width           =   165
   End
   Begin VB.Label EtEjP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   51
      Top             =   600
      Width           =   165
   End
   Begin VB.Label EtEjP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   8200
      TabIndex        =   50
      Top             =   600
      Width           =   165
   End
   Begin VB.Label EtEjP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   7780
      TabIndex        =   49
      Top             =   600
      Width           =   165
   End
   Begin VB.Image ImEjC1 
      Height          =   195
      Left            =   7520
      Picture         =   "FrmPp_Ju_Poker.frx":205E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   195
   End
   Begin VB.Image ImEjC5 
      Height          =   195
      Left            =   9240
      Picture         =   "FrmPp_Ju_Poker.frx":24A0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   195
   End
   Begin VB.Image ImEjC4 
      Height          =   195
      Left            =   8840
      Picture         =   "FrmPp_Ju_Poker.frx":28E2
      Stretch         =   -1  'True
      Top             =   600
      Width           =   195
   End
   Begin VB.Image ImEjC3 
      Height          =   200
      Left            =   8400
      Picture         =   "FrmPp_Ju_Poker.frx":2D24
      Stretch         =   -1  'True
      Top             =   600
      Width           =   200
   End
   Begin VB.Image ImEjC2 
      Height          =   195
      Left            =   7960
      Picture         =   "FrmPp_Ju_Poker.frx":3166
      Stretch         =   -1  'True
      Top             =   600
      Width           =   195
   End
   Begin VB.Label EtEjP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   7330
      TabIndex        =   48
      Top             =   600
      Width           =   165
   End
   Begin VB.Image Im_Son 
      Height          =   240
      Index           =   1
      Left            =   240
      Picture         =   "FrmPp_Ju_Poker.frx":35A8
      Top             =   165
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Im_Son 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "FrmPp_Ju_Poker.frx":3B32
      Top             =   165
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image In_Int 
      Height          =   280
      Index           =   0
      Left            =   600
      Picture         =   "FrmPp_Ju_Poker.frx":40BC
      Stretch         =   -1  'True
      ToolTipText     =   "Idioma Español"
      Top             =   120
      Width           =   400
   End
   Begin VB.Image In_Int 
      Height          =   285
      Index           =   1
      Left            =   600
      Picture         =   "FrmPp_Ju_Poker.frx":4AFC
      Stretch         =   -1  'True
      ToolTipText     =   "Language English"
      Top             =   120
      Width           =   405
   End
   Begin VB.Image In_Din 
      Height          =   285
      Index           =   0
      Left            =   1155
      Picture         =   "FrmPp_Ju_Poker.frx":56D9
      Stretch         =   -1  'True
      ToolTipText     =   "Coin"
      Top             =   100
      Width           =   435
   End
   Begin VB.Label EtAuBa 
      BackStyle       =   0  'Transparent
      Height          =   450
      Index           =   1
      Left            =   585
      TabIndex        =   46
      Top             =   0
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0007A618&
      X1              =   0
      X2              =   8520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label EtBe 
      BackColor       =   &H00078118&
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar tipo moneda, solo se permite el cambio cuando no se juega en red por dinero y cuando es un juego nuevo "
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   5160
      Width           =   9375
   End
   Begin VB.Label EtApu 
      BackStyle       =   0  'Transparent
      Caption         =   "$ 10"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   3130
      TabIndex        =   17
      Top             =   735
      Width           =   720
   End
   Begin VB.Label EtApu 
      BackStyle       =   0  'Transparent
      Caption         =   "$ 5"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   2290
      TabIndex        =   18
      Top             =   735
      Width           =   720
   End
   Begin VB.Image Scar 
      Height          =   780
      Left            =   3000
      Picture         =   "FrmPp_Ju_Poker.frx":59E3
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   645
   End
   Begin VB.Label Acerca2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed for: Martin DLS"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6720
      TabIndex        =   35
      Top             =   3975
      Width           =   2535
   End
   Begin VB.Image Acerca5 
      Height          =   900
      Left            =   6000
      Picture         =   "FrmPp_Ju_Poker.frx":6038
      Top             =   3720
      Width           =   915
   End
   Begin VB.Label Acerca7 
      BackStyle       =   0  'Transparent
      Caption         =   "Your questions-More SW..."
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6960
      TabIndex        =   36
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Acerca6 
      Height          =   240
      Left            =   9120
      Picture         =   "FrmPp_Ju_Poker.frx":63F5
      Stretch         =   -1  'True
      Top             =   3630
      Width           =   240
   End
   Begin VB.Label Acerca4 
      BackStyle       =   0  'Transparent
      Caption         =   "martsoft@gmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6840
      TabIndex        =   33
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Acerca0 
      BackStyle       =   0  'Transparent
      Caption         =   "MdlS-POKER Version 1.2"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6840
      TabIndex        =   34
      Top             =   3750
      Width           =   2295
   End
   Begin VB.Label etProgmer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sobre el programador"
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   7560
      TabIndex        =   20
      ToolTipText     =   "Martin-DLS"
      Top             =   4920
      Width           =   1605
   End
   Begin VB.Image carta 
      Height          =   1395
      Index           =   2
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   1455
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   375
      Left            =   5040
      TabIndex        =   32
      Top             =   360
      Width           =   375
      _cx             =   661
      _cy             =   661
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Transparent"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nada...................-X2                                   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   0
      Left            =   7350
      TabIndex        =   31
      Tag             =   "-2"
      ToolTipText     =   "Pierde dos veces lo apostado"
      Top             =   3120
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Royal Flush...........X8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   9
      Left            =   7350
      TabIndex        =   30
      Tag             =   "8"
      ToolTipText     =   "Multiplica lo apostado por ocho"
      Top             =   960
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Poker...................X6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   7
      Left            =   7350
      TabIndex        =   29
      Tag             =   "6"
      ToolTipText     =   "Multiplica lo apostado por ceis"
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Escalera Limpia.....X7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   8
      Left            =   7350
      TabIndex        =   28
      Tag             =   "7"
      ToolTipText     =   "Multiplica lo apostado por siete"
      Top             =   1200
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Escalera Sucia......X3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   4
      Left            =   7350
      TabIndex        =   27
      Tag             =   "3"
      ToolTipText     =   "Multiplica lo apostado por tres"
      Top             =   2160
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Color....................X4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   5
      Left            =   7350
      TabIndex        =   26
      Tag             =   "4"
      ToolTipText     =   "Multiplica lo apostado por cuatro"
      Top             =   1920
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Full......................X5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   6
      Left            =   7350
      TabIndex        =   25
      Tag             =   "5"
      ToolTipText     =   "Multiplica lo apostado por cinco"
      Top             =   1680
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Par Doble ............X1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   2
      Left            =   7350
      TabIndex        =   24
      Tag             =   "1"
      ToolTipText     =   "Multiplica lo apostado por uno"
      Top             =   2640
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pierna..................X2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   3
      Left            =   7350
      TabIndex        =   23
      Tag             =   "2"
      ToolTipText     =   "Multiplica lo apostado por dos"
      Top             =   2400
      Width           =   2100
   End
   Begin VB.Label EtDes 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Par simple.............Nada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   1
      Left            =   7350
      TabIndex        =   22
      Tag             =   "Nada"
      ToolTipText     =   "No pierde ni gana nada "
      Top             =   2880
      Width           =   2025
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   3015
      Left            =   7200
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   3100
      Picture         =   "FrmPp_Ju_Poker.frx":653F
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   3050
      Picture         =   "FrmPp_Ju_Poker.frx":6B94
      Stretch         =   -1  'True
      Top             =   4150
      Width           =   675
   End
   Begin VB.Label EtApu 
      Alignment       =   2  'Center
      BackColor       =   &H0007A618&
      BackStyle       =   0  'Transparent
      Caption         =   "Apuesta:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   855
   End
   Begin VB.Label EtApu 
      BackStyle       =   0  'Transparent
      Caption         =   "$ 1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   16
      Top             =   735
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARTA 5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARTA 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARTA 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARTA 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARTA 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label estado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label estado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label estado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image carta 
      Height          =   1395
      Index           =   4
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Image carta 
      Height          =   1395
      Index           =   3
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Image carta 
      Height          =   1395
      Index           =   1
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Image carta 
      Height          =   1395
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   0
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   1245
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   1
      Left            =   1710
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   1245
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   3
      Left            =   4340
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   1245
   End
   Begin VB.Label estado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   4
      Left            =   5670
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   1245
   End
   Begin VB.Label ErDinero 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Dinero:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   945
   End
   Begin VB.Label dinero 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6000
      TabIndex        =   6
      Top             =   705
      Width           =   930
   End
   Begin VB.Label EtDin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   5520
      TabIndex        =   12
      Top             =   720
      Width           =   450
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label estado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1500
      Index           =   2
      Left            =   3025
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   2295
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Shape Acerca1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   3150
      Picture         =   "FrmPp_Ju_Poker.frx":71E9
      Stretch         =   -1  'True
      Top             =   4250
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   8
      Left            =   3240
      Picture         =   "FrmPp_Ju_Poker.frx":783E
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   675
   End
   Begin VB.Shape Acerca3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H80000010&
      FillStyle       =   2  'Horizontal Line
      Height          =   1215
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   3495
   End
   Begin VB.Label EtAuBa 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   45
      Top             =   0
      Width           =   400
   End
   Begin VB.Label EtAuBa 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   1150
      TabIndex        =   47
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      Height          =   420
      Left            =   1080
      Top             =   45
      Width           =   600
   End
   Begin VB.Shape ShIdi 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      Height          =   420
      Left            =   560
      Top             =   45
      Width           =   540
   End
   Begin VB.Shape ShSon 
      BackColor       =   &H0007A618&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      Height          =   420
      Left            =   120
      Top             =   45
      Width           =   450
   End
   Begin VB.Menu mnu_juego 
      Caption         =   "Juego"
      Begin VB.Menu mnu_nuevo 
         Caption         =   "Nuevo"
         Begin VB.Menu mnjr 
            Caption         =   "Jugar en Red por dinero"
         End
         Begin VB.Menu msjdemo 
            Caption         =   "Localmente practica"
            Begin VB.Menu msNuC100 
               Caption         =   "Cargar $ 100"
            End
            Begin VB.Menu msNuCm 
               Caption         =   "Cargar manualmente"
            End
         End
      End
      Begin VB.Menu se 
         Caption         =   "-"
      End
      Begin VB.Menu mnidioma 
         Caption         =   "Idioma"
         Begin VB.Menu mnIdEn 
            Caption         =   "English"
         End
         Begin VB.Menu mnIdEsp 
            Caption         =   "Español"
         End
      End
      Begin VB.Menu se1 
         Caption         =   "-"
      End
      Begin VB.Menu mnop 
         Caption         =   "Opciones"
         Begin VB.Menu mnSonido 
            Caption         =   "Sonido"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnMoneda 
            Caption         =   "Moneda"
            Begin VB.Menu mnDolar 
               Caption         =   "Dolares"
            End
            Begin VB.Menu MnEuro 
               Caption         =   "Euros"
            End
            Begin VB.Menu MnPesos 
               Caption         =   "Pesos"
            End
         End
      End
      Begin VB.Menu kjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
'Module/Modulo: frm_main
'Author/Autor : Mdls --> (martsoft@gmail.com)
'You can utilize this code stop that to you seem him necessary ... enjoy it!!!
'Puedes utilizar este código para lo que a usted le parezca necesario ...disfrutadlo!!!
'---------------------------------------------------------------------------------------
Dim i As Integer
Dim Maso(51) As Integer  'Array with Pack of cards/Maso de cartas
Dim CartasDm(4) As Integer    'cards than after mixing/Cartas que despues de mezclar
Dim CartasDmDo(1) As Byte    'cards than after mixing/Cartas que despues de mezclar
Dim pinta(4) As String   'la pinta de las cartas finales
Dim saldo As Integer     'lo que queda de plata
Dim apuesta As Byte   'lo que apuesta
Dim pcarta As String     'esto es para buscar la carta
Dim a As Integer         'llama funcion barajar
Dim njugadas As Integer  'el numero de veces que juega
Dim cambios As Integer   'numero de cartas cambiadas
Dim SeCC(4) As Byte    'Array que indica cuando una carta está marcada para cambiarla o no
Dim resultado As Integer    'guarda el resultado final
Dim nada As Boolean      'no tiene nada
Dim IdcDes As Byte       'indice para descripcion marcada cuando titila
Private SeRep As Byte    'seña repartiendo
Private SeApostando As Byte    'seña de estar apostando
Private SeInfc As Byte
Private IdxRec As Long
Public SeSonido As Byte
Dim AuxApuesta As Integer    'Auxiliar del importe de la apuesta
Dim SeDoblando As Byte    'se esta doblando la apuesta
Private CartaADoblar(1) As Integer    'Las dos cartas a comparar para doblar
Private MarcarCartas(4) As Byte
Dim SeAcerca As Byte
Dim SinbMoneda As String * 3
Private ArrEjPo(0 To 9) As Byte

Private SeComNuevo As Byte 'Seña nuevo juego para poder cambiar moneda
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const SW_SHOWNORMAL = 1


Private Sub Acerca2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Acerca4.FontBold = False
End Sub
Private Sub Acerca3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Acerca4.FontBold = False
End Sub
Private Sub Acerca4_Click()
    ShellExecute Me.hwnd, vbNullString, "mailto:martsoft@gmail.com?subject=Mdls-->Game-->POKER", vbNullString, "C:\", SW_SHOWNORMAL    'Shell ("explore.exe mailto:martsoft@gmail.com")
End Sub
Private Sub Acerca4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Acerca4.FontBold = True
End Sub
Private Sub Acerca5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Acerca4.FontBold = False
End Sub
Private Sub Acerca6_Click()
    SeAcerca = 0
    AcercaDe ByVal False
End Sub
Private Sub carta_Click(Index As Integer)
    If Index > 4 Or SeDoblando = 1 Then Exit Sub
    If SeRep = 1 Then Exit Sub
    Dim Respuesta As Long
    Dim TmpCamOman(1) As String: TmpCamOman(0) = LoadResString(SeInfc + 123)
    TmpCamOman(1) = LoadResString(SeInfc + 125)
    Dim TmpRep$: TmpRep$ = LoadResString(SeInfc + 5)
    If cmd_jugar.Caption = TmpRep$ Then
        If SeInfc = 0 Then
            Son ByVal "msrepcar", ByVal "WAVSSISTEMAGRAL"
            DarVueltaCartas
            Respuesta = MsgBox("¡Debes repartir las cartas!" & Chr(13) & "¿Deseas repartir ahora?", vbYesNo + vbInformation + vbDefaultButton1, "Poker")
        Else
            Son ByVal "in_msrepcar", ByVal "WAVSSISTEMAGRAL"
            DarVueltaCartas
            Respuesta = MsgBox("You must distribute the cards!" & Chr(13) & "Do you wish to do it now?", vbYesNo + vbInformation + vbDefaultButton1, "Poker")

        End If
        If Respuesta = vbYes Then   ' El usuario eligió el botón Sí.
            Call cmd_jugar_Click: Exit Sub
        Else
            Exit Sub
        End If
    End If
    If njugadas Mod 2 = 0 Then
        If estado(Index).Caption = TmpCamOman(1) Then    'Cancelo cambiar
            estado(Index).Caption = TmpCamOman(0)
            estado(Index).BackColor = &H8000&
            Label2(Index).BackColor = &H8000&
            SeCC(Index) = 0    'Seña de hay que canceló cambiar la carta, pues estaba marcada para cambiarla
            cambios = cambios - 1
        Else
            estado(Index).Caption = TmpCamOman(1)    '"Cambiar"
            estado(Index).BackColor = &HFF&
            Label2(Index).BackColor = &HFF&
            SeCC(Index) = 1    'Seña de hay que cambiar la carta
        End If
    End If
End Sub

Private Sub carta_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Acerca1.Visible = True Then AcercaDe ByVal False
    Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
    Set carta(Index).MouseIcon = LoadResPicture(105, vbResCursor)
End Sub
Private Sub carta_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set carta(Index).MouseIcon = LoadResPicture(102, vbResCursor)
MsgBE ByVal 11

End Sub
Private Sub carta_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set carta(Index).MouseIcon = LoadResPicture(106, vbResCursor)
End Sub
Private Sub Cmayor_Click()
    DesDobct 0
    DarParaDoblar 1
    CartaCara ByVal CByte(6)
    If CartaADoblar(1) > CartaADoblar(0) Then
        AuxApuesta = AuxApuesta * 2    'manor
        EtImApDo.Caption = AuxApuesta & ".-"
        If msgApDoblada(AuxApuesta) = 1 Then
            DarVueltaCarta 6
            DarParaDoblar 0
            CartaCara ByVal CByte(5)
        Else
            Call Doblar_Click(0)
        End If
    Else
        If SeInfc = 0 Then
            MsgBox "La carta no es Menor, ¡Pierdes: " & SinbMoneda & EtImApDo.Caption & "!", vbInformation, "Doblar Apuesta"
        Else
            MsgBox "The card is not Minor, You Lose: " & SinbMoneda & EtImApDo.Caption & "!", vbInformation, "Bending Bet"
        End If

        EtImApDo.Caption = " 0.-"
        ActiDoblar ByVal 0
        CarDobDV
        cmd_jugar.Enabled = True
        DarVueltaCartas
        Deshop ByVal 0
    End If
    DesDobct 1
End Sub

Private Sub Cmayor_GotFocus()
MsgBE ByVal 2
End Sub

Private Sub Cmayor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 2
    CmdMm Cmayor, 1
End Sub

Private Sub cmd_jugar_Click()
    Dim SeDoblar As Byte
    SeDoblar = 0
    SeRep = 1
    cmd_jugar.Enabled = False
    Timer1.Enabled = False
    MarcarTablaDesc 1, 1
    ' Cls
    ' CartaCara
    If saldo <= 0 Then
        'MsgBox "No tines nada que apostar"'/You do not have nothing that to bet
        Dim Resp As Long
        If SeInfc = 0 Then
            Son ByVal "MsgTerJue", ByVal "WAVSSISTEMAGRAL"
            Resp = MsgBox("El juego Terminó!!!" & Chr(13) & "¿Deseas comenzar otro juego?", vbExclamation + vbYesNo, "POKER MDLS")

            If Resp = 6 Then msNuC100_Click

        Else
            Son ByVal "In_MsgTerJue", ByVal "WAVSSISTEMAGRAL"
            Resp = MsgBox("Game Over!!!" & Chr(13) & " Do you wish to begin another game?", vbExclamation + vbYesNo, "POKER MDLS")

            If Resp = 6 Then msNuC100_Click
        End If
        ColEstCarta
        Exit Sub
    End If
    If njugadas Mod 2 <> 0 Then
        cmd_jugar.Tag = "0"
        Deshop ByVal 1
        cambios = 0
        For i = 0 To 4
            If SeInfc = 0 Then
                estado(i).Caption = LoadResString(123)
            Else
                estado(i).Caption = LoadResString(124)
            End If

        Next i
        a = barajar(Maso(), 51)
        For i = 0 To 4
            CartasDm(i) = Maso(51 - i)
        Next i
        For i = 0 To 4
            DoEvents
            pcarta = ""
            pcarta = cargar(CartasDm(i))

            Son ByVal "repartir", ByVal "WAVSSISTEMAGRAL"
            'Sleep 110
            EfectoRepartir CByte(i) + 1
            carta(i) = LoadPicture(App.Path & "\CARTAS\" + pcarta)
            CartaCara ByVal CByte(i)
        Next i
        '  CartaCara
        '  cmd_jugar.Enabled = True
        'IIf SeInfc = 0, cmd_jugar.Caption = LoadResString(7), cmd_jugar.Caption = LoadResString(8)
        If SeInfc = 0 Then
            cmd_jugar.Caption = LoadResString(7)
        Else
            cmd_jugar.Caption = LoadResString(8)
        End If

        Debug.Print LoadResString(8)
    Else
       If SeComNuevo = 1 Then SeComNuevo = 0 'No es nuevo juego
        cmd_jugar.Tag = "1"
        Deshop ByVal 0
        For i = 0 To 4
            If SeCC(i) = 1 Then    'cambiar
                CartasDm(i) = Maso(46 - i)
                pcarta = ""
                pcarta = cargar(CartasDm(i))
                Son ByVal "repartir", ByVal "WAVSSISTEMAGRAL"
                EfectoRepartir CByte((i)) + 1
                carta(i) = LoadPicture(App.Path & "\CARTAS\" + pcarta)
            End If
            SeCC(i) = 0
        Next i
        'evalua la apuesta
        nada = True
        If opt_1 = True Then
            apuesta = 1
        ElseIf opt_5 = True Then
            apuesta = 5
        Else: apuesta = 10
        End If
        a = convertiryord(CartasDm())
        resultado = unpar(CartasDm())
        If resultado = 1 Then
            saldo = saldo
            If SeInfc = 0 Then
                ' MsgBox "Tienes par simple"'you have one pair
                Son ByVal "hepar", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_hepar", ByVal "WAVSSISTEMAGRAL"
            End If
            MarcarTablaDesc 1
            'dinero.Caption = Str(saldo)
            nada = False
        End If
        resultado = dospares(CartasDm())
        If resultado = 2 Then
            AuxApuesta = apuesta
            saldo = saldo + apuesta
            MarcarTablaDesc 2
            If SeInfc = 0 Then
                Son ByVal "hepardo", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_hepardo", ByVal "WAVSSISTEMAGRAL"
            End If
            ' MsgBox "Tienes par doble"
            dinero.Caption = Str(saldo)
            nada = False
        End If
        resultado = trio(CartasDm())
        If resultado = 3 Then
            AuxApuesta = (apuesta * 2)
            saldo = saldo + (apuesta * 2)
            MarcarTablaDesc 3
            ' MsgBox "Tienes Pierna"'you have tree of king
            If SeInfc = 0 Then
                Son ByVal "hepierna", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_hepierna", ByVal "WAVSSISTEMAGRAL"
            End If
            dinero.Caption = Str(saldo)
            nada = False
        End If
        resultado = poker(CartasDm())
        If resultado = 4 Then
            AuxApuesta = (apuesta * 6)
            saldo = saldo + (apuesta * 6)
            MarcarTablaDesc 7
            ' MsgBox "Tienes un poker"'four of the king
            If SeInfc = 0 Then
                Son ByVal "hepoker", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_hepoker", ByVal "WAVSSISTEMAGRAL"

            End If
            dinero.Caption = Str(saldo)
            nada = False
        End If
        resultado = full(CartasDm())
        If resultado = 5 Then
            AuxApuesta = (apuesta * 5)
            saldo = saldo + (apuesta * 5)
            MarcarTablaDesc 6

            If SeInfc = 0 Then
                Son ByVal "hefull", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_hefull", ByVal "WAVSSISTEMAGRAL"
            End If

            ' MsgBox "Tienes un full"
            dinero.Caption = Str(saldo)
            nada = False
        End If
        resultado = cosuclimreal(pinta(), CartasDm())
        If resultado = 6 Then
            AuxApuesta = (apuesta * 4)
            saldo = saldo + (apuesta * 4)

            If SeInfc = 0 Then
                Son ByVal "hecolor", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_hecolor", ByVal "WAVSSISTEMAGRAL"
            End If

            MarcarTablaDesc 5
            'MsgBox "Tienes Color"
            dinero.Caption = Str(saldo)
            nada = False
        End If
        If resultado = 7 Then
            AuxApuesta = (apuesta * 7)
            saldo = saldo + (apuesta * 7)
            If SeInfc = 0 Then
                Son ByVal "HEESCLIM", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_HEESCLIM", ByVal "WAVSSISTEMAGRAL"
            End If
            MarcarTablaDesc 7
            ' MsgBox "Tienes Limpia"'you have Straight Flush!
            dinero.Caption = Str(saldo)
            nada = False
        End If
        If resultado = 8 Then
            AuxApuesta = (apuesta * 8)
            saldo = saldo + (apuesta * 8)
            MarcarTablaDesc 9
            If SeInfc = 0 Then
                Son ByVal "heescreal", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_heescreal", ByVal "WAVSSISTEMAGRAL"
            End If
            ' MsgBox "Tienes Real"'Royal Flush
            dinero.Caption = Str(saldo)
            nada = False
        End If
        If resultado = 9 Then
            AuxApuesta = (apuesta * 3)
            saldo = saldo + (apuesta * 3)
            MarcarTablaDesc 4
            If SeInfc = 0 Then
                Son ByVal "heescsuc", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "In_heescsuc", ByVal "WAVSSISTEMAGRAL"
            End If
            '  MsgBox "Tienes sucia"'Straight
            nada = False
        End If
        If nada <> True Then
            If AuxApuesta <> 0 Then
                If msgDoblarAp((AuxApuesta)) = 1 Then    'doblo la apuesta
                    '   DesHCtlDob 0
                    ActiDoblar ByVal 1
                    DarParaDoblar 0
                    CartaCara ByVal CByte(5)
                    EtImApDo.Caption = CStr(AuxApuesta)
                    SeDoblar = 1    ' DarParaDoblar 1 '    Image2.ZOrder 1
                    cmd_jugar.Enabled = False    'CartaCara ByVal CByte(6)
                    Deshop ByVal 1
                Else
                    dinero.Caption = Str(saldo)
                End If
            End If
            'Bending the bet
            'pregunta si dobla la apuesta

        Else
            AuxApuesta = (apuesta * 2)
            saldo = saldo - (apuesta * 2)
            MarcarTablaDesc 0

            If SeInfc = 0 Then
                Son ByVal "henada", ByVal "WAVSSISTEMAGRAL"
            Else
                Son ByVal "in_henada", ByVal "WAVSSISTEMAGRAL"
            End If
            ' MsgBox "No tienes nada"'you haven't nothing!
            dinero.Caption = Str(saldo)
        End If
        If SeInfc = 0 Then
            cmd_jugar.Caption = LoadResString(5)
        Else
            cmd_jugar.Caption = LoadResString(6)
        End If
        cambios = 0
        For i = 0 To 4
            If SeInfc = 0 Then
                estado(i).Caption = LoadResString(123)
            Else
                estado(i).Caption = LoadResString(124)
            End If
        Next i
        ColEstCarta
        'DarVueltaCartas
    End If
    njugadas = njugadas + 1

    SeRep = 0
    If SeDoblar <> 1 Then AuxApuesta = 0: cmd_jugar.Enabled = True
End Sub
Private Sub cmd_jugar_KeyDown(KeyCode As Integer, Shift As Integer)
    Son ByVal "click", ByVal "WAVSSISTEMAGRAL"
End Sub
Private Sub cmd_jugar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Son ByVal "click", ByVal "WAVSSISTEMAGRAL"
End Sub

Private Sub cmd_jugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 4
    CmdMm cmd_jugar, 1
End Sub

Private Sub Cmenor_Click()
    DesDobct 0
    DarParaDoblar 1
    CartaCara ByVal CByte(6)
    If CartaADoblar(1) < CartaADoblar(0) Then
        AuxApuesta = AuxApuesta * 2    'manor
        EtImApDo.Caption = AuxApuesta & ".-"
        If msgApDoblada(AuxApuesta) = 1 Then
            DarVueltaCarta 6
            DarParaDoblar 0
            CartaCara ByVal CByte(5)
        Else
            Call Doblar_Click(0)
        End If
    Else
        If SeInfc = 0 Then
            MsgBox "La carta no es Mayor, ¡Pierdes: " & SinbMoneda & EtImApDo.Caption & "!", vbInformation, "Doblar Apuesta"
        Else
            MsgBox "The card is not Major, You Lose: " & SinbMoneda & EtImApDo.Caption & "!", vbInformation, "Bending Bet"
        End If
        EtImApDo.Caption = " 0.-"
        ActiDoblar ByVal 0
        CarDobDV
        cmd_jugar.Enabled = True
        DarVueltaCartas
        Deshop ByVal 0
    End If
    DesDobct 1
End Sub

Private Sub Cmenor_GotFocus()
MsgBE ByVal 0

End Sub

Private Sub Cmenor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 0
CmdMm Cmenor, 1
End Sub


Private Sub Doblar_Click(Index As Integer)
    DesDobct 0
    Select Case Index
    Case 0    'plantarse
        dinero.Caption = dinero + Val(AuxApuesta)
        CarDobDV
        EtImApDo.Caption = SinbMoneda & " 0.-"
        ActiDoblar ByVal 0
        ' DesHCtlDob 1
        cmd_jugar.Enabled = True
        DarVueltaCartas
        Deshop ByVal 0    'habilitar apuestas
    Case 1
        DarParaDoblar 1    'Doblar
        CartaCara ByVal CByte(6)

    End Select
    DesDobct 1
End Sub

Private Sub Doblar_GotFocus(Index As Integer)
    CmdMm Doblar(0), 1
End Sub

Private Sub Doblar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdMm Doblar(0), 1
    MsgBE ByVal 3

End Sub



Private Sub estado_Click(Index As Integer)
    Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
    Call carta_Click(Index)
End Sub

Private Sub estado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 11

End Sub

Private Sub EtAuBa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 2
MsgBE ByVal 5
Shape10.BackColor = 8454016
Case 1
  ShIdi.BackColor = &H80FF80
MsgBE ByVal 6
Case 0
 MsgBE ByVal 7
    ShSon.BackColor = &H80FF80
End Select
End Sub

Private Sub EtDes_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 EjPo ByVal Index
End Sub

Private Sub etProgmer_Click()
    ShellExecute Me.hwnd, vbNullString, "mailto:martsoft@gmail.com?subject=Mdls-->Game-->POKER", vbNullString, "C:\", SW_SHOWNORMAL    'Shell ("explore.exe mailto:martsoft@gmail.com")

End Sub

Private Sub etProgmer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AcercaDe ByVal True
    etProgmer.FontBold = True
End Sub
Private Sub Form_Activate()
    PonerMouseMano
End Sub
Private Function pauseMusica()
'for mp3 sounds
'MMControl.Command = "pause"
End Function
Private Function Musica()
'MMControl.Command = "close"
'MMControl.Command = "open"
'MMControl.Command = "play"
End Function
Private Function Archivo()
    If Len(App.Path) < 4 Then
        'MMControl.FileName = App.Path & "intro.mp3"
    Else
        'MMControl.FileName = "C:\Documents and Settings\Mar\Mis documentos\MP3\CASTELLANO\has echo\color.mp3" 'App.Path & "\intro.mp3"

    End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 49
        Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
         If cmd_jugar.Enabled Then Call carta_Click(0)
    Case 50
        Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
        If cmd_jugar.Enabled Then Call carta_Click(1)
    Case 51
        Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
       If cmd_jugar.Enabled Then Call carta_Click(2)       '
    Case 52
        Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
         If cmd_jugar.Enabled Then Call carta_Click(3)
    Case 53
        Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
      If cmd_jugar.Enabled Then Call carta_Click(4)
    Case 13
        Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
       If cmd_jugar.Enabled Then Call cmd_jugar_Click
    Case 112 'letra P
    If Picdob.Visible Then Call Doblar_Click(0)
    Case 45
    If Picdob.Visible Then Cmayor_Click
    Case 43
    If Picdob.Visible Then Cmenor_Click
    Case Else
        If SeInfc = 0 Then
            Son ByVal "ErrPret", ByVal "WAVSSISTEMAGRAL"
        Else
            Son ByVal "in_ErrPret", ByVal "WAVSSISTEMAGRAL"
        End If

    End Select
End Sub

Private Sub Form_Load()

    If GetSetting(App.EXEName, "Interface", "Lenguaje", "1") = "1" Then
        CargarInterface ByVal 1
    Else
        CargarInterface ByVal 0
    End If
    AcercaDe ByVal False
    '   Val (GetSetting(App.Title, "preferences", "backcolor", "&H00FFFFFF&"))
    If GetSetting(App.EXEName, "Interface", "Sonido", "1") = "1" Then
        Im_Son(0).Visible = False
        Im_Son(1).Visible = True
        SeSonido = 1
        Me.mnSonido.Checked = True
    Else
        Im_Son(0).Visible = True
        Im_Son(1).Visible = False
        Me.mnSonido.Checked = False
        SeSonido = 0
    End If
    'Ver La moneda elegida en el registro
    SeComNuevo = 1 'para que no salte el msgerror nuevo juego
    InvisibleEj ByVal 0 'invisible el ejemplo poker

    Dim Resp$: Resp = GetSetting(App.EXEName, "Interface", "Moneda", "1")
    If Resp = "1" Then    'Dollar
      mnDolar_Click
    ElseIf Resp = "2" Then    'Euro
        MnEuro_Click
    Else    'Pesos
        MnPesos_Click
    End If
    JuegoNuevo
   ActiDoblar ByVal 0
    DarVueltaCartas
    ColEstCarta
    dinero.Caption = 100
    saldo = 100
    njugadas = 1
    cambios = 0
    For i = 0 To 6
        carta(i) = LoadPicture(App.Path & "\FONDO\Pintas.GIF")
    Next i
    saldo = 100
    cmd_jugar.Enabled = True
End Sub
Private Function cargar(a As Integer, Optional CarDob As Byte) As String
    Dim aux As Integer
    cargar = ""
    Select Case a
    Case 0 To 12
        cargar = "c"
        aux = a + 1
    Case 13 To 25
        cargar = "d"
        aux = (a - 13) + 1
    Case 26 To 38
        cargar = "p"
        aux = (a - 26) + 1
    Case 39 To 51
        cargar = "t"
        aux = (a - 39) + 1
    End Select
    cargar = cargar + Trim(Str(aux)) + ".gif"
    If CarDob = 1 Then CartaADoblar(0) = CByte(aux)
    If CarDob = 2 Then CartaADoblar(1) = CByte(aux)

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnu_juego
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
InvisibleEj ByVal 0
MsgBE ByVal 1
Shape10.BackColor = &H7A618
    ShIdi.BackColor = &H7A618
    ShSon.BackColor = &H7A618
    CmdMm Doblar(0), 0
    CmdMm Cmenor, 0
    CmdMm Cmayor, 0
    CmdMm cmd_jugar, 0
    If SeAcerca = 0 Then AcercaDe ByVal False
    Acerca4.FontBold = False
    etProgmer.FontBold = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GetSetting(App.EXEName, "Interface", "Lenguaje", "1") = "1" Then
        SaveSetting App.EXEName, "Interface", "Lenguaje", "1"
    Else
        SaveSetting App.EXEName, "Interface", "Lenguaje", "0"
    End If
End Sub

Private Sub Im_Son_Click(Index As Integer)
    mnSonido_Click
End Sub

Private Sub Im_Son_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.PopupMenu mnidioma, , ShIdi.Left, ShIdi.Top + ShIdi.Height
End Sub

Private Sub Im_Son_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBE ByVal 7
    ShSon.BackColor = &H80FF80
End Sub

Private Sub Image4_Click()

End Sub

Private Sub In_Din_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.PopupMenu mnMoneda, , Shape10.Left, Shape10.Top + Shape10.Height
End Sub
Private Sub In_Din_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 5
Shape10.BackColor = 8454016
End Sub

Private Sub In_Int_Click(Index As Integer)
    If Index = 0 Then
        CargarInterface ByVal 1
    Else
        CargarInterface ByVal 0
    End If
End Sub

Private Sub In_Int_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.PopupMenu mnidioma, , ShIdi.Left, ShIdi.Top + ShIdi.Height


End Sub

Private Sub In_Int_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShIdi.BackColor = &H80FF80
MsgBE ByVal 6
End Sub

 

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 11
End Sub

Private Sub mnDolar_Click()
 
If SeComNuevo = 0 Then
   If msgIniNuevo = 0 Then Exit Sub
End If

    Dim o As Byte, i As Byte
    o = 1
    For i = 0 To 2
        EtApu(i).Caption = "u$s" & CStr(o)
        If o = 1 Then o = 0
        o = o + 5
    Next
    SinbMoneda = "u$s"
    EtDin.Caption = "u$s"
    EtDoSig.Caption = "u$s"
    mnDolar.Checked = True
    MnEuro.Checked = False
    MnPesos.Checked = False
    If SeInfc = 0 Then
        msNuC100.Caption = "Cargar " & SinbMoneda & "100."
    Else
        msNuC100.Caption = "Load " & SinbMoneda & "100."
    End If
    'Guarda en el registro la entrada seleccionada
    SaveSetting App.EXEName, "Interface", "Moneda", "1"
End Sub

Private Sub MnEuro_Click()
If SeComNuevo = 0 Then
   If msgIniNuevo = 0 Then Exit Sub
End If
Dim o As Byte, i As Byte
    o = 1
    For i = 0 To 2
        EtApu(i).Caption = "€" & CStr(o)
        If o = 1 Then o = 0
        o = o + 5
    Next
    SinbMoneda = "€"
    EtDin.Caption = "€"
    EtDoSig.Caption = "€"
    mnDolar.Checked = False
    MnEuro.Checked = True
    MnPesos.Checked = False
    If SeInfc = 0 Then
        msNuC100.Caption = "Cargar " & SinbMoneda & "100."
    Else
        msNuC100.Caption = "Load " & SinbMoneda & "100."
    End If
    'Guarda en el registro la entrada seleccionada
    SaveSetting App.EXEName, "Interface", "Moneda", "2"


End Sub

Private Sub mnIdEn_Click()
    CargarInterface ByVal 1
End Sub
Private Sub mnIdEsp_Click()
    CargarInterface ByVal 0
End Sub
Private Sub Label2_Click(Index As Integer)
    Son ByVal "selec", ByVal "WAVSSISTEMAGRAL"
    Call carta_Click(Index)

End Sub

Private Sub MnPesos_Click()
'si es un juego ya iniciado se pregunta si quiere reiniciar para cambiar la moneda
'
If SeComNuevo = 0 Then
   If msgIniNuevo = 0 Then Exit Sub
End If
Dim o As Byte, i As Byte
    o = 1
    For i = 0 To 2
        EtApu(i).Caption = "$" & CStr(o)
        If o = 1 Then o = 0
        o = o + 5
    Next
    SinbMoneda = "$"
    EtDin.Caption = "$"
    EtDoSig.Caption = "$"
    mnDolar.Checked = False
    MnEuro.Checked = False
    MnPesos.Checked = True
    If SeInfc = 0 Then
        msNuC100.Caption = "Cargar " & SinbMoneda & "100."
    Else
        msNuC100.Caption = "Load " & SinbMoneda & "100."
    End If
    'Guarda en el registro la entrada seleccionada
    SaveSetting App.EXEName, "Interface", "Moneda", "3"


End Sub

Private Sub mnSonido_Click()
    mnSonido.Checked = Not (mnSonido.Checked)
    If GetSetting(App.EXEName, "Interface", "Sonido", "1") = "0" Then
        Im_Son(0).Visible = False
        Im_Son(1).Visible = True
        SeSonido = 1
        SaveSetting App.EXEName, "Interface", "Sonido", "1"
    Else
        Im_Son(0).Visible = True
        Im_Son(1).Visible = False
        SeSonido = 0
        SaveSetting App.EXEName, "Interface", "Sonido", "0"
    End If
End Sub
Private Sub mnu_acerca_Click()
    
End Sub

Private Sub mnu_salir_Click()
    End
End Sub

Public Function barajar(vec() As Integer, tam As Integer) As Integer
'genera un vector con los numeros del 0 al 51
'y dependiendo del numero es corazon,diamantes,picas,treboles
'por ejemplo si es 10 seria la j de corazones
    Randomize
    Dim sw As Boolean
    Dim cont As Integer
    Dim X, Y As Integer
    sw = False
    While sw = False
        vec(0) = Int(Rnd * tam + 1)
        vec(1) = Int(Rnd * tam + 1)
        If vec(0) <> vec(1) Then
            sw = True
        End If
    Wend
    sw = False
    For X = 2 To tam
        sw = False
        While sw = False
            vec(X) = Int(Rnd * 52)
            cont = 0
            For Y = 0 To X - 1
                If vec(Y) <> vec(X) Then
                    cont = cont + 1
                End If
            Next Y
            If cont = X Then
                sw = True
            End If
        Wend
    Next X
End Function
Public Function unpar(vec() As Integer) As Byte
    Dim X, Y As Integer
    For X = 0 To 3
        Debug.Print vec(X) & "=" & vec(X + 1)

        If vec(X) = vec(X + 1) Then
            Y = Y + 1
        End If
    Next X
    If Y = 1 Then
        unpar = 1
    End If
End Function
Public Function dospares(vec() As Integer) As Byte
    Dim X, Y As Integer
    If (vec(0) = vec(1)) And (vec(2) = vec(3)) And (vec(3) <> vec(4)) And (vec(0) <> vec(3)) Then
        dospares = 2
    End If
    If (vec(1) = vec(2)) And (vec(3) = vec(4)) And (vec(0) <> vec(1)) And (vec(1) <> vec(4)) Then
        dospares = 2
    End If
    If (vec(0) = vec(1)) And (vec(3) = vec(4)) Then
        dospares = 2
    End If
End Function
Public Function trio(vec() As Integer) As Byte
    Dim X, Y As Integer
    For X = 0 To 2
        Select Case X
        Case 0
            If vec(X) = vec(X + 1) And vec(X) = vec(X + 2) And vec(X + 3) <> vec(X + 4) Then
                Y = Y + 1
            End If
        Case 1
            If vec(X) = vec(X + 1) And vec(X) = vec(X + 2) Then
                Y = Y + 1
            End If
        Case 2
            If vec(X) = vec(X + 1) And vec(X) = vec(X + 2) And vec(3 - X) <> vec(4 - X) Then
                Y = Y + 1
            End If
        End Select
    Next X
    If Y = 1 Then
        trio = 3
    End If
End Function
Public Function full(vec() As Integer) As Byte
    If vec(0) = vec(1) And (vec(2) = vec(3) And vec(2) = vec(4)) Then
        full = 5
    End If
    If vec(3) = vec(4) And (vec(0) = vec(1) And vec(0) = vec(2)) Then
        full = 5
    End If
End Function
Public Function cosuclimreal(vec() As String, vec1() As Integer) As Integer
'6 flush/ color
'7  straight flush/limpia
'8  Royal flush/Escalera real
'9 straight/ sucia
    Dim X, Y As Integer
    Dim a, b As Integer
    'Esto es para ver si se tiene una escalera
    For a = 1 To 4
        If vec1(a) = (vec1(a - 1) + 1) Then
            b = b + 1
        End If
    Next a
    'esto es para saber si todas son de la misma pinta
    For X = 0 To 3
        If vec(X) = vec(X + 1) Then
            Y = Y + 1
        End If
    Next X
    'si todas son iguales
    If Y = 4 Then
        'si no hay escalera y es color
        If b <> 4 Then
            cosuclimreal = 6
            '  straight  si hay escalera
        Else
            cosuclimreal = 7
        End If
        'if it's  Royal flush/si es una real!!!!
        If vec1(0) = 1 And vec1(1) = 10 And vec1(2) = 11 And vec1(3) = 12 And vec1(4) = 13 Then
            cosuclimreal = 8
        End If
    Else
        'it's straight/si es una sucia
        If b = 4 Then
            cosuclimreal = 9
        End If
    End If
End Function
Public Function poker(vec() As Integer) As Byte
    Dim X, Y As Integer
    For X = 0 To 1
        If vec(X) = vec(X + 1) And vec(X) = vec(X + 2) And vec(X) = vec(X + 3) Then
            Y = Y + 1
        End If
    Next X
    If Y = 1 Then
        poker = 4
    End If
End Function

Public Function convertiryord(vec() As Integer)
    Dim l, m, aux, aux1 As Integer
    For l = 0 To 4
        Select Case vec(l)
        Case 0 To 12
            vec(l) = vec(l) + 1
            pinta(l) = "c"
        Case 13 To 25
            vec(l) = (vec(l) - 13) + 1
            pinta(l) = "d"
        Case 26 To 38
            vec(l) = (vec(l) - 26) + 1
            pinta(l) = "p"
        Case 39 To 51
            vec(l) = (vec(l) - 39) + 1
            pinta(l) = "t"
        End Select
    Next l
    For l = 4 To 0 Step -1
        For m = 1 To l
            If vec(m - 1) > vec(m) Then
                aux = vec(m - 1)
                vec(m - 1) = vec(m)
                vec(m) = aux
            End If
            If Asc(pinta(m - 1)) > Asc(pinta(m)) Then
                aux1 = Asc(pinta(m - 1))
                pinta(m - 1) = pinta(m)
                pinta(m) = Chr(aux1)
            End If
        Next m
    Next l
End Function
Public Sub ColEstCarta()
    Dim i As Byte
    For i = 0 To 4
        ' estado(i).Caption = "Mantener"
        estado(i).BackColor = &H8000&
        'Shape3(i).BackStyle = 0
        Label2(i).BackColor = &H8000&
    Next
End Sub
Public Sub JuegoNuevo(Optional CanDinero As Long)
'   DarVueltaCartas
    ColEstCarta
  '  Debug.Print CanDinero
    If CanDinero <> 0 Then
        dinero.Caption = CStr(CanDinero)
    Else
        dinero.Caption = 100
    End If
    ' saldo = 100
    njugadas = 1
    DarVueltaCartas
SeComNuevo = 1
    'For i = 0 To 4
    
    '  carta(i) = LoadPicture(App.Path & "\FONDO\Pintas.GIF")
    ' Next i
    IIf Len(CanDinero) > 0, saldo = CInt(CanDinero), saldo = 100
ActiDoblar ByVal False
 cmd_jugar.Enabled = True
 
End Sub

Public Sub DarVueltaCartas()
'dar vuelta los naipes/Giving turn the cards 1-5
    Dim i As Byte
    For i = 0 To 4
        estado(i).Visible = False
        carta(i) = LoadPicture(App.Path & "\FONDO\Pintas.GIF")
        Shape3(i).BackColor = &HC0C0C0
        Label2(i).Visible = False
    Next
End Sub
Public Sub CartaCara(ByVal NroCarta As Byte)
    If NroCarta < 5 Then
        estado(NroCarta).Visible = True
        Label2(NroCarta).Visible = True
    End If
    Shape3(NroCarta).BackColor = &HFFFFFF
End Sub
Public Sub PonerMouseMano()
'Colocar diferentes cursores segun el control/Placing different cursors according to control
    Dim i As Byte
    For i = 0 To 4
         estado(i).MousePointer = 99
        Set estado(i).MouseIcon = LoadResPicture(101, vbResCursor)
        'Shape3(i).MousePointer = 99
        ' Set Shape3(i).MouseIcon = LoadResPicture(101, vbResCursor)
        carta(i).MousePointer = 99
        Set carta(i).MouseIcon = LoadResPicture(102, vbResCursor)
        Label2(i).MousePointer = 99
        Set Label2(i).MouseIcon = LoadResPicture(101, vbResCursor)
        cmd_jugar.MousePointer = 99
        Set cmd_jugar.MouseIcon = LoadResPicture(101, vbResCursor)
    Next
    For i = 0 To 1
        In_Int(i).MousePointer = 99
        Set In_Int(i).MouseIcon = LoadResPicture(101, vbResCursor)
        Im_Son(i).MousePointer = 99
        Set Im_Son(i).MouseIcon = LoadResPicture(101, vbResCursor)
      
    Next
   For i = 0 To 9
        EtDes(i).MousePointer = 99
      Set EtDes(i).MouseIcon = LoadResPicture(101, vbResCursor)
    Next

    In_Din(0).MousePointer = 99
    Set In_Din(0).MouseIcon = LoadResPicture(101, vbResCursor)

    Acerca6.MousePointer = 99
    Set Acerca6.MouseIcon = LoadResPicture(101, vbResCursor)
    etProgmer.MousePointer = 99
    Set etProgmer.MouseIcon = LoadResPicture(101, vbResCursor)
    Acerca4.MousePointer = 99
    Set Acerca4.MouseIcon = LoadResPicture(101, vbResCursor)
    Cmenor.MousePointer = 99
    Set Cmenor.MouseIcon = LoadResPicture(101, vbResCursor)
    Cmayor.MousePointer = 99
    Set Cmayor.MouseIcon = LoadResPicture(101, vbResCursor)
    Doblar(0).MousePointer = 99
    Set Doblar(0).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub
Public Sub MarcarTablaDesc(Idx As Byte, Optional SeDmT As Byte)
    Dim i As Byte
    If SeDmT = 1 Then
        For i = 0 To 9
            Me.EtDes(CInt(i)).ForeColor = vbWhite
            EtDes(CInt(i)).FontBold = True
        Next
    Else
        For i = 0 To 9
            Me.EtDes(i).ForeColor = vbWhite
        Next
        Me.EtDes(Idx).ForeColor = &H80FF80  ' vbRed
        IdcDes = Idx
        Timer1.Enabled = True
    End If
End Sub

Private Sub msNuC100_Click()
    cmd_jugar.Enabled = True
    JuegoNuevo
End Sub

Private Sub msNuCm_Click()
    Dim Resp$
    If SeInfc = 0 Then
        Resp = InputBox("Ingrese el la cantidad de " & SinbMoneda & " que desee." & Chr(13) & "Debe ser un numero entero entre 1 y 10000", "Cargar dinero")
        If Len(Trim(Resp)) = 0 Then Exit Sub

        If Not IsNumeric(Val(Resp)) Then Exit Sub
        If (Val(Resp)) < 10000 And (Val(Resp)) > 0 Then
         cmd_jugar.Enabled = True
        JuegoNuevo CLng(Resp)
        Else
        MsgBox "Error en el numero"
         End If
    Else
        Resp = InputBox("Enter the the quantity of " & SinbMoneda & Chr(13) & "Must Be an integer among 1 and 10000", "Cargar dinero")
        If Len(Trim(Resp)) = 0 Then Exit Sub

        If Not IsNumeric(Val(Resp)) Then Exit Sub
        If (Val(Resp)) < 10000 And (Val(Resp)) > 0 Then
          cmd_jugar.Enabled = True
        JuegoNuevo CLng(Resp)
     Else
        MsgBox "Error in the number"
      End If
   
    End If

End Sub

Private Sub opt_1_Click()
    EtApu(0).ForeColor = vbYellow
    EtApu(0).FontBold = True
    EtApu(1).ForeColor = vbWhite
    EtApu(1).FontBold = False
    EtApu(2).ForeColor = vbWhite
    EtApu(2).FontBold = False
End Sub
Private Sub opt_1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Son ByVal "click", ByVal "WAVSSISTEMAGRAL"
End Sub

Private Sub opt_1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 8

End Sub

Private Sub opt_10_Click()
    EtApu(0).ForeColor = vbWhite
    EtApu(0).FontBold = False
    EtApu(1).ForeColor = vbWhite
    EtApu(1).FontBold = False
    EtApu(2).ForeColor = vbYellow
    EtApu(2).FontBold = True
End Sub
Private Sub opt_10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Son ByVal "click", ByVal "WAVSSISTEMAGRAL"
End Sub

Private Sub opt_10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 10

End Sub

Private Sub opt_5_Click()
    EtApu(0).ForeColor = vbWhite
    EtApu(0).FontBold = False
    EtApu(1).ForeColor = vbYellow
    EtApu(1).FontBold = True
    EtApu(2).ForeColor = vbWhite
    EtApu(2).FontBold = False
End Sub
Private Sub opt_5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Son ByVal "click", ByVal "WAVSSISTEMAGRAL"
End Sub

Private Sub opt_5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBE ByVal 9

End Sub

Private Sub Picdob_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdMm Doblar(0), 0
    CmdMm Cmenor, 0
    CmdMm Cmayor, 0
    CmdMm cmd_jugar, 0
End Sub
Private Sub Timer1_Timer()

    If Me.EtDes(IdcDes).ForeColor = &H80FF80 Then   ' vbRed Then
        EtDes(IdcDes).ForeColor = vbWhite
        EtDes(IdcDes).BackColor = &H80FF80  ' vbRed
EtDes(IdcDes).FontBold = False
    Else
        EtDes(IdcDes).ForeColor = &H80FF80  ' vbRed
        EtDes(IdcDes).BackColor = vbWhite
    EtDes(IdcDes).FontBold = True

    End If
End Sub
Public Sub EfectoRepartir(Optional carta As Byte)
    Dim Izq As Long: Izq = Shape3(2).Left
    Const Arr As Long = 4080
    Dim i As Long    '4080
    Dim con As Byte
    Dim Idx As Byte
    For i = 1 To 10    'Step -1
        DoEvents
        Image1(1).ZOrder 1
        Image1(1).Top = Image1(1).Top - i
        Scar.ZOrder 0
        'Sleep 10'
    Next
    Select Case carta
    Case 1
        Me.Scar.Top = Arr
        Me.Scar.Left = Izq
        Me.Scar.Visible = True
        CarMov (1)
        con = 0
        Idx = 0
        For i = Me.Scar.Left To 390 Step -50
            'Do Until o = 1
            DoEvents
            con = con + 1
            If con = 13 Then
                Idx = Idx + 1: CarMov (Idx): con = 0
            End If
            DoEvents
            Me.Scar.Left = i
            Me.Scar.Top = i + 1200
        Next
        ' CarMov 0
        Me.Scar.Top = Me.carta(0).Top
        Me.Scar.Left = Me.carta(0).Left
        Me.Scar.Visible = False
    Case 2
        '''''
        Me.Scar.Top = Arr
        Me.Scar.Left = Izq
        Me.Scar.Visible = True
        CarMov (1)
        con = 0
        For i = Me.Scar.Left To 1710 Step -20
            DoEvents
            con = con + 1
            If con = 16 Then    'si se pasa en el bucle x veces se divide 1710 por X
                Idx = Idx + 1: CarMov (Idx): con = 0
            End If
            Me.Scar.Left = i
            Me.Scar.Top = i + 600
        Next
        Debug.Print con
        Me.Scar.Top = Me.carta(1).Top
        Me.Scar.Left = Me.carta(1).Left
        Me.Scar.Visible = False
    Case 3
        Me.Scar.Top = Arr
        Me.Scar.Left = Izq
        Me.Scar.Visible = True
        CarMov (1)
        con = 0
        Idx = 0
        For i = Me.Scar.Top To Me.carta(2).Top Step -10
            DoEvents
            'Me.Scar.Left = i
            Me.Scar.Top = i
            con = con + 1
            If con = 56 Then
                Idx = Idx + 1: CarMov (Idx): con = 0
            End If
        Next
        Me.Scar.Top = Me.carta(2).Top
        Me.Scar.Left = Me.carta(2).Left
        Me.Scar.Visible = False
    Case 4
        'carta cuatro
        Me.Scar.Top = Arr
        Me.Scar.Left = Izq
        Me.Scar.Visible = True
        CarMov (1)
        Dim Ct As Integer
        Ct = Arr
        con = 0
        For i = 1 To 1335 Step 10
            DoEvents
            Me.Scar.Left = 3000 + i
            Scar.Top = Ct
            Ct = Ct - 18
            con = con + 1
            If con = 33 Then
                Idx = Idx + 1: CarMov (Idx): con = 0
            End If
        Next
        Debug.Print con
        Me.Scar.Top = Me.carta(3).Top
        Me.Scar.Left = Me.carta(3).Left
        Me.Scar.Visible = False
    Case 5
        'carta cinco
        Me.Scar.Top = Arr
        Me.Scar.Left = Izq
        Me.Scar.Visible = True
        Ct = Arr
        con = 0
        CarMov (1)
        For i = 1 To 2670 Step 10
            con = con + 1: DoEvents
            Me.Scar.Left = 3000 + i
            Scar.Top = Ct    ' i + 100
            Ct = Ct - 9
            If con = 66 Then
                Idx = Idx + 1: CarMov (Idx): con = 0
            End If
        Next
        Debug.Print con
        Me.Scar.Top = Me.carta(4).Top
        Me.Scar.Left = Me.carta(4).Left
        Me.Scar.Visible = False
    End Select
    Image1(1).Top = 4150
End Sub

Public Sub CarMov(Ncar As Byte)
    Scar.Stretch = True
    Select Case Ncar
    Case 1
        Scar.Height = 900
        Scar.Width = 679
    Case 2
        Scar.Height = 1140
        Scar.Width = 915
    Case 3
        Scar.Height = 1260
        Scar.Width = 1035
    Case 4
        Scar.Height = 1380
        Scar.Width = 1125
    Case 0
        Scar.Stretch = False
    End Select
End Sub
Public Sub CargarInterface(ByVal SeInt As Byte)
' Cambiar idioma/Changing idiom
    Dim i As Byte
    Dim Idx As Integer
    If SeInt = 0 Then
        SeInfc = 0
        EtApu(3).Caption = "Apuesta:"
        For i = 0 To 4
            Label2(i).Caption = "Carta " & i
            estado(i).Caption = "Mantener"
        Next
        Idx = 30
        EtDes(1).Tag = "nada"
        For i = 0 To 9
            EtDes(i).Caption = LoadResString(Idx + i)
               EtDes(i).ToolTipText = "Multiplica por " & EtDes(i).Tag & " lo apostado"
        Next
        Im_Son(1).ToolTipText = "Sonido Activado"
        Im_Son(0).ToolTipText = "Sonido desactivado"
       
        msNuC100.Caption = "Cargar " & SinbMoneda & "100."
        msNuCm.Caption = "Ingresar dinero manualmente"
        Cmenor.ToolTipText = "Haga click aquí ó presione la tecla (-) si cree que la carta será menor"
        Cmayor.ToolTipText = "Haga click aquí ó presione la tecla (+), si cree que la carta será mayor"
        Doblar(0).ToolTipText = "Haga click aquí ó presione la tecla (P), no quiere doblar la apuesta"
        cmd_jugar.Caption = "Repartir"
        ErDinero.Caption = "Dinero:"
        Me.mnu_juego.Caption = "&Juego"
        Me.mnu_nuevo.Caption = "&Nuevo"
        mnidioma.Caption = "Idioma"
        mnIdEn.Caption = "Inglés"
        mnIdEsp.Caption = "Español"
        mnu_salir.Caption = "&Salir"
        mnSonido.Caption = "Sonido"
        Me.mnop.Caption = "Opciones"
        mnIdEsp.Checked = True
        mnIdEn.Checked = False
        'Bending bet/doblar apuestas
        Cmenor.Caption = "Menor"
        Cmayor.Caption = "Mayor"
        Doblar(0).Caption = "Plantarse"
        mnjr.Caption = "Jugar en Red por dinero"
        msjdemo.Caption = "Jugar Practicando localmente"
        mnMoneda.Caption = "Moneda"
        mnDolar.Caption = "Dolares (u$s)"
        MnEuro.Caption = "Euros (€)"
        MnPesos.Caption = "Pesos ($)"
        etProgmer.Caption = "Acerca del programador"
        In_Int(0).Visible = True    'imagenes interface
        In_Int(1).Visible = False
        Acerca2.Caption = "Programado por: Martin DLS"
        Acerca7.Caption = "Preguntas-Mas SW..."
        In_Din(0).ToolTipText = "Moneda"
       EtEjPo.Caption = "Ejemplo:"
        EtApDo.Caption = "Apuesta doblada:"
          
        SaveSetting App.EXEName, "Interface", "Lenguaje", "0"
    Else
        SeInfc = 1
        EtApu(3).Caption = "Bet: "
        For i = 0 To 4
            Label2(i).Caption = LoadResString(1 + SeInfc) & " " & i
            estado(i).Caption = LoadResString(124)
        Next
         EtDes(1).Tag = "Nothing"
       
        Idx = 40
        For i = 0 To 9
            EtDes(i).Caption = LoadResString(Idx + i)
             EtDes(i).ToolTipText = "Multiply for " & EtDes(i).Tag & " the bet"
         Next
         Cmenor.ToolTipText = "Make click here ó press the key (-) if you believes that the cart will be minor"
        Cmayor.ToolTipText = "Make click here ó press the key (+) if you believes that the cart will be major"
        Doblar(0).ToolTipText = "Make click here ó press the key (P), you do not want to bend the bet"
        EtEjPo.Caption = "Example:"
       msNuC100.Caption = "Load " & SinbMoneda & "100."
        msNuCm.Caption = "Entering manually"
      In_Din(0).ToolTipText = "Coin"
        Im_Son(0).ToolTipText = "Deactivated sound"
        Im_Son(1).ToolTipText = "Activated Sound"
        etProgmer.Caption = "About the programmer"
        cmd_jugar.Caption = "Deal"
        ErDinero.Caption = "Money:"
        Me.mnu_juego.Caption = "&Game"
        Me.mnu_nuevo.Caption = "&New"
        mnjr.Caption = "Play in Net for Money"
        msjdemo.Caption = "Playing Practicing locally"
        mnMoneda.Caption = "Coin"
        mnDolar.Caption = "Dollars (U$S)"
        MnEuro.Caption = "Euros (€)"
        MnPesos.Caption = "Pesos ($)"
        EtApDo.Caption = "Bets folded:"
        mnidioma.Caption = "&Language"
        mnIdEn.Caption = "&English"
        mnIdEsp.Caption = "&Spanish"
        Me.mnu_salir.Caption = "&Exit"
        mnSonido.Caption = "Sound"
        Me.mnop.Caption = "&Options"
        mnIdEsp.Checked = False
        mnIdEn.Checked = True
        'Bending bet/doblar apuestas
        Cmenor.Caption = "Minor"
        Cmayor.Caption = "Major"
        Doblar(0).Caption = "Stop"
        Acerca2.Caption = "Programmed for: Martin DLS"
        Acerca7.Caption = "Your questions-More SW..."
        In_Int(0).Visible = False    'imagenes interface
        In_Int(1).Visible = True
        SaveSetting App.EXEName, "Interface", "Lenguaje", "1"
    End If
End Sub
Public Sub Deshop(ByVal se As Byte)
    If se = 1 Then
        opt_1.Enabled = False
        opt_5.Enabled = False
        opt_10.Enabled = False
    Else
        opt_1.Enabled = True
        opt_5.Enabled = True
        opt_10.Enabled = True
    End If
End Sub

Public Sub AcercaDe(ByVal T_F As Boolean)
    Acerca0.Visible = T_F
    Acerca1.Visible = T_F
    Acerca2.Visible = T_F
 Acerca3.Visible = T_F
    Acerca4.Visible = T_F
    Acerca5.Visible = T_F
    Acerca6.Visible = T_F
    Acerca7.Visible = T_F
End Sub
Private Sub Timer2_Timer()
    
    If SeInfc = 0 Then
        Me.Caption = "M@rtsoft->Juegos->Poker  " & Format(Date, "Long Date") & " " & Time
    Else
        Me.Caption = "M@rtsoft->Games->Poker  " & Format(Date, "Long Date") & " " & Time
    End If
    If Acerca1.Visible = True Then
        SeAcerca = SeAcerca + 1
        If SeAcerca >= 7 Then AcercaDe ByVal False: SeAcerca = 0
    End If
   
    
  

End Sub
Public Sub ActiDoblar(ByVal T As Byte)
    Dim i As Byte

    Doblar(0).Enabled = IIf(T = 1, True, False)

    Cmenor.Enabled = IIf(T = 1, True, False)
    Cmayor.Enabled = IIf(T = 1, True, False)
   ' cmd_jugar.Enabled = IIf(T = 1, False, True)
    IIf T = 1, SeDoblando = 1, SeDoblando = 0
    Picdob.Visible = IIf(T = 1, True, False)
    For i = 0 To 4
        carta(i).Visible = IIf(T = 1, False, True)
        Label2(i).Visible = IIf(T = 1, False, True)
        estado(i).Visible = IIf(T = 1, False, True)
        Shape3(i).Visible = IIf(T = 1, False, True)
    Next
    cmd_jugar.Enabled = IIf(T = 1, False, False)
End Sub
Public Sub DarParaDoblar(Ncart As Byte)
' IIf Ncart = 5, Me.EfectoRepartir(0), Me.EfectoRepartir(3)
    If Ncart = 0 Then
        EfectoRepartir (0)
    Else
        Me.EfectoRepartir (4)
    End If
    a = barajar(Maso(), 51)
    CartasDmDo(Ncart) = Maso(51 - Ncart)
    CartaADoblar(Ncart) = CartasDmDo(Ncart)
    pcarta = ""
    pcarta = cargar(CByte(CartasDmDo(Ncart)), (Ncart + 1))
    carta(5 + Ncart) = LoadPicture(App.Path & "\CARTAS\" + pcarta)
End Sub

Public Function msgDoblarAp(dinero As Integer) As Byte
'--Inicio de Mensaje--

    Dim msbMensaje As String, msbOpcion As Integer, msbTitulo As String
    If SeInfc = 0 Then
        msbMensaje = "Ganaste " & SinbMoneda & CStr(dinero) + ".- " + vbCrLf
        msbMensaje = msbMensaje + "¿Desea doblar la apuesta?"
        msbTitulo = "Doblar Apuesta"
    Else
        msbMensaje = "You win!..." & SinbMoneda & CStr(dinero) + ".- " + vbCrLf
        msbMensaje = msbMensaje + "Do you Wish To bend the bet?"
        msbTitulo = "Bending the bet"
    End If
    msbOpcion = 4 + 32


    Select Case MsgBox(msbMensaje, msbOpcion, msbTitulo)
    Case 6  'Si/yes
        msgDoblarAp = 1
    Case 7  'No
        msgDoblarAp = 0
    End Select
    '--Fin de Mensaje--
End Function
Public Function msgApDoblada(DinDob As Integer) As Byte
'--Inicio de Mensaje--
    Dim msbMensaje As String, msbOpcion As Integer, msbTitulo As String
    If SeInfc = 0 Then
        msbTitulo = "Doblar la apuesta"
        msbMensaje = "Acertaste!!! ganas.. " & SinbMoneda & CStr(DinDob / 2) + vbCrLf
        msbMensaje = msbMensaje + "Subtotal Doblado: " & SinbMoneda & CStr(DinDob) & vbCrLf
        msbMensaje = msbMensaje + "Total acumulado: " & SinbMoneda & CStr(dinero.Caption) + vbCrLf
        msbMensaje = msbMensaje + "Total acumulado más Subtotoal Doblado: " & SinbMoneda & CStr(Val(dinero.Caption) + DinDob) + vbCrLf
        msbMensaje = msbMensaje + "¿Desea doblar la apuesta?" + vbCrLf
        msbMensaje = msbMensaje + "Presione Sí para Doblar ó No para plantarse"
    Else
        msbMensaje = "Yes, that's right! You win ..." & SinbMoneda & CStr(DinDob / 2) + vbCrLf
        msbMensaje = msbMensaje + "Bet total:" & SinbMoneda & CStr(DinDob) & vbCrLf
        msbMensaje = msbMensaje + "Accumulated total: " & SinbMoneda & CStr(dinero.Caption) + vbCrLf
        msbMensaje = msbMensaje + "Accumulated total more the bet total: " & SinbMoneda & CStr(Val(dinero.Caption) + DinDob) + vbCrLf
        msbMensaje = msbMensaje + "Do you Wish To bend the bet?" + vbCrLf
        msbMensaje = msbMensaje + "Press Yes (to Bend) or No (to stop)"
        msbTitulo = "Bending the bet"
    End If
    msbOpcion = 4 + 32
    Select Case MsgBox(msbMensaje, msbOpcion, msbTitulo)
    Case 6  'Si/yes
        msgApDoblada = 1
    Case 7  'No
        msgApDoblada = 0
    End Select
    '--Fin de Mensaje--
End Function

Public Sub CarDobDV()
    Dim i As Byte
    For i = 5 To 6
        carta(i) = LoadPicture(App.Path & "\FONDO\Pintas.GIF")
        Shape3(i).BackColor = &HC0C0C0
    Next i
End Sub

Public Sub DarVueltaCarta(Nc As Byte)
    carta(Nc) = LoadPicture(App.Path & "\FONDO\Pintas.GIF")
    Shape3(Nc).BackColor = &HC0C0C0
End Sub
Public Sub DesHCtlDob(Hab As Byte)
    Exit Sub
    Dim i As Byte
    For i = 1 To 4
        carta(i).Enabled = IIf(Hab = 1, True, False)
        Label2(i).Enabled = IIf(Hab = 1, True, False)
        estado(i).Enabled = IIf(Hab = 1, True, False)
    Next
    cmd_jugar.Enabled = IIf(Hab = 1, True, False)
End Sub
Public Sub DesDobct(S As Byte)
    Cmenor.Enabled = IIf(S = 1, True, False)
    Cmayor.Enabled = IIf(S = 1, True, False)
    Doblar(0).Enabled = IIf(S = 1, True, False)
End Sub
Public Sub VerificarCartas()
    If opt_1 = True Then
        apuesta = 1
    ElseIf opt_5 = True Then
        apuesta = 5
    Else: apuesta = 10
    End If
    a = convertiryord(CartasDm())
    resultado = unpar(CartasDm())
    If resultado = 1 Then
        saldo = saldo
        MarcarTablaDesc 1
        nada = False
    End If
    resultado = dospares(CartasDm())
    If resultado = 2 Then
        AuxApuesta = apuesta
        saldo = saldo + apuesta
        MarcarTablaDesc 2
        dinero.Caption = Str(saldo)
        nada = False
    End If
    resultado = trio(CartasDm())
    If resultado = 3 Then
        AuxApuesta = (apuesta * 2)
        saldo = saldo + (apuesta * 2)
        MarcarTablaDesc 3
        ' MsgBox "Tienes Pierna"'you have tree of king
        If SeInfc = 0 Then
            ' Son ByVal "hepierna", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "In_hepierna", ByVal "WAVSSISTEMAGRAL"
        End If
        dinero.Caption = Str(saldo)
        nada = False
    End If
    resultado = poker(CartasDm())
    If resultado = 4 Then
        AuxApuesta = (apuesta * 6)
        saldo = saldo + (apuesta * 6)
        MarcarTablaDesc 7
        ' MsgBox "Tienes un poker"'four of the king
        If SeInfc = 0 Then
            ' Son ByVal "hepoker", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "In_hepoker", ByVal "WAVSSISTEMAGRAL"

        End If
        dinero.Caption = Str(saldo)
        nada = False
    End If
    resultado = full(CartasDm())
    If resultado = 5 Then
        AuxApuesta = (apuesta * 5)
        saldo = saldo + (apuesta * 5)
        MarcarTablaDesc 6

        If SeInfc = 0 Then
            ' Son ByVal "hefull", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "In_hefull", ByVal "WAVSSISTEMAGRAL"
        End If

        ' MsgBox "Tienes un full"
        dinero.Caption = Str(saldo)
        nada = False
    End If
    resultado = cosuclimreal(pinta(), CartasDm())
    If resultado = 6 Then
        AuxApuesta = (apuesta * 4)
        saldo = saldo + (apuesta * 4)

        If SeInfc = 0 Then
            ' Son ByVal "hecolor", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "In_hecolor", ByVal "WAVSSISTEMAGRAL"
        End If

        MarcarTablaDesc 5
        'MsgBox "Tienes Color"
        dinero.Caption = Str(saldo)
        nada = False
    End If
    If resultado = 7 Then
        AuxApuesta = (apuesta * 7)
        saldo = saldo + (apuesta * 7)
        If SeInfc = 0 Then
            ' Son ByVal "HEESCLIM", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "In_HEESCLIM", ByVal "WAVSSISTEMAGRAL"
        End If
        MarcarTablaDesc 7
        ' MsgBox "Tienes Limpia"'you have Straight Flush!
        dinero.Caption = Str(saldo)
        nada = False
    End If
    If resultado = 8 Then
        AuxApuesta = (apuesta * 8)
        saldo = saldo + (apuesta * 8)
        MarcarTablaDesc 9
        If SeInfc = 0 Then
            ' Son ByVal "heescreal", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "In_heescreal", ByVal "WAVSSISTEMAGRAL"
        End If
        ' MsgBox "Tienes Real"'Royal Flush
        dinero.Caption = Str(saldo)
        nada = False
    End If
    If resultado = 9 Then
        AuxApuesta = (apuesta * 3)
        saldo = saldo + (apuesta * 3)
        MarcarTablaDesc 4
        If SeInfc = 0 Then
            ' Son ByVal "heescsuc", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "In_heescsuc", ByVal "WAVSSISTEMAGRAL"
        End If
        '  MsgBox "Tienes sucia"'Straight
        nada = False
    End If
    If nada <> True Then
        If AuxApuesta <> 0 Then
            If msgDoblarAp((AuxApuesta)) = 1 Then    'doblo la apuesta
                '   DesHCtlDob 0
                ActiDoblar ByVal 1
                DarParaDoblar 0
                CartaCara ByVal CByte(5)
                EtImApDo.Caption = CStr(AuxApuesta)
                'SeDoblar = 1    ' DarParaDoblar 1 '    Image2.ZOrder 1
                cmd_jugar.Enabled = False    'CartaCara ByVal CByte(6)
            Else
                dinero.Caption = Str(saldo)
            End If
        End If
        'Bending the bet
        'pregunta si dobla la apuesta

    Else
        AuxApuesta = (apuesta * 2)
        saldo = saldo - (apuesta * 2)
        MarcarTablaDesc 0

        If SeInfc = 0 Then
            ' Son ByVal "henada", ByVal "WAVSSISTEMAGRAL"
        Else
            ' Son ByVal "in_henada", ByVal "WAVSSISTEMAGRAL"
        End If
        ' MsgBox "No tienes nada"'you haven't nothing!
        dinero.Caption = Str(saldo)
    End If
    If SeInfc = 0 Then
        cmd_jugar.Caption = LoadResString(5)
    Else
        cmd_jugar.Caption = LoadResString(6)
    End If
    cambios = 0
    For i = 0 To 4
        If SeInfc = 0 Then
            estado(i).Caption = LoadResString(123)
        Else
            estado(i).Caption = LoadResString(124)
        End If
    Next i
    ColEstCarta
    'DarVueltaCartas

End Sub
'efect bold in  mouse_move Event--> (CommandButton)
Public Sub CmdMm(C As Object, se As Byte)
    With C
        .FontBold = IIf(se = 1, True, False)
    End With
End Sub

Public Function msgIniNuevo() As Byte

Dim msbMensaje As String, msbOpcion As Integer, msbTitulo As String
'--Fin de Mensaje--


      If SeInfc = 0 Then
msbMensaje = "Debes iniciar un juego nuevo para poder cambiar la moneda" + vbCrLf
msbMensaje = msbMensaje + "¿Deseas hacerlo?"
msbOpcion = 4 + 48
msbTitulo = " Cambio de moneda"
Select Case MsgBox(msbMensaje, msbOpcion, msbTitulo)
    Case 6  'Si
JuegoNuevo
msgIniNuevo = 1
    Case 7  'No
Exit Function
End Select


Else
msbMensaje = "You must initiate a new game to be able to change the coin" + vbCrLf
msbMensaje = msbMensaje + "Do you wish to do it?"
msbOpcion = 4 + 48
msbTitulo = "Change of coin"

Select Case MsgBox(msbMensaje, msbOpcion, msbTitulo)
    Case 6  'Si
JuegoNuevo
msgIniNuevo = 1
Case 7  'No
Exit Function
End Select

End If

End Function
'mensajes para barra de estado
Public Sub MsgBE(ByVal m As Byte)
Select Case m
Case 0 'doblar menor
    If SeInfc = 0 Then
    EtBe.Caption = "Presione la tecla '-' si usa el teclado."
    Else
      EtBe.Caption = "Press the key '-' if you make use of the keyboard."
    End If
    
Case 1
EtBe.Caption = ""
Case 2 'boblar mayor
   If SeInfc = 0 Then
    EtBe.Caption = "Presione la tecla '+' si usa el teclado."
    Else
     EtBe.Caption = "Press the key '+' if you make use of the keyboard."
    End If
Case 3 'parar doblar
 If SeInfc = 0 Then
    EtBe.Caption = "Presione la tecla 'P' si usa el teclado."
    Else
     EtBe.Caption = "Press the key 'P' if you make use of the keyboard."
    End If
Case 4
 If SeInfc = 0 Then
  If Trim(cmd_jugar.Caption) = "Repartir" Then
      EtBe.Caption = "Reparte las cartas" & "" & "presione la tecla 'Entrar' ó 'barra espaciadora', si usa el teclado."
  Else
      EtBe.Caption = "Renuava las cartas seleccionadas si es que seleccionó alguna." & "" & "presione la tecla 'Entrar' ó 'barra espaciadora', si usa el teclado."
  End If
  
    Else
   
   If Trim(cmd_jugar.Caption) = "Deal" Then
      EtBe.Caption = "Distribute the cards" & "" & "Press the key 'Intro' or 'space bar', if you make use of the keyboard."
  Else
      EtBe.Caption = "Renew the selected cards if you selected any one." & "" & "Press the key 'Intro' or 'space bar', if you make use of the keyboard."
  End If
  End If
  Case 5 'barra de herramientas cambio de moneda
    If SeInfc = 0 Then
       EtBe.Caption = "Cambiar tipo de moneda, solo se permite el cambio cuando no se juega en red por dinero y cuando es un juego nuevo "
    Else
    EtBe.Caption = "Changing type of coin, only she affords the change if not one plays in net for money and when is a new game"
    End If
  Case 6 'Cambio de idioma/change language
    
    If SeInfc = 0 Then
       EtBe.Caption = "Cambiar interfaz Ingles/Español"
    Else
    EtBe.Caption = "Changing interface Spanish/English"
    End If
  Case 7
    If SeInfc = 0 Then
       EtBe.Caption = "Desactivar/Activar todos los sonidos"
    Else
    EtBe.Caption = "Deactivating/activating all of the sounds"
    End If
 Case 8
   If SeInfc = 0 Then
       EtBe.Caption = "Apuesta " & EtApu(0)
    Else
    EtBe.Caption = "The " & EtApu(0) & " bets "
    End If

 Case 9
   If SeInfc = 0 Then
       EtBe.Caption = "Apuesta " & EtApu(1)
    Else
    EtBe.Caption = "The " & EtApu(1) & " bets "
    End If

  Case 10
    If SeInfc = 0 Then
       EtBe.Caption = "Apuesta " & EtApu(2)
    Else
    EtBe.Caption = "Apuesta " & EtApu(2)
    End If
Case 11
   If SeInfc = 0 Then
       EtBe.Caption = "Seleccione las cartas que desee cambiar. Presione las teclas del 1 al 5 si utiliza el teclado"
    Else
    EtBe.Caption = "Make click to select the cards that she wishes to change ....Press the key 1 to 5 if utilizes the keyboard "
    End If
Case 12

End Select
End Sub

Private Sub Timer3_Timer()
'efecto rebote marquesina/Effect rebound maquee
 Static S As Byte
 Static a As Byte
If a = 8 Then
Acerca0.Caption = Right(Acerca0.Caption, Len(Acerca0.Caption) - 1)
 S = S + 1
 If S = 8 Then S = 0: a = 0
Else
 Acerca0.Caption = " " & Acerca0.Caption
  a = a + 1
End If
End Sub

Public Sub EjPo(ByVal Ej As Byte)
InvisibleEj ByVal 1

Select Case Ej
Case 9
'Escalera real/Royal Flush
Set ImEjC1.Picture = Me.ImaCor.Picture
Set ImEjC2.Picture = Me.ImaCor.Picture
Set ImEjC3.Picture = Me.ImaCor.Picture
Set ImEjC4.Picture = Me.ImaCor.Picture
Set ImEjC5.Picture = Me.ImaCor.Picture

EtEjP(0).Caption = "10"
EtEjP(1).Caption = "J"
EtEjP(2).Caption = "Q"
EtEjP(3).Caption = "K"
EtEjP(4).Caption = "A"

Case 8
'Escalera limpia/Straight Flush
Set ImEjC1.Picture = Me.ImTre.Picture
Set ImEjC2.Picture = Me.ImTre.Picture
Set ImEjC3.Picture = Me.ImTre.Picture
Set ImEjC4.Picture = Me.ImTre.Picture
Set ImEjC5.Picture = Me.ImTre.Picture

EtEjP(0).Caption = "4"
EtEjP(1).Caption = "5"
EtEjP(2).Caption = "6"
EtEjP(3).Caption = "7"
EtEjP(4).Caption = "8"

Case 7
'Poker/Four of a Kind
Set ImEjC1.Picture = Me.Impic.Picture
Set ImEjC2.Picture = Me.Impic.Picture
Set ImEjC3.Picture = Me.Impic.Picture
Set ImEjC4.Picture = Me.Impic.Picture
Set ImEjC5.Picture = Me.ImRom.Picture
EtEjP(0).Caption = "8"
EtEjP(1).Caption = "8"
EtEjP(2).Caption = "8"
EtEjP(3).Caption = "8"
EtEjP(4).Caption = "J"
Case 6
'full/full House
Set ImEjC1.Picture = Me.ImRom.Picture
Set ImEjC2.Picture = Me.ImTre.Picture
Set ImEjC3.Picture = Me.ImaCor.Picture
Set ImEjC4.Picture = Me.ImTre.Picture
Set ImEjC5.Picture = Me.Impic.Picture
EtEjP(0).Caption = "2"
EtEjP(1).Caption = "2"
EtEjP(2).Caption = "2"
EtEjP(3).Caption = "5"
EtEjP(4).Caption = "5"
Case 5
'Flush/flush
Set ImEjC1.Picture = Me.ImRom.Picture
Set ImEjC2.Picture = Me.ImRom.Picture
Set ImEjC3.Picture = Me.ImRom.Picture
Set ImEjC4.Picture = Me.ImRom.Picture
Set ImEjC5.Picture = Me.ImRom.Picture
EtEjP(0).Caption = "5"
EtEjP(1).Caption = "7"
EtEjP(2).Caption = "2"
EtEjP(3).Caption = "K"
EtEjP(4).Caption = "10"
Case 4
'Escalera sucia/Straight
Set ImEjC1.Picture = Me.ImRom.Picture
Set ImEjC2.Picture = Me.ImaCor.Picture
Set ImEjC3.Picture = Me.ImTre.Picture
Set ImEjC4.Picture = Me.ImRom.Picture
Set ImEjC5.Picture = Me.Impic.Picture
EtEjP(0).Caption = "3"
EtEjP(1).Caption = "4"
EtEjP(2).Caption = "5"
EtEjP(3).Caption = "6"
EtEjP(4).Caption = "7"
Case 3
'trio/Three of a Kind
Set ImEjC1.Picture = Me.ImaCor.Picture
Set ImEjC2.Picture = Me.ImaCor.Picture
Set ImEjC3.Picture = Me.ImaCor.Picture
Set ImEjC4.Picture = Me.ImRom.Picture
Set ImEjC5.Picture = Me.Impic.Picture
EtEjP(0).Caption = "7"
EtEjP(1).Caption = "7"
EtEjP(2).Caption = "7"
EtEjP(3).Caption = "2"
EtEjP(4).Caption = "9"
Case 2
'par doble/Two Pair
Set ImEjC1.Picture = Me.Impic.Picture
Set ImEjC2.Picture = Me.Impic.Picture
Set ImEjC3.Picture = Me.ImaCor.Picture
Set ImEjC4.Picture = Me.ImRom.Picture
Set ImEjC5.Picture = Me.ImRom.Picture

EtEjP(0).Caption = "8"
EtEjP(1).Caption = "8"
EtEjP(2).Caption = "2"
EtEjP(3).Caption = "5"
EtEjP(4).Caption = "5"
Case 1
'par simple/one pair
Set ImEjC1.Picture = Me.ImaCor.Picture
Set ImEjC2.Picture = Me.ImaCor.Picture
Set ImEjC3.Picture = Me.Impic.Picture
Set ImEjC4.Picture = Me.ImRom.Picture
Set ImEjC5.Picture = Me.Impic.Picture
EtEjP(0).Caption = "8"
EtEjP(1).Caption = "8"
EtEjP(2).Caption = "A"
EtEjP(3).Caption = "K"
EtEjP(4).Caption = "5"
Case 0
Set ImEjC1.Picture = Me.ImaCor.Picture
Set ImEjC2.Picture = Me.ImaCor.Picture
Set ImEjC3.Picture = Me.Impic.Picture
Set ImEjC4.Picture = Me.ImRom.Picture
Set ImEjC5.Picture = Me.ImTre.Picture

EtEjP(0).Caption = "10"
EtEjP(1).Caption = "7"
EtEjP(2).Caption = "K"
EtEjP(3).Caption = "J"
EtEjP(4).Caption = "2"
End Select

End Sub

Public Sub InvisibleEj(ByVal S As Byte)
    Dim i As Byte
    For i = 0 To 4
        EtEjP(i).Visible = IIf(S = 1, True, False)
    Next
    If SeInfc = 0 Then
        If S = 1 Then
            With EtEjPo
                .Caption = "Ejemplo:"
              '  .FontSize = 9
              '  .FontBold = True
            End With
        ElseIf S = 0 Then
            With EtEjPo
                .Caption = Space(43) & "Tabla de puntuaciones"
             '   .FontSize = 9
             '   .FontBold = True
            End With
       End If
        'EtEjPo.Caption = IIf(S = 1, "Ejemplo:", "                                      Tabla de puntuaciones")
    Else
      
  If S = 1 Then
            With EtEjPo
                .Caption = "Example:"
            '    .FontSize = 9
            '    .FontBold = False
            End With
        ElseIf S = 0 Then
            With EtEjPo
                .Caption = Space(48) & "List of punctuations"
              '  .FontSize = 9
             '   .FontBold = True
            End With
       End If
    
    End If

    ImEjC1.Visible = IIf(S = 1, True, False)
    ImEjC2.Visible = IIf(S = 1, True, False)
    ImEjC3.Visible = IIf(S = 1, True, False)
    ImEjC4.Visible = IIf(S = 1, True, False)
    ImEjC5.Visible = IIf(S = 1, True, False)
End Sub
