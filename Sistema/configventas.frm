VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgConfigVentas 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreos de Atributos de Cotiza"
   ClientHeight    =   9555
   ClientLeft      =   225
   ClientTop       =   390
   ClientWidth     =   11535
   LinkTopic       =   "Form2"
   ScaleHeight     =   9555
   ScaleWidth      =   11535
   Begin VB.TextBox Clave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   70
      Text            =   " "
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox ClaveII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   68
      Text            =   " "
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Operador 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Salida 
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   11668
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "configventas.frx":0000
      Tab(0).ControlCount=   8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Titulo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Titulo2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Titulo3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Titulo4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Opcion1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Opcion2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Opcion3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Opcion4"
      Tab(0).Control(7).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "configventas.frx":001C
      Tab(1).ControlCount=   12
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Titulo11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Titulo12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Titulo13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Titulo14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Titulo16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Titulo15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Opcion11"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "Opcion12"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "Opcion13"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Opcion14"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Opcion16"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Opcion15"
      Tab(1).Control(11).Enabled=   -1  'True
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "configventas.frx":0038
      Tab(2).ControlCount=   38
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Titulo24"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Titulo23"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Titulo22"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Titulo21"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Titulo28"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Titulo27"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Titulo26"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Titulo25"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Titulo212"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Titulo211"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Titulo210"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Titulo29"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Titulo217"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Titulo218"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Titulo219"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Titulo213"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Titulo214"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Titulo215"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Titulo216"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Opcion24"
      Tab(2).Control(19).Enabled=   -1  'True
      Tab(2).Control(20)=   "Opcion23"
      Tab(2).Control(20).Enabled=   -1  'True
      Tab(2).Control(21)=   "Opcion22"
      Tab(2).Control(21).Enabled=   -1  'True
      Tab(2).Control(22)=   "opcion21"
      Tab(2).Control(22).Enabled=   -1  'True
      Tab(2).Control(23)=   "Opcion28"
      Tab(2).Control(23).Enabled=   -1  'True
      Tab(2).Control(24)=   "Opcion27"
      Tab(2).Control(24).Enabled=   -1  'True
      Tab(2).Control(25)=   "Opcion26"
      Tab(2).Control(25).Enabled=   -1  'True
      Tab(2).Control(26)=   "Opcion25"
      Tab(2).Control(26).Enabled=   -1  'True
      Tab(2).Control(27)=   "Opcion212"
      Tab(2).Control(27).Enabled=   -1  'True
      Tab(2).Control(28)=   "Opcion211"
      Tab(2).Control(28).Enabled=   -1  'True
      Tab(2).Control(29)=   "Opcion210"
      Tab(2).Control(29).Enabled=   -1  'True
      Tab(2).Control(30)=   "Opcion29"
      Tab(2).Control(30).Enabled=   -1  'True
      Tab(2).Control(31)=   "Opcion217"
      Tab(2).Control(31).Enabled=   -1  'True
      Tab(2).Control(32)=   "Opcion218"
      Tab(2).Control(32).Enabled=   -1  'True
      Tab(2).Control(33)=   "Opcion219"
      Tab(2).Control(33).Enabled=   -1  'True
      Tab(2).Control(34)=   "Opcion213"
      Tab(2).Control(34).Enabled=   -1  'True
      Tab(2).Control(35)=   "Opcion214"
      Tab(2).Control(35).Enabled=   -1  'True
      Tab(2).Control(36)=   "Opcion215"
      Tab(2).Control(36).Enabled=   -1  'True
      Tab(2).Control(37)=   "Opcion216"
      Tab(2).Control(37).Enabled=   -1  'True
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "configventas.frx":0054
      Tab(3).ControlCount=   4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Titulo31"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Titulo32"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Opcion31"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "Opcion32"
      Tab(3).Control(3).Enabled=   -1  'True
      Begin VB.CheckBox Opcion32 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69960
         TabIndex        =   66
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Opcion31 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69960
         TabIndex        =   64
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Opcion216 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   56
         Top             =   6000
         Width           =   615
      End
      Begin VB.CheckBox Opcion215 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   55
         Top             =   5700
         Width           =   615
      End
      Begin VB.CheckBox Opcion214 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   54
         Top             =   5280
         Width           =   615
      End
      Begin VB.CheckBox Opcion213 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   53
         Top             =   4920
         Width           =   615
      End
      Begin VB.CheckBox Opcion219 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64800
         TabIndex        =   52
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox Opcion218 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64800
         TabIndex        =   51
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Opcion217 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64800
         TabIndex        =   50
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Opcion29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   45
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox Opcion210 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   44
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox Opcion211 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   43
         Top             =   4200
         Width           =   615
      End
      Begin VB.CheckBox Opcion212 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   42
         Top             =   4560
         Width           =   615
      End
      Begin VB.CheckBox Opcion25 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   37
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox Opcion26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   36
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox Opcion27 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   35
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox Opcion28 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   34
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox opcion21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   29
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Opcion22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Opcion23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   27
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox Opcion24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   26
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox Opcion15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   23
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Opcion16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   22
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Opcion4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox Opcion3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   1380
         Width           =   615
      End
      Begin VB.CheckBox Opcion2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   17
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox Opcion1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   660
         Width           =   615
      End
      Begin VB.CheckBox Opcion14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Opcion11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Titulo32 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   67
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Titulo31 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   65
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Titulo216 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   63
         Top             =   6000
         Width           =   4335
      End
      Begin VB.Label Titulo215 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   62
         Top             =   5640
         Width           =   4335
      End
      Begin VB.Label Titulo214 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   61
         Top             =   5280
         Width           =   4335
      End
      Begin VB.Label Titulo213 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   60
         Top             =   4920
         Width           =   4335
      End
      Begin VB.Label Titulo219 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   59
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Titulo218 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   58
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Titulo217 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   57
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Titulo29 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   49
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label Titulo210 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   48
         Top             =   3840
         Width           =   4095
      End
      Begin VB.Label Titulo211 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   47
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label Titulo212 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   46
         Top             =   4560
         Width           =   3975
      End
      Begin VB.Label Titulo25 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   41
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label Titulo26 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   40
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Titulo27 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   39
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Label Titulo28 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   38
         Top             =   3120
         Width           =   4095
      End
      Begin VB.Label Titulo21 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   33
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Titulo22 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   32
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Titulo23 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   31
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label Titulo24 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   30
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Titulo15 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   25
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Titulo16 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   24
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label Titulo4 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1740
         Width           =   3975
      End
      Begin VB.Label Titulo3 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1380
         Width           =   3855
      End
      Begin VB.Label Titulo2 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Titulo1 
         Caption         =   "Ingreso de Bancos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Titulo14 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   12
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Titulo13 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   10
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Titulo12 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   8
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Titulo11 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   6
         Top             =   720
         Width           =   3975
      End
   End
   Begin VB.Label lblLabels 
      Caption         =   "Clave Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   71
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      Caption         =   "Clave Presupuesto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   69
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgConfigVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Load()

    Operador.Text = ""
    DesOperador.Caption = ""
    
    Tablas.TabCaption(0) = "Clientes"
    Tablas.TabCaption(1) = "Novedades"
    Tablas.TabCaption(2) = "Listados"
    Tablas.TabCaption(3) = "Procesos"
    
    Rem titulo1
    
    Titulo1.Caption = "Ingreso de Cuentas Contables"
    Titulo2.Caption = "Ingreso de Proveedores"
    Titulo3.Caption = "Ingreso de Conceptos de Compras"
    Titulo4.Caption = "Ingeso de Bancos"

    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    
    
    Rem titulo2
    
    Titulo11.Caption = "Ingreso de Comprobantes de Proveedores"
    Titulo12.Caption = "Ingreso de Pagos"
    Titulo13.Caption = "Ingreso de Cobranzas"
    Titulo14.Caption = "Ingreso de Depositos"
    Titulo15.Caption = "Ingreso de Depositos"
    Titulo16.Caption = "Ingresos de Gastos"
    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0

    
    Rem titulo3
    
    Titulo21.Caption = "Listado de Cuenta Corrientes de Proveedoes"
    Titulo22.Caption = "Consulta de Cuenta Corriente de Proveedores"
    Titulo23.Caption = "Proyeccion de Pagos"
    Titulo24.Caption = "Listado de Cuenta Corriente de Proveedores a Fecha"
    Titulo25.Caption = "Listado de Pagos por Fecha"
    Titulo26.Caption = "Listado de Pagos por Proveedor"
    Titulo27.Caption = "Listado de Pagos por Proveedor"
    Titulo28.Caption = "Listado de Cobranzas"
    Titulo29.Caption = "Subdiario de Caja"
    Titulo210.Caption = "Listado de Movimientos Bancarios"
    Titulo211.Caption = "Listado de Cheques en Cartera"
    Titulo212.Caption = "Listado de Cheques en Cartera por Cliente"
    Titulo213.Caption = "Listado de Cheques en Cartera por Nro de Cheque"
    Titulo214.Caption = "Listado de Cheques en Cartera por Proveedor    "
    Titulo215.Caption = "Listado de Cheques en Cartera por Fecha de Ingreso"
    Titulo216.Caption = "Iva Compras"
    Titulo217.Caption = "Listado de Compras por concepto"
    Titulo218.Caption = "Listado de Imputaciones Contables"
    Titulo219.Caption = "Listado de Caja"

    opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion210.Value = 0
    Opcion211.Value = 0
    Opcion212.Value = 0
    Opcion213.Value = 0
    Opcion214.Value = 0
    Opcion215.Value = 0
    Opcion216.Value = 0
    Opcion217.Value = 0
    Opcion218.Value = 0
    Opcion219.Value = 0
    
    
    
    Rem titulo4
    
    Titulo31.Caption = "Ingreso de Configuracion del Sistema"
    Titulo32.Caption = "Ingreso de Datos de la Empresa"

    Opcion31.Value = 0
    Opcion32.Value = 0
    
End Sub


Private Sub Graba_Click()
    
    WAtributo1 = ""
    WAtributo2 = ""
    WAtributo3 = ""
    WAtributo4 = ""
    WAtributo5 = ""
    WAtributo6 = ""
    WAtributo7 = ""
    WAtributo8 = ""
    WAtributo9 = ""
    WAtributo10 = ""
    
    
    
    If Opcion1.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion2.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion3.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion4.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    
    
    
    If Opcion11.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion12.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion13.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion14.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion15.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion16.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    
    
    
    
    If opcion21.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion22.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion23.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion24.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion25.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion26.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion27.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion28.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion29.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion210.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion211.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion212.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion213.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion214.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion215.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion216.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion217.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion218.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion219.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    
    
    If Opcion31.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion32.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    
    txtUserName = "SA"
    txtPassword = "Sw58125812"
    txtOdbc = "Fragancias"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            
    ZSql = ""
    ZSql = ZSql + "UPDATE OperadorIngreso SET "
    ZSql = ZSql + " Campo11 = " + "'" + WAtributo1 + "',"
    ZSql = ZSql + " Campo12 = " + "'" + WAtributo2 + "',"
    ZSql = ZSql + " Campo13 = " + "'" + WAtributo3 + "',"
    ZSql = ZSql + " Campo14 = " + "'" + WAtributo4 + "'"
    ZSql = ZSql + " Where Operador = " + "'" + Operador.Text + "'"
    ZSql = ZSql + " and Modulo = " + "'" + "1" + "'"
    spOperadorIngreso = ZSql
    Set rstOperadorIngreso = db.OpenRecordset(spOperadorIngreso, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    txtUserName = "SA"
    txtPassword = "Sw58125812"
    txtOdbc = "FraganciasII"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            
    ZSql = ""
    ZSql = ZSql + "UPDATE OperadorIngreso SET "
    ZSql = ZSql + " Campo11 = " + "'" + WAtributo1 + "',"
    ZSql = ZSql + " Campo12 = " + "'" + WAtributo2 + "',"
    ZSql = ZSql + " Campo13 = " + "'" + WAtributo3 + "',"
    ZSql = ZSql + " Campo14 = " + "'" + WAtributo4 + "'"
    ZSql = ZSql + " Where Operador = " + "'" + Operador.Text + "'"
    ZSql = ZSql + " and Modulo = " + "'" + "1" + "'"
    spOperadorIngreso = ZSql
    Set rstOperadorIngreso = db.OpenRecordset(spOperadorIngreso, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    txtUserName = "SA"
    txtPassword = "Sw58125812"
    txtOdbc = "FraganciasII"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    
    Operador.Text = ""
    DesOperador.Caption = ""
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
                    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0

    opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion210.Value = 0
    Opcion211.Value = 0
    Opcion212.Value = 0
    Opcion213.Value = 0
    Opcion214.Value = 0
    Opcion215.Value = 0
    Opcion216.Value = 0
    Opcion217.Value = 0
    Opcion218.Value = 0
    Opcion219.Value = 0
                    
    Opcion31.Value = 0
    Opcion32.Value = 0
    
    Operador.SetFocus
    
    Tablas.Tab = 0

End Sub

Private Sub Operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Operador.Text <> "" Then
            
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operador"
            ZSql = ZSql + " Where Operador.OPerador = " + "'" + Operador.Text + "'"
            spOperador = ZSql
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                DesOperador.Caption = rstOperador!Nombre
                Clave.Text = rstOperador!Clave
                ClaveII.Text = rstOperador!ClaveII
                rstOperador.Close
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM OperadorIngreso"
                ZSql = ZSql + " Where OperadorIngreso.OPerador = " + "'" + Operador.Text + "'"
                ZSql = ZSql + " and OperadorIngreso.Modulo = " + "'" + "1" + "'"
                spOperadorIngreso = ZSql
                Set rstOperadorIngreso = db.OpenRecordset(spOperadorIngreso, dbOpenSnapshot, dbSQLPassThrough)
                If rstOperadorIngreso.RecordCount > 0 Then
                    
                    Opcion1.Value = 0
                    Opcion2.Value = 0
                    Opcion3.Value = 0
                    Opcion4.Value = 0
                        
                    Opcion11.Value = 0
                    Opcion12.Value = 0
                    Opcion13.Value = 0
                    Opcion14.Value = 0
                    Opcion15.Value = 0
                    Opcion16.Value = 0
    
                    opcion21.Value = 0
                    Opcion22.Value = 0
                    Opcion23.Value = 0
                    Opcion24.Value = 0
                    Opcion25.Value = 0
                    Opcion26.Value = 0
                    Opcion27.Value = 0
                    Opcion28.Value = 0
                    Opcion29.Value = 0
                    Opcion210.Value = 0
                    Opcion211.Value = 0
                    Opcion212.Value = 0
                    Opcion213.Value = 0
                    Opcion214.Value = 0
                    Opcion215.Value = 0
                    Opcion216.Value = 0
                    Opcion217.Value = 0
                    Opcion218.Value = 0
                    Opcion219.Value = 0
                        
                    Opcion31.Value = 0
                    Opcion32.Value = 0
                
                
                    Opcion1.Value = Val(Mid$(rstOperadorIngreso!Campo11, 1, 1))
                    Opcion2.Value = Val(Mid$(rstOperadorIngreso!Campo11, 2, 1))
                    Opcion3.Value = Val(Mid$(rstOperadorIngreso!Campo11, 3, 1))
                    Opcion4.Value = Val(Mid$(rstOperadorIngreso!Campo11, 4, 1))
                    
                    Opcion11.Value = Val(Mid$(rstOperadorIngreso!Campo12, 1, 1))
                    Opcion12.Value = Val(Mid$(rstOperadorIngreso!Campo12, 2, 1))
                    Opcion13.Value = Val(Mid$(rstOperadorIngreso!Campo12, 3, 1))
                    Opcion14.Value = Val(Mid$(rstOperadorIngreso!Campo12, 4, 1))
                    Opcion15.Value = Val(Mid$(rstOperadorIngreso!Campo12, 5, 1))
                    Opcion16.Value = Val(Mid$(rstOperadorIngreso!Campo12, 6, 1))
                    
                    opcion21.Value = Val(Mid$(rstOperadorIngreso!Campo13, 1, 1))
                    Opcion22.Value = Val(Mid$(rstOperadorIngreso!Campo13, 2, 1))
                    Opcion23.Value = Val(Mid$(rstOperadorIngreso!Campo13, 3, 1))
                    Opcion24.Value = Val(Mid$(rstOperadorIngreso!Campo13, 4, 1))
                    Opcion25.Value = Val(Mid$(rstOperadorIngreso!Campo13, 5, 1))
                    Opcion26.Value = Val(Mid$(rstOperadorIngreso!Campo13, 6, 1))
                    Opcion27.Value = Val(Mid$(rstOperadorIngreso!Campo13, 7, 1))
                    Opcion28.Value = Val(Mid$(rstOperadorIngreso!Campo13, 8, 1))
                    Opcion29.Value = Val(Mid$(rstOperadorIngreso!Campo13, 9, 1))
                    Opcion210.Value = Val(Mid$(rstOperadorIngreso!Campo13, 10, 1))
                    Opcion211.Value = Val(Mid$(rstOperadorIngreso!Campo13, 11, 1))
                    Opcion212.Value = Val(Mid$(rstOperadorIngreso!Campo13, 12, 1))
                    Opcion213.Value = Val(Mid$(rstOperadorIngreso!Campo13, 13, 1))
                    Opcion214.Value = Val(Mid$(rstOperadorIngreso!Campo13, 14, 1))
                    Opcion215.Value = Val(Mid$(rstOperadorIngreso!Campo13, 15, 1))
                    Opcion216.Value = Val(Mid$(rstOperadorIngreso!Campo13, 16, 1))
                    Opcion217.Value = Val(Mid$(rstOperadorIngreso!Campo13, 17, 1))
                    Opcion218.Value = Val(Mid$(rstOperadorIngreso!Campo13, 18, 1))
                    Opcion219.Value = Val(Mid$(rstOperadorIngreso!Campo13, 19, 1))
                    
                    Opcion31.Value = Val(Mid$(rstOperadorIngreso!Campo14, 1, 1))
                    Opcion32.Value = Val(Mid$(rstOperadorIngreso!Campo14, 2, 1))
                    rstOperadorIngreso.Close
                    
                End If
            
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salida_Click()
    PrgConfigVentas.Hide
    Unload Me
    Menu.Show
End Sub

