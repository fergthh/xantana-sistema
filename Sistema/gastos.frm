VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGastos 
   AutoRedraw      =   -1  'True
   Caption         =   "Gastos de Importacion"
   ClientHeight    =   8235
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   11730
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   11730
   Begin VB.Frame Frame2 
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11535
      Begin Crystal.CrystalReport Listado 
         Left            =   5760
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.Frame PantaImpre 
         Height          =   3015
         Left            =   2880
         TabIndex        =   78
         Top             =   960
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton CancelaImpre 
            Caption         =   "Cancela"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   3120
            MouseIcon       =   "gastos.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "gastos.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Salida"
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton ConfirmaImpre 
            Caption         =   "Confirma"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1560
            MouseIcon       =   "gastos.frx":0B4C
            MousePointer    =   99  'Custom
            Picture         =   "gastos.frx":0E56
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Graba los Datos Ingresados"
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox Margen 
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
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   82
            Text            =   " "
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Paridad 
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
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   81
            Text            =   " "
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label28 
            Caption         =   "Margen"
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
            Left            =   240
            TabIndex        =   80
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label27 
            Caption         =   "Paridad"
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
            Left            =   240
            TabIndex        =   79
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.CommandButton Lista 
         Caption         =   "Listado F9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10440
         MouseIcon       =   "gastos.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "gastos.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Impresion "
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox Total2 
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
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   76
         Text            =   " "
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox Total1 
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
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   75
         Text            =   " "
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox Impo19 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   74
         Text            =   " "
         Top             =   7440
         Width           =   1695
      End
      Begin VB.TextBox Impo19No 
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   73
         Text            =   " "
         Top             =   7440
         Width           =   1695
      End
      Begin VB.TextBox Impo18 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   72
         Text            =   " "
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox Impo18No 
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   71
         Text            =   " "
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox Impo17 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   70
         Text            =   " "
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox Impo16 
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
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   69
         Text            =   " "
         Top             =   6360
         Width           =   1695
      End
      Begin VB.TextBox Impo15 
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
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   68
         Text            =   " "
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox Impo14 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   67
         Text            =   " "
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox Impo13 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   66
         Text            =   " "
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox Impo13No 
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   65
         Text            =   " "
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox Impo12 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   64
         Text            =   " "
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox Impo11 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   63
         Text            =   " "
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox Descri19 
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
         MaxLength       =   20
         TabIndex        =   61
         Text            =   " "
         Top             =   7440
         Width           =   2655
      End
      Begin VB.TextBox Descri18 
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
         MaxLength       =   20
         TabIndex        =   59
         Text            =   " "
         Top             =   7080
         Width           =   2655
      End
      Begin VB.TextBox Descri17 
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
         MaxLength       =   20
         TabIndex        =   57
         Text            =   " "
         Top             =   6720
         Width           =   2655
      End
      Begin VB.TextBox Descri16 
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
         MaxLength       =   20
         TabIndex        =   55
         Text            =   " "
         Top             =   6360
         Width           =   2655
      End
      Begin VB.TextBox Descri15 
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
         MaxLength       =   20
         TabIndex        =   53
         Text            =   " "
         Top             =   6000
         Width           =   2655
      End
      Begin VB.TextBox Descri14 
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
         MaxLength       =   20
         TabIndex        =   51
         Text            =   " "
         Top             =   5640
         Width           =   2655
      End
      Begin VB.TextBox Descri13 
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
         MaxLength       =   20
         TabIndex        =   49
         Text            =   " "
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox Descri12 
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
         MaxLength       =   20
         TabIndex        =   47
         Text            =   " "
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox Descri11 
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
         MaxLength       =   20
         TabIndex        =   45
         Text            =   " "
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox Impo10 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   44
         Text            =   " "
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox Descri10 
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
         MaxLength       =   20
         TabIndex        =   42
         Text            =   " "
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox Impo9 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   41
         Text            =   " "
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox Descri9 
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
         MaxLength       =   20
         TabIndex        =   39
         Text            =   " "
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox Impo8 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   38
         Text            =   " "
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Descri8 
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
         MaxLength       =   20
         TabIndex        =   36
         Text            =   " "
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox Impo7 
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   35
         Text            =   " "
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Descri7 
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
         MaxLength       =   20
         TabIndex        =   33
         Text            =   " "
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox Impo6 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   32
         Text            =   " "
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Descri6 
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
         MaxLength       =   20
         TabIndex        =   30
         Text            =   " "
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox Impo5 
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   29
         Text            =   " "
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Descri5 
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
         MaxLength       =   20
         TabIndex        =   27
         Text            =   " "
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox Impo4 
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   26
         Text            =   " "
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Descri4 
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
         MaxLength       =   20
         TabIndex        =   24
         Text            =   " "
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Impo3 
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   23
         Text            =   " "
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Descri3 
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
         MaxLength       =   20
         TabIndex        =   21
         Text            =   " "
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox Impo2 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   20
         Text            =   " "
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Descri2 
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
         MaxLength       =   20
         TabIndex        =   18
         Text            =   " "
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Impo1 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   13
         Text            =   " "
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpia F3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10440
         MouseIcon       =   "gastos.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "gastos.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpia la pantalla"
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Proceso 
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
         Height          =   975
         Left            =   10440
         MouseIcon       =   "gastos.frx":2D30
         MousePointer    =   99  'Custom
         Picture         =   "gastos.frx":303A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Proveedor 
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
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   9
         Text            =   " "
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Codigo 
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Descri1 
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
         MaxLength       =   20
         TabIndex        =   3
         Text            =   " "
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Menu F10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10440
         MouseIcon       =   "gastos.frx":387C
         MousePointer    =   99  'Custom
         Picture         =   "gastos.frx":3B86
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salida"
         Top             =   5040
         Width           =   855
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label26 
         Caption         =   "Varios"
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
         Left            =   120
         TabIndex        =   62
         Top             =   7440
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "Varios"
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
         Left            =   120
         TabIndex        =   60
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "PLQPP"
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
         Left            =   120
         TabIndex        =   58
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "Textil"
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
         Left            =   120
         TabIndex        =   56
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "Valor Criterio"
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
         Left            =   120
         TabIndex        =   54
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Fedex"
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
         Left            =   120
         TabIndex        =   52
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Gtos. Bancarios"
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
         Left            =   120
         TabIndex        =   50
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Gtos. Transferencia"
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
         Left            =   120
         TabIndex        =   48
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label18 
         Caption         =   "Honorarios"
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
         Left            =   120
         TabIndex        =   46
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "Dpto Fiscal"
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
         Left            =   120
         TabIndex        =   43
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Acarreo"
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
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Flete"
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
         Left            =   120
         TabIndex        =   37
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Ingresos Brutos CF"
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
         Left            =   120
         TabIndex        =   34
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Oficializacion"
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
         Left            =   120
         TabIndex        =   31
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Imp. Ganancias"
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
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "IVA Adicional"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "IVA"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Tasa Estadistica"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Textil/Criterio"
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
         Left            =   8640
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recuperable"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Recuperable"
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
         Left            =   6840
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
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
         Left            =   2280
         TabIndex        =   14
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label DesProveedor 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   7560
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Carpeta"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Proveedor"
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
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Derechos Importacion"
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
   End
End
Attribute VB_Name = "PrgGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(100, 10) As String
Dim ZImpo As Double
Dim ZImpoTextil As Double
Dim ZImpoCriterio As Double
Dim ZPorceGastos As Double
Dim ZPorceTextil As Double
Dim ZPorceCriterio As Double

Private Sub CancelaImpre_Click()

    PantaImpre.Visible = False

End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Fecha.Text = "  /  /    "
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Total1.Text = ""
    Total2.Text = ""
    
    Impo1.Text = ""
    Descri1.Text = ""
    Impo2.Text = ""
    Descri2.Text = ""
    Impo3.Text = ""
    Descri3.Text = ""
    Impo4.Text = ""
    Descri4.Text = ""
    Impo5.Text = ""
    Descri5.Text = ""
    Impo6.Text = ""
    Descri6.Text = ""
    Impo7.Text = ""
    Descri7.Text = ""
    Impo8.Text = ""
    Descri8.Text = ""
    Impo9.Text = ""
    Descri9.Text = ""
    Impo10.Text = ""
    Descri10.Text = ""
    Impo11.Text = ""
    Descri11.Text = ""
    Impo12.Text = ""
    Descri12.Text = ""
    Impo13.Text = ""
    Impo13No.Text = ""
    Descri13.Text = ""
    Impo14.Text = ""
    Descri14.Text = ""
    Impo15.Text = ""
    Descri15.Text = ""
    Impo16.Text = ""
    Descri16.Text = ""
    Impo17.Text = ""
    Descri17.Text = ""
    Impo18.Text = ""
    Impo18No.Text = ""
    Descri18.Text = ""
    Impo19.Text = ""
    Impo19No.Text = ""
    Descri19.Text = ""
    Call Calcula
    
    Codigo.SetFocus

End Sub



Private Sub ConfirmaImpre_Click()

    Erase ZVector
    WNeto = 0
    WNetoII = 0
    WNetoIII = 0
    WNetoIV = 0
    WNetoV = 0
    
    For WRenglon = 1 To 100
    
        Auxi = Codigo.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenImportacion"
        ZSql = ZSql + " Where OrdenImportacion.Clave = " + "'" + WClave + "'"
        spOrdenImportacion = ZSql
        Set rstOrdenImportacion = db.OpenRecordset(spOrdenImportacion, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenImportacion.RecordCount > 0 Then
            
            ZVector(WRenglon, 1) = rstOrdenImportacion!Articulo
            ZVector(WRenglon, 3) = Str$(rstOrdenImportacion!Cantidad)
            ZVector(WRenglon, 4) = Str$(rstOrdenImportacion!Fob)
            ZImpo = rstOrdenImportacion!Cantidad * rstOrdenImportacion!Fob
            ZCantidad = rstOrdenImportacion!Cantidad
            Call Redondeo(ZImpo)
            ZVector(WRenglon, 5) = Str$(ZImpo)
            ZVector(WRenglon, 6) = "0"
            ZVector(WRenglon, 7) = rstOrdenImportacion!Clave
            
            WNeto = WNeto + ZImpo
            
            rstOrdenImportacion.Close
                
            ZZTextil = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZVector(WRenglon, 1) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZVector(WRenglon, 2) = rstArticulo!Descripcion
                ZZTextil = IIf(IsNull(rstArticulo!Textil), "0", rstArticulo!Textil)
                ZVector(WRenglon, 6) = Str$(ZZTextil)
                rstArticulo.Close
            End If
            
            If ZZTextil = 1 Then
                WNetoII = WNetoII + ZImpo
                WNetoIV = WNetoIV + ZCantidad
            End If
            
            If ZZTextil = 2 Then
                WNetoIII = WNetoIII + ZImpo
                WNetoV = WNetoV + ZCantidad
            End If
                    
        End If
    
    Next WRenglon
    
    ZPorceGastos = 0
    If WNeto <> 0 Then
        ZPorceGastos = (Val(Total2.Text) / WNeto) * 100
        Call Redondeo(ZPorceGastos)
    End If
    
    ZPorceTextil = 0
    If WNetoII <> 0 Then
        ZPorceTextil = (Val(Impo16.Text) / WNetoII) * 100
        Call Redondeo(ZPorceTextil)
    End If
    
    ZPorceCriterio = 0
    If WNetoIII <> 0 Then
        ZPorceCriterio = (Val(Impo15.Text) / WNetoIII) * 100
        Call Redondeo(ZPorceCriterio)
    End If
    
    For ZZCiclo = 1 To 100
    
        If ZVector(ZZCiclo, 1) <> "" Then
        
            ZPorce = Val(ZVector(ZZCiclo, 5)) / WNeto
            ZImpo = (Val(Total2.Text) * ZPorce) / Val(ZVector(ZZCiclo, 3))
            Call Redondeo3(ZImpo)
            
            ZImpoTextil = 0
            If Val(ZVector(ZZCiclo, 6)) = 1 Then
                ZPorce = Val(ZVector(ZZCiclo, 3)) / WNetoIV
                ZImpoTextil = (Val(Impo16.Text) * ZPorce) / Val(ZVector(ZZCiclo, 3))
                Call Redondeo3(ZImpoTextil)
            End If
            If Val(ZVector(ZZCiclo, 6)) = 2 Then
                ZPorce = Val(ZVector(ZZCiclo, 3)) / WNetoV
                ZImpoTextil = (Val(Impo15.Text) * ZPorce) / Val(ZVector(ZZCiclo, 3))
                Call Redondeo3(ZImpoTextil)
            End If
            
            ZCif = Val(ZVector(ZZCiclo, 4)) + ZImpo + ZImpoTextil
            ZPesos = ZCif * Val(Paridad.Text)
            ZPrecio = ZPesos * Val(Margen.Text)
            
    
            ZSql = ""
            ZSql = ZSql + "UPDATE OrdenImportacion SET "
            ZSql = ZSql + " PorceGastos = " + "'" + Str$(ZPorceGastos) + "',"
            ZSql = ZSql + " Gastos = " + "'" + Str$(ZImpo) + "',"
            ZSql = ZSql + " PorceTextil = " + "'" + Str$(ZPorceTextil) + "',"
            ZSql = ZSql + " Textil = " + "'" + Str$(ZImpoTextil) + "',"
            ZSql = ZSql + " Paridad = " + "'" + Paridad.Text + "',"
            ZSql = ZSql + " Cif = " + "'" + Str$(ZCif) + "',"
            ZSql = ZSql + " Pesos = " + "'" + Str$(ZPesos) + "',"
            ZSql = ZSql + " Margen = " + "'" + Margen.Text + "',"
            ZSql = ZSql + " Precio = " + "'" + Str$(ZPrecio) + "',"
            ZSql = ZSql + " Total1 = " + "'" + Total1.Text + "',"
            ZSql = ZSql + " Total2 = " + "'" + Total2.Text + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZVector(ZZCiclo, 7) + "'"
            spOrdenImportacion = ZSql
            Set rstOrdenImportacion = db.OpenRecordset(spOrdenImportacion, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
    Next ZZCiclo
    
    Listado.WindowTitle = "Calculo de Costos de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height


    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT OrdenImportacion.Clave, OrdenImportacion.Numero, OrdenImportacion.Proveedor, OrdenImportacion.Articulo, OrdenImportacion.Cantidad, OrdenImportacion.Fob, OrdenImportacion.PorceGastos, OrdenImportacion.Gastos, OrdenImportacion.PorceTextil, OrdenImportacion.Textil, OrdenImportacion.Paridad, OrdenImportacion.Cif, OrdenImportacion.Pesos, OrdenImportacion.Margen, OrdenImportacion.Precio, OrdenImportacion.Total1, OrdenImportacion.Total2, OrdenImportacion.Fecha, " _
            + "Proveedor.Nombre, " _
            + "Articulo.Descripcion, " _
            + "Gastos.Impo1, Gastos.Descri1, Gastos.Impo2, Gastos.Descri2, Gastos.Impo3, Gastos.Descri3, Gastos.Impo4, Gastos.Descri4, Gastos.Impo5, Gastos.Descri5, Gastos.Impo6, Gastos.Descri6, Gastos.Impo7, Gastos.Descri7, Gastos.Impo8, Gastos.Descri8, Gastos.Impo9, Gastos.Descri9, Gastos.Impo10, Gastos.Descri10, Gastos.Impo11, Gastos.Descri11, Gastos.Impo12, Gastos.Descri12, Gastos.Impo13, Gastos.Impo13No, Gastos.Descri13, Gastos.Impo14, Gastos.Descri14, Gastos.Impo15, Gastos.Descri15, Gastos.Impo16, Gastos.Descri16, Gastos.Impo17, Gastos.Descri17, Gastos.Impo18, Gastos.Impo18No, Gastos.Descri18, Gastos.Impo19, Gastos.Impo19No, Gastos.Descri19 " _
            + "From " _
            + DSQ + ".dbo.OrdenImportacion OrdenImportacion, " _
            + DSQ + ".dbo.Proveedor Proveedor, " _
            + DSQ + ".dbo.Articulo Articulo, " _
            + DSQ + ".dbo.Gastos Gastos " _
            + "Where " _
            + "OrdenImportacion.Proveedor = Proveedor.Proveedor AND " _
            + "OrdenImportacion.Articulo = Articulo.Codigo AND " _
            + "OrdenImportacion.Numero = Gastos.Codigo AND " _
            + "OrdenImportacion.Numero >= " + Codigo.Text + " AND " _
            + "OrdenImportacion.Numero <= " + Codigo.Text
    
    Uno = "{OrdenImportacion.Numero} in " + Codigo.Text + " to " + Codigo.Text
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
            
    Listado.ReportFileName = "ImpreOrdenImpoGastos.rpt"
    
    Listado.Destination = 1
    Listado.Destination = 0
    
    Listado.Action = 1

    PantaImpre.Visible = False


End Sub

Private Sub Lista_Click()

    XCodigo = Codigo.Text
    Call Proceso_Click
    Codigo.Text = XCodigo
    Call Codigo_KeyPress(13)
    
    Paridad.Text = ""
    Margen.Text = ""
    
    PantaImpre.Visible = True
    Paridad.SetFocus
    

End Sub

Private Sub Proceso_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Gastos"
    ZSql = ZSql + " Where Gastos.Codigo = " + "'" + Codigo.Text + "'"
    spGastos = ZSql
    Set rstGastos = db.OpenRecordset(spGastos, dbOpenSnapshot, dbSQLPassThrough)
    If rstGastos.RecordCount > 0 Then
    
        rstGastos.Close
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Gastos SET "
        ZSql = ZSql + " Impo1 = " + "'" + Impo1.Text + "',"
        ZSql = ZSql + " Descri1 = " + "'" + Descri1.Text + "',"
        ZSql = ZSql + " Impo2 = " + "'" + Impo2.Text + "',"
        ZSql = ZSql + " Descri2 = " + "'" + Descri2.Text + "',"
        ZSql = ZSql + " Impo3 = " + "'" + Impo3.Text + "',"
        ZSql = ZSql + " Descri3 = " + "'" + Descri3.Text + "',"
        ZSql = ZSql + " Impo4 = " + "'" + Impo4.Text + "',"
        ZSql = ZSql + " Descri4 = " + "'" + Descri4.Text + "',"
        ZSql = ZSql + " Impo5 = " + "'" + Impo5.Text + "',"
        ZSql = ZSql + " Descri5 = " + "'" + Descri5.Text + "',"
        ZSql = ZSql + " Impo6 = " + "'" + Impo6.Text + "',"
        ZSql = ZSql + " Descri6 = " + "'" + Descri6.Text + "',"
        ZSql = ZSql + " Impo7 = " + "'" + Impo7.Text + "',"
        ZSql = ZSql + " Descri7 = " + "'" + Descri7.Text + "',"
        ZSql = ZSql + " Impo8 = " + "'" + Impo8.Text + "',"
        ZSql = ZSql + " Descri8 = " + "'" + Descri8.Text + "',"
        ZSql = ZSql + " Impo9 = " + "'" + Impo9.Text + "',"
        ZSql = ZSql + " Descri9 = " + "'" + Descri9.Text + "',"
        ZSql = ZSql + " Impo10 = " + "'" + Impo10.Text + "',"
        ZSql = ZSql + " Descri10 = " + "'" + Descri10.Text + "',"
        ZSql = ZSql + " Impo11 = " + "'" + Impo11.Text + "',"
        ZSql = ZSql + " Descri11 = " + "'" + Descri11.Text + "',"
        ZSql = ZSql + " Impo12 = " + "'" + Impo12.Text + "',"
        ZSql = ZSql + " Descri12 = " + "'" + Descri12.Text + "',"
        ZSql = ZSql + " Impo13 = " + "'" + Impo13.Text + "',"
        ZSql = ZSql + " Impo13No = " + "'" + Impo13No.Text + "',"
        ZSql = ZSql + " Descri13 = " + "'" + Descri13.Text + "',"
        ZSql = ZSql + " Impo14 = " + "'" + Impo14.Text + "',"
        ZSql = ZSql + " Descri14 = " + "'" + Descri14.Text + "',"
        ZSql = ZSql + " Impo15 = " + "'" + Impo15.Text + "',"
        ZSql = ZSql + " Descri15 = " + "'" + Descri15.Text + "',"
        ZSql = ZSql + " Impo16 = " + "'" + Impo16.Text + "',"
        ZSql = ZSql + " Descri16 = " + "'" + Descri16.Text + "',"
        ZSql = ZSql + " Impo17 = " + "'" + Impo17.Text + "',"
        ZSql = ZSql + " Descri17 = " + "'" + Descri17.Text + "',"
        ZSql = ZSql + " Impo18 = " + "'" + Impo18.Text + "',"
        ZSql = ZSql + " Impo18No = " + "'" + Impo18No.Text + "',"
        ZSql = ZSql + " Descri18 = " + "'" + Descri18.Text + "',"
        ZSql = ZSql + " Impo19 = " + "'" + Impo19.Text + "',"
        ZSql = ZSql + " Impo19No = " + "'" + Impo19No.Text + "',"
        ZSql = ZSql + " Descri19 = " + "'" + Descri19.Text + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        
        spGastos = ZSql
        Set rstGastos = db.OpenRecordset(spGastos, dbOpenSnapshot, dbSQLPassThrough)
    
            Else

        ZSql = ""
        ZSql = ZSql + "INSERT INTO Gastos ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Impo1 ,"
        ZSql = ZSql + "Descri1 ,"
        ZSql = ZSql + "Impo2 ,"
        ZSql = ZSql + "Descri2 ,"
        ZSql = ZSql + "Impo3 ,"
        ZSql = ZSql + "Descri3 ,"
        ZSql = ZSql + "Impo4 ,"
        ZSql = ZSql + "Descri4 ,"
        ZSql = ZSql + "Impo5 ,"
        ZSql = ZSql + "Descri5 ,"
        ZSql = ZSql + "Impo6 ,"
        ZSql = ZSql + "Descri6 ,"
        ZSql = ZSql + "Impo7 ,"
        ZSql = ZSql + "Descri7 ,"
        ZSql = ZSql + "Impo8 ,"
        ZSql = ZSql + "Descri8 ,"
        ZSql = ZSql + "Impo9 ,"
        ZSql = ZSql + "Descri9 ,"
        ZSql = ZSql + "Impo10 ,"
        ZSql = ZSql + "Descri10 ,"
        ZSql = ZSql + "Impo11 ,"
        ZSql = ZSql + "Descri11 ,"
        ZSql = ZSql + "Impo12 ,"
        ZSql = ZSql + "Descri12 ,"
        ZSql = ZSql + "Impo13 ,"
        ZSql = ZSql + "Impo13No ,"
        ZSql = ZSql + "Descri13 ,"
        ZSql = ZSql + "Impo14 ,"
        ZSql = ZSql + "Descri14 ,"
        ZSql = ZSql + "Impo15 ,"
        ZSql = ZSql + "Descri15 ,"
        ZSql = ZSql + "Impo16 ,"
        ZSql = ZSql + "Descri16 ,"
        ZSql = ZSql + "Impo17 ,"
        ZSql = ZSql + "Descri17 ,"
        ZSql = ZSql + "Impo18 ,"
        ZSql = ZSql + "Impo18No ,"
        ZSql = ZSql + "Descri18 ,"
        ZSql = ZSql + "Impo19 ,"
        ZSql = ZSql + "Impo19No ,"
        ZSql = ZSql + "Descri19 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + Impo1.Text + "',"
        ZSql = ZSql + "'" + Descri1.Text + "',"
        ZSql = ZSql + "'" + Impo2.Text + "',"
        ZSql = ZSql + "'" + Descri2.Text + "',"
        ZSql = ZSql + "'" + Impo3.Text + "',"
        ZSql = ZSql + "'" + Descri3.Text + "',"
        ZSql = ZSql + "'" + Impo4.Text + "',"
        ZSql = ZSql + "'" + Descri4.Text + "',"
        ZSql = ZSql + "'" + Impo5.Text + "',"
        ZSql = ZSql + "'" + Descri5.Text + "',"
        ZSql = ZSql + "'" + Impo6.Text + "',"
        ZSql = ZSql + "'" + Descri6.Text + "',"
        ZSql = ZSql + "'" + Impo7.Text + "',"
        ZSql = ZSql + "'" + Descri7.Text + "',"
        ZSql = ZSql + "'" + Impo8.Text + "',"
        ZSql = ZSql + "'" + Descri8.Text + "',"
        ZSql = ZSql + "'" + Impo9.Text + "',"
        ZSql = ZSql + "'" + Descri9.Text + "',"
        ZSql = ZSql + "'" + Impo10.Text + "',"
        ZSql = ZSql + "'" + Descri10.Text + "',"
        ZSql = ZSql + "'" + Impo11.Text + "',"
        ZSql = ZSql + "'" + Descri11.Text + "',"
        ZSql = ZSql + "'" + Impo12.Text + "',"
        ZSql = ZSql + "'" + Descri12.Text + "',"
        ZSql = ZSql + "'" + Impo13.Text + "',"
        ZSql = ZSql + "'" + Impo13No.Text + "',"
        ZSql = ZSql + "'" + Descri13.Text + "',"
        ZSql = ZSql + "'" + Impo14.Text + "',"
        ZSql = ZSql + "'" + Descri14.Text + "',"
        ZSql = ZSql + "'" + Impo15.Text + "',"
        ZSql = ZSql + "'" + Descri15.Text + "',"
        ZSql = ZSql + "'" + Impo16.Text + "',"
        ZSql = ZSql + "'" + Descri16.Text + "',"
        ZSql = ZSql + "'" + Impo17.Text + "',"
        ZSql = ZSql + "'" + Descri17.Text + "',"
        ZSql = ZSql + "'" + Impo18.Text + "',"
        ZSql = ZSql + "'" + Impo18No.Text + "',"
        ZSql = ZSql + "'" + Descri18.Text + "',"
        ZSql = ZSql + "'" + Impo19.Text + "',"
        ZSql = ZSql + "'" + Impo19No.Text + "',"
        ZSql = ZSql + "'" + Descri19.Text + "')"
                                
        spGastos = ZSql
        Set rstGastos = db.OpenRecordset(spGastos, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    Call CmdLimpiar_Click
    
End Sub

Private Sub Cancela_click()
    PrgGastos.Hide
    Unload Me
    Menu23.Show
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenImportacion"
        ZSql = ZSql + " Where OrdenImportacion.Numero = " + "'" + Codigo.Text + "'"
        spOrdenImportacion = ZSql
        Set rstOrdenImportacion = db.OpenRecordset(spOrdenImportacion, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenImportacion.RecordCount > 0 Then
        
            Fecha.Text = rstOrdenImportacion!Fecha
            Proveedor.Text = rstOrdenImportacion!Proveedor
            
            rstOrdenImportacion.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                DesProveedor.Caption = rstProveedor!Nombre
                rstProveedor.Close
                    Else
                DesProveedor.Caption = ""
            End If
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Gastos"
            ZSql = ZSql + " Where Gastos.Codigo = " + "'" + Codigo.Text + "'"
            spGastos = ZSql
            Set rstGastos = db.OpenRecordset(spGastos, dbOpenSnapshot, dbSQLPassThrough)
            If rstGastos.RecordCount > 0 Then
                Impo1.Text = Str$(rstGastos!Impo1)
                Descri1.Text = Trim(rstGastos!Descri1)
                Impo2.Text = Str$(rstGastos!Impo2)
                Descri2.Text = Trim(rstGastos!Descri2)
                Impo3.Text = Str$(rstGastos!Impo3)
                Descri3.Text = Trim(rstGastos!Descri3)
                Impo4.Text = Str$(rstGastos!Impo4)
                Descri4.Text = Trim(rstGastos!Descri4)
                Impo5.Text = Str$(rstGastos!Impo5)
                Descri5.Text = Trim(rstGastos!Descri5)
                Impo6.Text = Str$(rstGastos!Impo6)
                Descri6.Text = Trim(rstGastos!Descri6)
                Impo7.Text = Str$(rstGastos!Impo7)
                Descri7.Text = Trim(rstGastos!Descri7)
                Impo8.Text = Str$(rstGastos!Impo8)
                Descri8.Text = Trim(rstGastos!Descri8)
                Impo9.Text = Str$(rstGastos!Impo9)
                Descri9.Text = Trim(rstGastos!Descri9)
                Impo10.Text = Str$(rstGastos!Impo10)
                Descri10.Text = Trim(rstGastos!Descri10)
                Impo11.Text = Str$(rstGastos!Impo11)
                Descri11.Text = Trim(rstGastos!Descri11)
                Impo12.Text = Str$(rstGastos!Impo12)
                Descri12.Text = Trim(rstGastos!Descri12)
                Impo13.Text = Str$(rstGastos!Impo13)
                Impo13No.Text = Str$(rstGastos!Impo13No)
                Descri13.Text = Trim(rstGastos!Descri13)
                Impo14.Text = Str$(rstGastos!Impo14)
                Descri14.Text = Trim(rstGastos!Descri14)
                Impo15.Text = Str$(rstGastos!Impo15)
                Descri15.Text = Trim(rstGastos!Descri15)
                Impo16.Text = Str$(rstGastos!Impo16)
                Descri16.Text = Trim(rstGastos!Descri16)
                Impo17.Text = Str$(rstGastos!Impo17)
                Descri17.Text = Trim(rstGastos!Descri17)
                Impo18.Text = Str$(rstGastos!Impo18)
                Impo18No.Text = Str$(rstGastos!Impo18No)
                Descri18.Text = Trim(rstGastos!Descri18)
                Impo19.Text = Str$(rstGastos!Impo19)
                Impo19No.Text = Str$(rstGastos!Impo19No)
                Descri19.Text = Trim(rstGastos!Descri19)
                
                rstGastos.Close
                
                Impo1.Text = Pusing("###,###.##", Impo1.Text)
                Impo2.Text = Pusing("###,###.##", Impo2.Text)
                Impo3.Text = Pusing("###,###.##", Impo3.Text)
                Impo4.Text = Pusing("###,###.##", Impo4.Text)
                Impo5.Text = Pusing("###,###.##", Impo5.Text)
                Impo6.Text = Pusing("###,###.##", Impo6.Text)
                Impo7.Text = Pusing("###,###.##", Impo7.Text)
                Impo8.Text = Pusing("###,###.##", Impo8.Text)
                Impo9.Text = Pusing("###,###.##", Impo9.Text)
                Impo10.Text = Pusing("###,###.##", Impo10.Text)
                Impo11.Text = Pusing("###,###.##", Impo11.Text)
                Impo12.Text = Pusing("###,###.##", Impo12.Text)
                Impo13.Text = Pusing("###,###.##", Impo13.Text)
                Impo13No.Text = Pusing("###,###.##", Impo13No.Text)
                Impo14.Text = Pusing("###,###.##", Impo14.Text)
                Impo15.Text = Pusing("###,###.##", Impo15.Text)
                Impo16.Text = Pusing("###,###.##", Impo16.Text)
                Impo17.Text = Pusing("###,###.##", Impo17.Text)
                Impo18.Text = Pusing("###,###.##", Impo18.Text)
                Impo18No.Text = Pusing("###,###.##", Impo18No.Text)
                Impo19.Text = Pusing("###,###.##", Impo19.Text)
                Impo19No.Text = Pusing("###,###.##", Impo19No.Text)
                
                Call Calcula
                
            End If
            
            Descri1.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Suma()
                
    ZTotal1 = Val(Impo3.Text) + Val(Impo4.Text) + Val(Impo5.Text) + Val(Impo7.Text) + Val(Impo13No.Text) + Val(Impo18No.Text) + Val(Impo19No.Text)
    ZTotal2 = Val(Impo1.Text) + Val(Impo2.Text) + Val(Impo6.Text) + Val(Impo8.Text) + Val(Impo9.Text) + Val(Impo10.Text)
    ZTotal2 = ZTotal2 + Val(Impo11.Text) + Val(Impo12.Text) + Val(Impo13.Text) + Val(Impo14.Text)
    ZTotal2 = ZTotal2 + Val(Impo17.Text) + Val(Impo18.Text) + Val(Impo19.Text)
    
    Total1.Text = Str$(ZTotal1)
    Total2.Text = Str$(ZTotal2)
    
    Total1.Text = Pusing("###,###.##", Total1.Text)
    Total2.Text = Pusing("###,###.##", Total2.Text)

End Sub

Private Sub Descri1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo1.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri1.Text = ""
    End If
End Sub

Private Sub Impo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo1.Text = Pusing("###,###.##", Impo1.Text)
        Call Calcula
        Descri2.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo2.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri2.Text = ""
    End If
End Sub

Private Sub Impo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo2.Text = Pusing("###,###.##", Impo2.Text)
        Call Calcula
        Descri3.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo2.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo3.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri3.Text = ""
    End If
End Sub

Private Sub Impo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo3.Text = Pusing("###,###.##", Impo3.Text)
        Call Calcula
        Descri4.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo3.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo4.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri4.Text = ""
    End If
End Sub

Private Sub Impo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo4.Text = Pusing("###,###.##", Impo4.Text)
        Call Calcula
        Descri5.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo4.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo5.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri5.Text = ""
    End If
End Sub

Private Sub Impo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo5.Text = Pusing("###,###.##", Impo5.Text)
        Call Calcula
        Descri6.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo5.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo6.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri6.Text = ""
    End If
End Sub

Private Sub Impo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo6.Text = Pusing("###,###.##", Impo6.Text)
        Call Calcula
        Descri7.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo6.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo7.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri7.Text = ""
    End If
End Sub

Private Sub Impo7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo7.Text = Pusing("###,###.##", Impo7.Text)
        Call Calcula
        Descri8.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo7.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo8.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri8.Text = ""
    End If
End Sub

Private Sub Impo8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo8.Text = Pusing("###,###.##", Impo8.Text)
        Call Calcula
        Descri9.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo8.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo9.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri9.Text = ""
    End If
End Sub

Private Sub Impo9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo9.Text = Pusing("###,###.##", Impo9.Text)
        Call Calcula
        Descri10.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo9.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo10.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri10.Text = ""
    End If
End Sub

Private Sub Impo10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo10.Text = Pusing("###,###.##", Impo10.Text)
        Call Calcula
        Descri11.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo10.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo11.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri11.Text = ""
    End If
End Sub

Private Sub Impo11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo11.Text = Pusing("###,###.##", Impo11.Text)
        Call Calcula
        Descri12.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo11.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo12.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri12.Text = ""
    End If
End Sub

Private Sub Impo12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo12.Text = Pusing("###,###.##", Impo12.Text)
        Call Calcula
        Descri13.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo12.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo13No.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri13.Text = ""
    End If
End Sub

Private Sub Impo13No_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo13No.Text = Pusing("###,###.##", Impo13No.Text)
        Call Calcula
        Impo13.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo13No.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Impo13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo13.Text = Pusing("###,###.##", Impo13.Text)
        Call Calcula
        Descri14.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo13.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo14.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri14.Text = ""
    End If
End Sub

Private Sub Impo14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo14.Text = Pusing("###,###.##", Impo14.Text)
        Call Calcula
        Descri15.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo14.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo15.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri15.Text = ""
    End If
End Sub

Private Sub Impo15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo15.Text = Pusing("###,###.##", Impo15.Text)
        Call Calcula
        Descri16.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo15.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo16.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri16.Text = ""
    End If
End Sub

Private Sub Impo16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo16.Text = Pusing("###,###.##", Impo16.Text)
        Call Calcula
        Descri17.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo16.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri17_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo17.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri17.Text = ""
    End If
End Sub

Private Sub Impo17_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo17.Text = Pusing("###,###.##", Impo17.Text)
        Call Calcula
        Descri18.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo17.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descri18_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo18No.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri18.Text = ""
    End If
End Sub

Private Sub Impo18No_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo18No.Text = Pusing("###,###.##", Impo18No.Text)
        Call Calcula
        Impo18.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo18No.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Impo18_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo18.Text = Pusing("###,###.##", Impo18.Text)
        Call Calcula
        Descri19.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo18.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Descri19_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo19No.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri19.Text = ""
    End If
End Sub

Private Sub Impo19no_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo19No.Text = Pusing("###,###.##", Impo19No.Text)
        Call Calcula
        Impo19.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo19No.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Impo19_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Impo19.Text = Pusing("###,###.##", Impo19.Text)
        Call Calcula
        Descri1.SetFocus
    End If
    If KeyAscii = 27 Then
        Impo19.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Paridad.Text = Pusing("###,###.##", Paridad.Text)
        Margen.SetFocus
    End If
    If KeyAscii = 27 Then
        Paridad.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Margen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Margen.Text = Pusing("###,###.##", Margen.Text)
        Paridad.SetFocus
    End If
    If KeyAscii = 27 Then
        Margen.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Sub Form_Load()

    Codigo.Text = ""
    Fecha.Text = "  /  /    "
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Total1.Text = ""
    Total2.Text = ""
    
    Impo1.Text = ""
    Descri1.Text = ""
    Impo2.Text = ""
    Descri2.Text = ""
    Impo3.Text = ""
    Descri3.Text = ""
    Impo4.Text = ""
    Descri4.Text = ""
    Impo5.Text = ""
    Descri5.Text = ""
    Impo6.Text = ""
    Descri6.Text = ""
    Impo7.Text = ""
    Descri7.Text = ""
    Impo8.Text = ""
    Descri8.Text = ""
    Impo9.Text = ""
    Descri9.Text = ""
    Impo10.Text = ""
    Descri10.Text = ""
    Impo11.Text = ""
    Descri11.Text = ""
    Impo12.Text = ""
    Descri12.Text = ""
    Impo13.Text = ""
    Impo13No.Text = ""
    Descri13.Text = ""
    Impo14.Text = ""
    Descri14.Text = ""
    Impo15.Text = ""
    Descri15.Text = ""
    Impo16.Text = ""
    Descri16.Text = ""
    Impo17.Text = ""
    Descri17.Text = ""
    Impo18.Text = ""
    Impo18No.Text = ""
    Descri18.Text = ""
    Impo19.Text = ""
    Impo19No.Text = ""
    Descri19.Text = ""
    
End Sub

Private Sub Impo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo12_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri12_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo13_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri13_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo14_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri14_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo15_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri15_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo16_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri16_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo17_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri17_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo18_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo18no_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri18_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo19_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impo19no_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descri19_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub

Private Sub Calcula()

    ZTotal1 = Val(Impo3.Text) + Val(Impo4.Text) + Val(Impo5.Text) + Val(Impo7.Text) + Val(Impo13No.Text) + Val(Impo18No.Text) + Val(Impo19No.Text)
    ZTotal2 = Val(Impo1.Text) + Val(Impo2.Text) + Val(Impo6.Text) + Val(Impo8.Text) + Val(Impo9.Text) + Val(Impo10.Text) + Val(Impo11.Text) + Val(Impo12.Text) + Val(Impo13.Text) + Val(Impo14.Text) + Val(Impo17.Text) + Val(Impo18.Text) + Val(Impo19.Text)
    
    Total1.Text = Str$(ZTotal1)
    Total2.Text = Str$(ZTotal2)
    
    Total1.Text = Pusing("###,###.##", Total1.Text)
    Total2.Text = Pusing("###,###.##", Total2.Text)

End Sub













