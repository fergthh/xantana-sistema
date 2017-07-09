VERSION 5.00
Begin VB.Form PrgPonerCero 
   AutoRedraw      =   -1  'True
   Caption         =   "Pone en cero los stock "
   ClientHeight    =   2715
   ClientLeft      =   3390
   ClientTop       =   1650
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   5790
   Begin VB.CommandButton CmdClose 
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
      Left            =   3240
      MouseIcon       =   "PonerCero.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "PonerCero.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salida"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
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
      Left            =   1440
      MouseIcon       =   "PonerCero.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "PonerCero.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "PrgPonerCero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    
    T$ = "Poner stock en cero"
    m$ = "Desea Borrar el stock "
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Stock = 0"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Insumo SET "
        ZSql = ZSql + " Stock = 0"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        
        m$ = "Proceso finalizado"
        aaaaaa% = MsgBox(m$, 0, "POner stock en cero")
    
    End If
    
End Sub

Private Sub cmdClose_Click()
    PrgPonerCero.Hide
    Unload Me
    MenuVen.Show
End Sub




























