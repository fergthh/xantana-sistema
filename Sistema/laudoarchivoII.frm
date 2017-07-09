VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgLaudoArchivoII 
   AutoRedraw      =   -1  'True
   Caption         =   "Seleccion de Planilla Excel a Importar"
   ClientHeight    =   6690
   ClientLeft      =   2850
   ClientTop       =   690
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   3975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   1080
         TabIndex        =   2
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   3600
         MouseIcon       =   "laudoarchivoII.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "laudoarchivoII.frx":030A
         ToolTipText     =   "Cancela el Proceso"
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Acepta 
         Height          =   480
         Left            =   2040
         MouseIcon       =   "laudoarchivoII.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "laudoarchivoII.frx":0A56
         ToolTipText     =   "Confirma el Proceso"
         Top             =   5280
         Width           =   480
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Listleg.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgLaudoArchivoII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    WDrive = Drive1.Drive
    WDir = Dir1.Path
    WLon = Len(WDir)
    If Right$(WDir, 1) = "\" Then
        WDir = Mid(WDir, 1, WLon - 1)
    End If
    XNombre = WDir + "\"
    
    ZZPasoArchivo = XNombre + File1.filename
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    ZZProcesoLaudo = 1
    PrgLaudoArchivoII.Hide
    Unload Me
    PrgLaudoRepuesto.Show
End Sub

Private Sub Form_Activate()
    ZZPasoArchivo = ""
    File1.Pattern = "*.xls"
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub

Sub Form_Load()
    Frame2.Visible = True
End Sub

