VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Proyecto1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileNew 
         Caption         =   "Nuev&o"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Ce&rrar"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "G&uardar como..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Guardar &todo"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Propiedades"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Co&nfigurar página..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "&Vista preliminar"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "I&mprimir..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "&Enviar..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edición"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cor&tar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Peg&ado especial..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&Nueva ventana"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Casca&da"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Mosaico &horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Mosaico &vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "Organi&zar iconos"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Buscar Ayuda acerca de..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de Proyecto1..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)




Private Sub mnuHelpContents_Click()
    

    Dim nRet As Integer


    'Si no hay archivo de Ayuda para este proyecto, muestra un mensaje al usuario
    'puede establecer el archivo de Ayuda para su aplicación en el cuadro de
    'diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "Imposible mostrar los contenidos de la Ayuda. No hay una Ayuda asociada con este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuHelpSearch_Click()
    

    Dim nRet As Integer


    'Si no hay archivo de Ayuda para este proyecto, muestra un mensaje al usuario
    'puede establecer el archivo de Ayuda para su aplicación en el cuadro de
    'diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "Imposible mostrar los contenidos de la Ayuda. No hay una Ayuda asociada con este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub


Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub mnuWindowNewWindow_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para nueva ventana"
End Sub


Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub


Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuEditCopy_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para copiar"
End Sub


Private Sub mnuEditCut_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para cortar"
End Sub


Private Sub mnuEditPaste_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para pegar"
End Sub


Private Sub mnuEditPasteSpecial_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de pegado especial"
End Sub


Private Sub mnuEditUndo_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para deshacer"
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
    'Para hacer
        'Establece los indicadores y atributos del
        'control Common Dialog
        .Filter = "Todos los archivos (*.*)|*.*"
        .ShowOpen
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
    'Para hacer
    'Procesa el archivo abierto
End Sub


Private Sub mnuFileClose_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para cerrar"
End Sub


Private Sub mnuFileSave_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para guardar"
End Sub


Private Sub mnuFileSaveAs_Click()
    'Para hacer
    'Configura el control Common Dialog
    'antes de llamar a ShowSave
    dlgCommonDialog.ShowSave
End Sub


Private Sub mnuFileSaveAll_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para guardar todo"
End Sub


Private Sub mnuFileProperties_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de las propiedades"
End Sub


Private Sub mnuFilePageSetup_Click()
    dlgCommonDialog.ShowPrinter
End Sub


Private Sub mnuFilePrintPreview_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de vista previa"
End Sub


Private Sub mnuFilePrint_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de impresión"
End Sub


Private Sub mnuFileSend_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de enviar"
End Sub


Private Sub mnuFileMRU_Click(Index As Integer)
    'Para hacer
    MsgBox "Aquí se sitúa el código de archivos recientes"
End Sub


Private Sub mnuFileExit_Click()
    'Descarga el formulario
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub



