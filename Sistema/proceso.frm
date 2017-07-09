VERSION 5.00
Begin VB.Form PrgProceso 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Grabacion de Cuentas Contables"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "PrgProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private Vector(10000) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim XParam As String

Private Sub Cancelar_Click()

    PrgProceso.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

    With rstCue
        .Index = "Cuenta"
        .MoveFirst
        Do
        
            If !Imputable = 1 Then
        
            WCuenta = !Cuenta
            WDescripcion = !Descripcion
        
            XParam = "'" + WCuenta + "','" _
                        + WDescripcion + "','" _
                        + "0" + "','" _
                        + "1" + "'"
        
            spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                Set rstCuenta = db.OpenRecordset("ModificaCuenta " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                Set rstCuenta = db.OpenRecordset("AltaCuenta " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            End If
        
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    


    
    
    Call Cancelar_Click

End Sub

