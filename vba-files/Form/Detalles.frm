VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Detalles 
   Caption         =   "Detalles"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6270
   OleObjectBlob   =   "Detalles.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Detalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncancelar_Click()
Dim answer As Integer
answer = MsgBox("¿Desea eliminar este registro?", vbQuestion + vbYesNo + vbDefaultButton2, "Clientes")
'Comprobar si acepta el cuadro de dialogo
If answer = vbYes Then
  'rs_delete.Open "DELETE * FROM categorias where categoria like '" & Val(txtcodigo.Value) & "' ", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  MsgBox ("Eliminado Correctamente")
 End If
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Call Conecta

    '''''''''''''''''''''''''''''''''''''''''''''''''''
    cbarticulo.Text = "Seleccioone un Articulo"
    
    Set rscantidad = New ADODB.Recordset
    rscantidad.Open "SELECT count(*) from articulos", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
     
    Set rsarticulos = New ADODB.Recordset
    rsarticulos.Open "SELECT NomArticulo & '-'& IdArticulo as nombre_codigo  from articulos", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    'le damos los items al cuadro combinado
    cbarticulo.AddItem (rsarticulos.Fields("nombre_codigo"))
   'iteramos para agregar los demas items al cuadro combinado
    Dim articulos As Integer
    articulos = Val(rscantidad.Fields("Expr1000")) - 1

    For x = 1 To articulos Step 1
    rsarticulos.MoveNext
    cbarticulo.AddItem (rsarticulos.Fields("nombre_codigo"))
    Next

    txtcantidad.Value = 0
    txtpreventa.Value = 0
    
End Sub
