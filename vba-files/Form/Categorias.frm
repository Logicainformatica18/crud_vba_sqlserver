VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Categorias 
   Caption         =   "CATEGORIA"
   ClientHeight    =   5880
   ClientLeft      =   -30
   ClientTop       =   315
   ClientWidth     =   5070
   OleObjectBlob   =   "Categorias.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "Categorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnanterior_Click()
' On error resumen next se ejecuta si no hay error
On Error Resume Next
'mueve al anterior registro
rs.MovePrevious
txtcodigo.Text = rs.Fields("Categoria")
txtnombre.Text = rs.Fields("Nombre")
End Sub

Private Sub btnbuscar_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs_search = New ADODB.Recordset
'ejecutamos una consulta a la tabla  categorias de la base de datos pcventas
'SELECT categorias.categoria, categorias.nombre FROM categorias where categorias.nombre Like "*"
rs_search.Open "SELECT categoria,nombre from  categorias where nombre Like '" & txtcriterio.Text & "%'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
'on error resume next ejecuta si no hay error
On Error Resume Next
'asignamos los numeros de columnas
With Me.lscategorias
    .ColumnCount = rs_search.Fields.Count
    End With
    'movernos al inicio
    rs_search.MoveFirst
    Dim i As Integer
    i = 1
    With Me.lscategorias
    .Clear
    'aï¿½adir los encabezados
    .AddItem
    'aca van los numeros de columnas (la columna uno es el indice cero,la segunda el indice uno y asi sucesivamente)
    For j = 0 To 1
    .List(0, j) = rs_search.Fields(j).Name
    Next j
    'llenado de los registros de la tabla categorias al listbox (cuadro de lista)
    Do
    .AddItem
    .List(i, 0) = rs_search![categoria]
    .List(i, 1) = rs_search![nombre]
    i = i + 1
    'avansamos un registro siguiente
    rs_search.MoveNext
    Loop Until rs_search.EOF

End With
'cerramos la conexion
miConexion.Close
End Sub

Private Sub btncancelar_Click()
btnanterior.Enabled = True
btnsiguiente.Enabled = True
btnprimero.Enabled = True
btnultimo.Enabled = True

btnguardar.Enabled = False
End Sub

Private Sub btneliminar_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs_delete = New ADODB.Recordset


Dim answer As Integer
answer = MsgBox("Â¿Desea eliminar este registro?", vbQuestion + vbYesNo + vbDefaultButton2, "Clientes")
'Comprobar si acepta el cuadro de dialogo
If answer = vbYes Then
  rs_delete.Open "DELETE  FROM categorias where categoria like '" & Val(txtcodigo.Value) & "' ", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  MsgBox ("Eliminado Correctamente")
  'invocar a la funcion
 Call reset
 End If

End Sub

Private Sub btnguardar_Click()
'on error resumen next se ejecuta si no hay error
On Error Resume Next
'agrega un nuevo registro
rs.AddNew
'asignamos los valores a las celdas de los campos
rs.Fields("categoria") = txtcodigo.Text
rs.Fields("nombre") = txtnombre.Text
'guarda registro
rs.Save
'mensaje de éxito
MsgBox ("Guardado correctamente")
'activa los controles de navegación
btnprimero.Enabled = True
btnultimo.Enabled = True
btnsiguiente.Enabled = True
btnanterior.Enabled = True
'bloquea el boton guardar
btnguardar.Enabled = False

End Sub

Private Sub btnmodificar_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs = New ADODB.Recordset
'ejecutamos una consulta a la tabla clientes de la base de datos pcventas
  rs.Open "UPDATE Categorias set  nombre='" & txtnombre.Text & "' where categoria like '" & txtcodigo & "' ;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

'Call activar_controles
MsgBox ("Modificado Correctamente")
Call reset
End Sub

Private Sub btnnuevo_Click()
'limpia los controles
txtcodigo.Text = ""
txtnombre.Text = ""
'activa el control guardar
btnguardar.Enabled = True
'bloquea los controles de navegacion
btnprimero.Enabled = False
btnultimo.Enabled = False
btnsiguiente.Enabled = False
btnanterior.Enabled = False
'activar el cursor de la caja de texto código
txtcodigo.SetFocus

End Sub

Private Sub btnprimero_Click()
'on error resumen next se ejecuta si no hay error
On Error Resume Next
' mueve al anterior registro
rs.MoveFirst
txtcodigo.Text = rs.Fields("Categoria")
txtnombre.Text = rs.Fields("nombre")

End Sub

Private Sub btnsiguiente_Click()
On Error Resume Next
rs.MoveNext
txtcodigo.Text = rs.Fields("categoria")
txtnombre.Text = rs.Fields("Nombre")

End Sub

Private Sub btnultimo_Click()
On Error Resume Next
rs.MoveLast
txtcodigo.Text = rs.Fields("Categoria")
txtnombre.Text = rs.Fields("Nombre")

End Sub

Private Sub Label1_Click()

End Sub

Private Sub txtcriterio_Change()

End Sub

Private Sub UserForm_Initialize()
Call reset

End Sub



Sub reset()
'ejecuta un procedimiento conectar
Call Conecta
Set rs = New ADODB.Recordset
'abre una consulta a la base de datos de la tabla categorias
rs.Open "SELECT Categorias.Categoria, Categorias.Nombre FROM Categorias;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
'asignamos los valores de la primera fila a las cajas de texto
txtcodigo.Text = rs.Fields("Categoria")
txtnombre.Text = rs.Fields("Nombre")
'bloquea el boton guardar
btnguardar.Enabled = False
End Sub
