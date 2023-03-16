VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clientes 
   Caption         =   "CLIENTES"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6810
   OleObjectBlob   =   "Clientes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnanterior_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'mueve al registro
rs.MovePrevious
'asignamos los valores a las cajas de texto
txtidcliente.Text = rs.Fields("Idcliente")
txtnombre.Text = rs.Fields("NomCliente")
txtdireccion.Text = rs.Fields("DirCliente")
txttelefono.Text = rs.Fields("TelCliente")
txtemail.Text = rs.Fields("Email")

End Sub


Private Sub btnbuscar_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs_search = New ADODB.Recordset
'ejecutamos una consulta a la tabla  categorias de la base de datos pcventas
'SELECT categorias.categoria, categorias.nombre FROM categorias where categorias.nombre Like "*"
rs_search.Open "SELECT Clientes.Idcliente, Clientes.NomCliente from Clientes where Clientes.NomCliente Like '" & txtcriterio.Text & "%'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
'on error resume next ejecuta si no hay error
On Error Resume Next
'asignamos los numeros de columnas
With Me.ListBox1
    .ColumnCount = rs_search.Fields.Count
    End With
    'movernos al inicio
    rs_search.MoveFirst
    Dim i As Integer
    i = 1
    With Me.ListBox1
    .Clear
    'a�adir los encabezados
    .AddItem
    'aca van los numeros de columnas (la columna uno es el indice cero,la segunda el indice uno y asi sucesivamente)
    For j = 0 To 1
    .List(0, j) = rs_search.Fields(j).Name
    Next j
    'llenado de los registros de la tabla categorias al listbox (cuadro de lista)
    Do
    .AddItem
    .List(i, 0) = rs_search![IDCLIENTE]
    .List(i, 1) = rs_search![NOMCLIENTE]
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

Private Sub btndelete_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs_delete = New ADODB.Recordset


Dim answer As Integer
answer = MsgBox("¿Desea eliminar este registro?", vbQuestion + vbYesNo + vbDefaultButton2, "Clientes")
'Comprobar si acepta el cuadro de dialogo
If answer = vbYes Then
  rs_delete.Open "DELETE * FROM Clientes where Clientes.IdCliente = '" & txtidcliente.Value & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  MsgBox ("Eliminado Correctamente")
  'invocar a la funcion
 Call reset
 
  
End If




End Sub

Private Sub btnguardar_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'agregar nuevo registro
rs.AddNew
'asignamos los valores de las cajas de texto a los campos del registro
rs.Fields("Idcliente") = txtidcliente.Text
rs.Fields("NomCliente") = txtnombre.Text
rs.Fields("DirCliente") = txtdireccion.Text
rs.Fields("TelCliente") = txttelefono.Text
rs.Fields("Email") = txtemail.Text
'guardar registro
rs.Save
'mostramos un mensaje de exito
MsgBox ("Datos Guardados correctamente")


'bloquear controles
btnanterior.Enabled = True
btnsiguiente.Enabled = True
btnprimero.Enabled = True
btnultimo.Enabled = True

End Sub


Private Sub btnnuevo_Click()
'limpiamos las cajas de texto
txtidcliente.Text = ""
txtnombre.Text = ""
txtdireccion.Text = ""
txttelefono.Text = ""
txtemail.Text = ""

'habilitar boton guardar
btnguardar.Enabled = True


'bloquear controles
btnanterior.Enabled = False
btnsiguiente.Enabled = False
btnprimero.Enabled = False
btnultimo.Enabled = False
End Sub

Private Sub btnprimero_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'mueve al registro siguiente
rs.MoveFirst
'asignamos los valores a las cajas de texto
txtidcliente.Text = rs.Fields("Idcliente")
txtnombre.Text = rs.Fields("NomCliente")
txtdireccion.Text = rs.Fields("DirCliente")
txttelefono.Text = rs.Fields("TelCliente")
txtemail.Text = rs.Fields("Email")
End Sub


Private Sub btnsiguiente_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'mueve al registro
rs.MoveNext
'asignamos los valores a las cajas de texto
txtidcliente.Text = rs.Fields("Idcliente")
txtnombre.Text = rs.Fields("NomCliente")
txtdireccion.Text = rs.Fields("DirCliente")
txttelefono.Text = rs.Fields("TelCliente")
txtemail.Text = rs.Fields("Email")
End Sub

Private Sub btnultimo_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'mueve al registro
rs.MoveLast
'asignamos los valores a las cajas de texto
txtidcliente.Text = rs.Fields("Idcliente")
txtnombre.Text = rs.Fields("NomCliente")
txtdireccion.Text = rs.Fields("DirCliente")
txttelefono.Text = rs.Fields("TelCliente")
txtemail.Text = rs.Fields("Email")
End Sub

Private Sub btnupdate_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs = New ADODB.Recordset
'ejecutamos una consulta a la tabla clientes de la base de datos pcventas
  rs.Open "UPDATE Cliente set  Clientes.NomCliente='" & txtnombre.Text & "',Clientes.DirCliente='" & txtdireccion.Text & "', Clientes.TelCliente= '" & txttelefono.Text & "',Clientes.Email='" & txtemail.Text & "' where clientes.idcliente ='" & txtidcliente & "' ;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call activar_controles
MsgBox (txtidcliente.Text)
End Sub

Private Sub UserForm_Initialize()
Call reset

End Sub

Public Function activar_controles()
'asignamos los valores a las cajas de texto
txtidcliente.Text = rs.Fields("Idcliente")
txtnombre.Text = rs.Fields("NomCliente")
txtdireccion.Text = rs.Fields("DirCliente")
txttelefono.Text = rs.Fields("TelCliente")
txtemail.Text = rs.Fields("Email")

'bloquear el boton guardar
btnguardar.Enabled = False
End Function

Public Function reset()
txtidcliente.Enabled = False
'ejecutar el procedimiento conecta
Call Conecta
Set rs = New ADODB.Recordset
'ejecutamos una consulta a la tabla clientes de la base de datos pcventas
  rs.Open "SELECT Clientes.Idcliente, Clientes.NomCliente, Clientes.DirCliente, Clientes.TelCliente, Clientes.Email FROM Clientes; ", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call activar_controles
End Function
