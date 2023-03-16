VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Facturas 
   Caption         =   "UserForm1"
   ClientHeight    =   5295
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9960.001
   OleObjectBlob   =   "FACTURAS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FACTURAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnanterior_Click()
On Error Resume Next
rsfacturas.MovePrevious
   txtidfactura.Text = rsfacturas.Fields("idfactura")
    txtfecha.Text = rsfacturas.Fields("fecha")
    cbidcliente.Text = rsfacturas.Fields("codigo_nombre")
    Call Detalle
End Sub

Private Sub btnbuscar_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs_search = New ADODB.Recordset
'ejecutamos una consulta a la tabla  categorias de la base de datos pcventas
'SELECT categorias.categoria, categorias.nombre FROM categorias where categorias.nombre Like "*"
rs_search.Open "SELECT F.IDFACTURA,F.FECHA, C.IDCLIENTE, C.NOMCLIENTE from FACTURAS F INNER JOIN CLIENTES C  ON F.IDCLIENTE = C.IDCLIENTE where C.NOMCLIENTE Like '" & txtcriterio.Text & "%'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
'on error resume next ejecuta si no hay error
On Error Resume Next
'asignamos los numeros de columnas
With Me.lsfacturas
    .ColumnCount = rs_search.Fields.Count
    End With
    'movernos al inicio
    rs_search.MoveFirst
    Dim i As Integer
    i = 1
    With Me.lsfacturas
    .Clear
    'a�adir los encabezados
    .AddItem
    'aca van los numeros de columnas (la columna uno es el indice cero,la segunda el indice uno y asi sucesivamente)
    For j = 0 To 3
    .List(0, j) = rs_search.Fields(j).Name
    Next j
    'llenado de los registros de la tabla  al listbox (cuadro de lista)
    Do
    .AddItem
    .List(i, 0) = rs_search![IdFactura]
    .List(i, 1) = rs_search![FECHA]
    .List(i, 2) = rs_search![IDCLIENTE]
    .List(i, 3) = rs_search![NOMCLIENTE]
    
    
    i = i + 1
    'avansamos un registro siguiente
    rs_search.MoveNext
    Loop Until rs_search.EOF

End With
'cerramos la conexion
miConexion.Close
End Sub





Private Sub btnnuevo_Click()
On Error Resume Next
rsfacturas.MoveLast
 Detalles.lblfactura = Val(rsfacturas.Fields("idfactura")) + 1
Detalles.Show
End Sub

Private Sub btnprimero_Click()
On Error Resume Next
rsfacturas.MoveFirst
   txtidfactura.Text = rsfacturas.Fields("idfactura")
    txtfecha.Text = rsfacturas.Fields("fecha")
    cbidcliente.Text = rsfacturas.Fields("codigo_nombre")
    Call Detalle
End Sub

Private Sub btnsiguiente_Click()
On Error Resume Next
rsfacturas.MoveNext
   txtidfactura.Text = rsfacturas.Fields("idfactura")
    txtfecha.Text = rsfacturas.Fields("fecha")
    cbidcliente.Text = rsfacturas.Fields("codigo_nombre")
    Call Detalle
End Sub

Private Sub btnultimo_Click()

On Error Resume Next
rsfacturas.MoveLast
   txtidfactura.Text = rsfacturas.Fields("idfactura")
    txtfecha.Text = rsfacturas.Fields("fecha")
    cbidcliente.Text = rsfacturas.Fields("codigo_nombre")
    Call Detalle
End Sub



Private Sub cbidcliente_Change()

End Sub

Private Sub UserForm_Click()

End Sub
Sub Detalle()

''''''''''''   LISTAR LOS DETALLES''''''''''''''''''''''
     
Set rs_detalle = New ADODB.Recordset
rs_detalle.Open "SELECT A.NOMARTICULO,D.CANTIDAD,D.PREVENTA, D.CANTIDAD * D.PREVENTA AS SUBTOTAL FROM  DETALLES D  INNER JOIN ARTICULOS A ON D.IDARTICULO = A.IDARTICULO WHERE D.IDFACTURA =" & Val(txtidfactura.Text) & ";", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
'on error resume next ejecuta si no hay error
On Error Resume Next
'asignamos los numeros de columnas
With Me.lsdetalle
    .ColumnCount = rs_detalle.Fields.Count
    End With
    'movernos al inicio
    rs_detalle.MoveFirst
    Dim i As Integer
    i = 1
    With Me.lsdetalle
    .Clear
    'a�adir los encabezados
    .AddItem
    'aca van los numeros de columnas (la columna uno es el indice cero,la segunda el indice uno y asi sucesivamente)
    For j = 0 To 3
    .List(0, j) = rs_detalle.Fields(j).Name
    Next j
    'llenado de los registros de la tabla  al listbox (cuadro de lista)
    Do
    .AddItem
    .List(i, 0) = rs_detalle![NomArticulo]
    .List(i, 1) = rs_detalle![Cantidad]
    .List(i, 2) = rs_detalle![PreVenta]
    .List(i, 3) = rs_detalle![Subtotal]


    i = i + 1
    'avansamos un registro siguiente
    rs_detalle.MoveNext
    Loop Until rs_detalle.EOF

End With
    
End Sub
Private Sub UserForm_Initialize()
    'ejecutar el procedimiento conecta
Call Conecta


    Call reset
    Set rs2 = New ADODB.Recordset
    Set rsclientes = New ADODB.Recordset
    rs2.Open "SELECT count(*) from clientes", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    rsclientes.Open "SELECT idcliente & '- '& nomcliente as codigo_nombre  from clientes order by Nomcliente asc;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
     'le damos los items al cuadro combinado
   cbidcliente.AddItem (rsfacturas.Fields("codigo_nombre"))
   'iteramos para agregar los demas items al cuadro combinado
    Dim registros As Integer
    registros = Val(rs2.Fields("Expr1000")) - 1

    For x = 1 To registros Step 1
        rsclientes.MoveNext
   cbidcliente.AddItem (rsclientes.Fields("codigo_nombre"))
    Next
    
    
    Call Detalle
    

    
    
 
    
End Sub
Sub reset()

Set rsfacturas = New ADODB.Recordset
'ejecutamos una consulta a la tabla clientes de la base de datos pcventas
  

  rsfacturas.Open "SELECT F.IDFACTURA,F.FECHA, F.IDCLIENTE,C.IDCLIENTE & '- ' & C.NOMCLIENTE as codigo_nombre  FROM FACTURAS F INNER JOIN CLIENTES C ON F.IDCLIENTE = C.IDCLIENTE order by f.fecha;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

    txtidfactura.Text = rsfacturas.Fields("idfactura")
    txtfecha.Text = rsfacturas.Fields("fecha")
    cbidcliente.Text = rsfacturas.Fields("codigo_nombre")
End Sub
