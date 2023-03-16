VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Articulos 
   Caption         =   "UserForm4"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7155
   OleObjectBlob   =   "Articulos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnanterior_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'mueve al registro
rsarticulos.MovePrevious
    txtidarticulo.Text = rsarticulos.Fields("IdArticulo")
    txtnombre.Text = rsarticulos.Fields("NomArticulo")
    txtprecio.Text = Val(rsarticulos.Fields("PreArticulo"))
    txtstock.Text = Val(rsarticulos.Fields("Stock"))
     cbcategorias.Text = rsarticulos.Fields("codigo_nombre")
End Sub

Private Sub btnbuscar_Click()
'ejecutar el procedimiento conecta
Call Conecta
Set rs_search = New ADODB.Recordset
'ejecutamos una consulta a la tabla  categorias de la base de datos pcventas
'SELECT categorias.categoria, categorias.nombre FROM categorias where categorias.nombre Like "*"
rs_search.Open "SELECT a.IdArticulo,a.NomArticulo,a.PreArticulo,a.Stock,c.nombre from articulos a inner join categorias c on a.categ=c.categoria where nomarticulo Like '" & txtcriterio.Text & "%'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
'on error resume next ejecuta si no hay error
On Error Resume Next
'asignamos los numeros de columnas
With Me.lsarticulos
    .ColumnCount = rs_search.Fields.Count
    End With
    'movernos al inicio
    rs_search.MoveFirst
    Dim i As Integer
    i = 1
    With Me.lsarticulos
    .Clear
    'a�adir los encabezados
    .AddItem
    'aca van los numeros de columnas (la columna uno es el indice cero,la segunda el indice uno y asi sucesivamente)
    For j = 0 To 4
    .List(0, j) = rs_search.Fields(j).Name
    Next j
    'llenado de los registros de la tabla categorias al listbox (cuadro de lista)
    Do
    .AddItem
    .List(i, 0) = rs_search![IdArticulo]
    .List(i, 1) = rs_search![NomArticulo]
    .List(i, 2) = rs_search![PreArticulo]
    .List(i, 3) = rs_search![Stock]
    .List(i, 4) = rs_search![nombre]
    i = i + 1
    'avansamos un registro siguiente
    rs_search.MoveNext
    Loop Until rs_search.EOF

End With
'cerramos la conexion
miConexion.Close
End Sub

Private Sub btndelete_Click()

End Sub

Private Sub btnguardar_Click()
Set rsarticulos_save = New ADODB.Recordset
  rsarticulos_save.Open "SELECT * from Articulos", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
'on error resume next encuentra error
On Error Resume Next
'agregar nuevo registro
rsarticulos_save.AddNew
'Obtener llave secundaria
Dim cod_categorias As String
cod_categorias = cbcategorias.Text
cod_categorias = Mid(cod_categorias, 1, InStr(1, cod_categorias, "-") - 1)
'asignamos los valores de las cajas de texto a los campos del registro
rsarticulos_save.Fields("IdArticulo") = txtidarticulo.Text
rsarticulos_save.Fields("NomArticulo") = txtnombre.Text
rsarticulos_save.Fields("PreArticulo") = Val(txtprecio.Text)
rsarticulos_save.Fields("Stock") = Val(txtstock.Text)
rsarticulos_save.Fields("Categ") = cod_categorias
'guardar registro
rsarticulos_save.Save
'mostramos un mensaje de exito
MsgBox ("Datos Guardados correctamente")
'actualizar registro
  Call reset

'bloquear controles
btnanterior.Enabled = True
btnsiguiente.Enabled = True
btnprimero.Enabled = True
btnultimo.Enabled = True
End Sub

Private Sub btnnuevo_Click()
'limpiamos las cajas de texto
txtidarticulo.Text = ""
txtnombre.Text = ""
txtprecio.Text = ""
txtstock.Text = ""


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
rsarticulos.MoveFirst
    txtidarticulo.Text = rsarticulos.Fields("IdArticulo")
    txtnombre.Text = rsarticulos.Fields("NomArticulo")
    txtprecio.Text = Val(rsarticulos.Fields("PreArticulo"))
    txtstock.Text = Val(rsarticulos.Fields("Stock"))
  cbcategorias.Text = rsarticulos.Fields("codigo_nombre")
End Sub

Private Sub btnsiguiente_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'mueve al registro
rsarticulos.MoveNext
    txtidarticulo.Text = rsarticulos.Fields("IdArticulo")
    txtnombre.Text = rsarticulos.Fields("NomArticulo")
    txtprecio.Text = Val(rsarticulos.Fields("PreArticulo"))
    txtstock.Text = Val(rsarticulos.Fields("Stock"))
    cbcategorias.Text = rsarticulos.Fields("codigo_nombre")
End Sub

Private Sub btnultimo_Click()
'on error resume next ejecuta si no hay error
On Error Resume Next
'mueve al registro
rsarticulos.MoveLast
    txtidarticulo.Text = rsarticulos.Fields("IdArticulo")
    txtnombre.Text = rsarticulos.Fields("NomArticulo")
    txtprecio.Text = Val(rsarticulos.Fields("PreArticulo"))
    txtstock.Text = Val(rsarticulos.Fields("Stock"))
  cbcategorias.Text = rsarticulos.Fields("codigo_nombre")
End Sub

Private Sub cbcategorias_Change()

End Sub

Private Sub UserForm_Initialize()
    'ejecutar el procedimiento conecta
Call Conecta


    Call reset
    
    
    Set rs2 = New ADODB.Recordset
    Set rscategorias = New ADODB.Recordset
    rs2.Open "SELECT count(*) from Categorias", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    rscategorias.Open "SELECT categoria & '- '& nombre as codigo_nombre  from Categorias order by Nombre asc;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
     'le damos los items al cuadro combinado
   cbcategorias.AddItem (rscategorias.Fields("codigo_nombre"))
   'iteramos para agregar los demas items al cuadro combinado
    Dim registros As Integer
    registros = Val(rs2.Fields("Expr1000")) - 1
                    
    For x = 1 To registros Step 1
        rscategorias.MoveNext
   cbcategorias.AddItem (rscategorias.Fields("codigo_nombre"))
    Next
End Sub
Public Sub reset()

Set rsarticulos = New ADODB.Recordset
'ejecutamos una consulta a la tabla clientes de la base de datos pcventas
  
  
  rsarticulos.Open "SELECT a.IdArticulo,a.NomArticulo,a.PreArticulo,a.Stock,a.Categ,c.categoria & '- ' & c.nombre as codigo_nombre from articulos a inner join categorias c on a.categ=c.categoria order by a.idarticulo;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  

    txtidarticulo.Text = rsarticulos.Fields("IdArticulo")
    txtnombre.Text = rsarticulos.Fields("NomArticulo")
    txtprecio.Text = Val(rsarticulos.Fields("PreArticulo"))
    txtstock.Text = Val(rsarticulos.Fields("Stock"))
    cbcategorias.Text = rsarticulos.Fields("codigo_nombre")
End Sub
