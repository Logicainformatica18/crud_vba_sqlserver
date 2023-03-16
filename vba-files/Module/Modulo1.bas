Attribute VB_Name = "Modulo1"
'creamos una variable publica
Public miConexion As New ADODB.Connection
'creamos una variable publica del tipo adodb
Public rs As New ADODB.Recordset
Public rsarticulos As New ADODB.Recordset
Public rsfacturas As New ADODB.Recordset
'creamos un procedimiento conecta
Sub Conecta()
'le damos un valor a la variable miconexion la propiedad .connection
Set miConexion = New ADODB.Connection

With miConexion
    'utilizamos la libreria importada (referencias)
    .ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.1.18;Initial Catalog=pcventas;User ID=sa;Password=1234;"
    'abrimos la conexion
    .Open
End With
End Sub

