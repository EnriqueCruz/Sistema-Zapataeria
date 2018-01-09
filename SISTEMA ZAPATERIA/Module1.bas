Attribute VB_Name = "Module1"
Public Const myConStr As String = "Provider=Microsoft.jet.OLEDB.4.0;Data source = C:\SISTEMA ZAPATERIA\codigo\BD_Sistema.mdb"
Public vpuesto As String
Public vtotal As Double
Public filas As Integer
Public vempleado As Integer
Public emple As String
Public cveprod As Integer
Public cveemp As Integer
Public cvepromo As String
Const strChecked = "þ"
Const strUnChecked = "q"

Public Sub dataemp()
    Dim numfil As Integer
    
    sqlfil = "select count(Id_Empleado) As NumRow from Empleado"
    Set fil = consulta(sqlfil, myConStr)
    numfil = fil!NumRow + 1
    
    Form7!MSFGEmp.Cols = 12
    Form7!MSFGEmp.Rows = numfil
    Form7!MSFGEmp.ColWidth(0) = 15
    Form7!MSFGEmp.Col = 2
    Form7!MSFGEmp.Row = 0
    Form7!MSFGEmp.Text = "Id_Empleado"
    Form7!MSFGEmp.Col = 3
    Form7!MSFGEmp.ColWidth(2) = 1000
    Form7!MSFGEmp.Text = "Usuario"
    Form7!MSFGEmp.Col = 4
    Form7!MSFGEmp.Text = "Password"
    Form7!MSFGEmp.Col = 5
    Form7!MSFGEmp.ColWidth(5) = 2000
    Form7!MSFGEmp.Text = "Nombre"
    Form7!MSFGEmp.Col = 6
    Form7!MSFGEmp.Text = "RFC"
    Form7!MSFGEmp.Col = 7
    Form7!MSFGEmp.Text = "Direccion"
    Form7!MSFGEmp.Col = 8
    Form7!MSFGEmp.Text = "Telefono"
    Form7!MSFGEmp.Col = 9
    Form7!MSFGEmp.Text = "Celular"
    Form7!MSFGEmp.Col = 10
    Form7!MSFGEmp.Text = "Correo"
    Form7!MSFGEmp.Col = 11
    Form7!MSFGEmp.ColWidth(11) = 1500
    Form7!MSFGEmp.Text = "Puesto"

    SQL = "SELECT  Empleado.id_Empleado,Usuario,Password,Nombre,RFC,Direccion,Telefono,Celular,Correo,Perfil.Tipo from Empleado,Perfil where Empleado.Id_Perfil = Perfil.Id_Perfil order by id_Empleado"
    Set datemp = consulta(SQL, myConStr)
    'myrs.Open SQL, , adOpenForwardOnly, adLockPessimistic
    If datemp.EOF And datemp.BOF Then
           MsgBox "No hay Registros", vbCritical
    Else
        For i = 1 To Form7!MSFGEmp.Rows - 1
                    Form7!MSFGEmp.Row = i
                    Form7!MSFGEmp.Col = 1
                    Form7!MSFGEmp.CellFontName = "Wingdings"
                    Form7!MSFGEmp.CellFontSize = 14
                    Form7!MSFGEmp.CellAlignment = flexAlignCenterCenter
                    Form7!MSFGEmp.Text = strUnChecked
                    Form7!MSFGEmp.Col = 2
                    Form7!MSFGEmp.TextMatrix(i, 2) = datemp!id_empleado
                    Form7!MSFGEmp.Col = 3
                    Form7!MSFGEmp.TextMatrix(i, 3) = datemp!Usuario
                    Form7!MSFGEmp.Col = 4
                    Form7!MSFGEmp.TextMatrix(i, 4) = datemp!Password
                    Form7!MSFGEmp.Col = 5
                    Form7!MSFGEmp.TextMatrix(i, 5) = datemp!Nombre
                    Form7!MSFGEmp.Col = 6
                    Form7!MSFGEmp.TextMatrix(i, 6) = datemp!RFC
                    Form7!MSFGEmp.Col = 7
                    Form7!MSFGEmp.TextMatrix(i, 7) = datemp!direccion
                    Form7!MSFGEmp.Col = 8
                    Form7!MSFGEmp.TextMatrix(i, 8) = datemp!Telefono
                    Form7!MSFGEmp.Col = 9
                    Form7!MSFGEmp.TextMatrix(i, 9) = datemp!Celular
                    Form7!MSFGEmp.Col = 10
                    Form7!MSFGEmp.TextMatrix(i, 10) = datemp!Correo
                    Form7!MSFGEmp.Col = 11
                    Form7!MSFGEmp.TextMatrix(i, 11) = datemp!Tipo
                    datemp.MoveNext
                    If datemp.EOF Then
                        Exit For
                    Else
                    End If
            Next
    End If
End Sub

Public Sub datprod()
    Dim numfil2 As Integer
    
    sqlfil2 = "select count(Id_Producto) As NumRow2 from Producto"
    Set fil2 = consulta(sqlfil2, myConStr)
    numfil2 = fil2!NumRow2 + 1
    
    Form12!MSFGProd.Cols = 13
    Form12!MSFGProd.Rows = numfil2
    Form12!MSFGProd.ColWidth(0) = 15
    Form12!MSFGProd.Col = 2
    Form12!MSFGProd.Row = 0
    Form12!MSFGProd.Text = "Id_Producto"
    Form12!MSFGProd.Col = 3
    Form12!MSFGProd.ColWidth(2) = 1000
    Form12!MSFGProd.Text = "Descripcion"
    Form12!MSFGProd.Col = 4
    Form12!MSFGProd.Text = "Material"
    Form12!MSFGProd.Col = 5
    Form12!MSFGProd.ColWidth(5) = 2000
    Form12!MSFGProd.Text = "Marca"
    Form12!MSFGProd.Col = 6
    Form12!MSFGProd.Text = "Talla"
    Form12!MSFGProd.Col = 7
    Form12!MSFGProd.Text = "Precio"
    Form12!MSFGProd.Col = 8
    Form12!MSFGProd.Text = "Modelo"
    Form12!MSFGProd.Col = 9
    Form12!MSFGProd.Text = "Genero"
    Form12!MSFGProd.Col = 10
    Form12!MSFGProd.Text = "Color"
    Form12!MSFGProd.Col = 11
    Form12!MSFGProd.Text = "Disponibilidad"
    Form12!MSFGProd.Col = 12
    Form12!MSFGProd.Text = "Tipo"
    
    sqlp = "SELECT * from Producto order by Id_Producto"
    Set datp = consulta(sqlp, myConStr)
    'myrs.Open SQL, , adOpenForwardOnly, adLockPessimistic
    If datp.EOF And datp.BOF Then
           MsgBox "No hay Registros", vbCritical
    Else
        For i = 1 To Form12!MSFGProd.Rows - 1
                    Form12!MSFGProd.Row = i
                    Form12!MSFGProd.Col = 1
                    Form12!MSFGProd.CellFontName = "Wingdings"
                    Form12!MSFGProd.CellFontSize = 14
                    Form12!MSFGProd.CellAlignment = flexAlignCenterCenter
                    Form12!MSFGProd.Text = strUnChecked
                    Form12!MSFGProd.Col = 2
                    Form12!MSFGProd.TextMatrix(i, 2) = datp!Id_Producto
                    Form12!MSFGProd.Col = 3
                    Form12!MSFGProd.TextMatrix(i, 3) = datp!Descripcion
                    Form12!MSFGProd.Col = 4
                    Form12!MSFGProd.TextMatrix(i, 4) = datp!Material
                    Form12!MSFGProd.Col = 5
                    Form12!MSFGProd.TextMatrix(i, 5) = datp!Marca
                    Form12!MSFGProd.Col = 6
                    Form12!MSFGProd.TextMatrix(i, 6) = datp!Talla
                    Form12!MSFGProd.Col = 7
                    Form12!MSFGProd.TextMatrix(i, 7) = datp!Precio
                    Form12!MSFGProd.Col = 8
                    Form12!MSFGProd.TextMatrix(i, 8) = datp!modelo
                    Form12!MSFGProd.Col = 9
                    Form12!MSFGProd.TextMatrix(i, 9) = datp!Genero
                    Form12!MSFGProd.Col = 10
                    Form12!MSFGProd.TextMatrix(i, 10) = datp!Color
                    Form12!MSFGProd.Col = 11
                    Form12!MSFGProd.TextMatrix(i, 11) = datp!Disponibilidad
                    Form12!MSFGProd.Col = 12
                    Form12!MSFGProd.TextMatrix(i, 12) = datp!Tipo
                    datp.MoveNext
                    If datp.EOF Then
                        Exit For
                    Else
                    End If
            Next
    End If
End Sub

Public Sub datprom()
    Dim numfil As Integer
    
    sqlfil3 = "select count(Id_Producto) As NumRow3 from Promocion"
    Set fil3 = consulta(sqlfil3, myConStr)
    numfil3 = fil3!NumRow3 + 1
    Form9!MSFGPromo.Cols = 6
    Form9!MSFGPromo.Rows = numfil3
    Form9!MSFGPromo.ColWidth(0) = 15
    Form9!MSFGPromo.Col = 2
    Form9!MSFGPromo.Row = 0
    Form9!MSFGPromo.Text = "Descripcion"
    Form9!MSFGPromo.Col = 3
    Form9!MSFGPromo.ColWidth(2) = 1000
    Form9!MSFGPromo.Text = "Modelo"
    Form9!MSFGPromo.Col = 4
    Form9!MSFGPromo.Text = "Descripcion de Promocion"
    Form9!MSFGPromo.ColWidth(4) = 2000
    Form9!MSFGPromo.Col = 5
    Form9!MSFGPromo.ColWidth(5) = 2000
    Form9!MSFGPromo.Text = "Descuento"
    
    SQL = "select * from Promocion   INNER JOIN  Producto on(Promocion.Id_Producto = Producto.Id_Producto);"
    Set datpromo = consulta(SQL, myConStr)
    'myrs.Open SQL, , adOpenForwardOnly, adLockPessimistic
    If datpromo.EOF And datpromo.BOF Then
           MsgBox "No hay Registros", vbCritical
    Else
        For i = 1 To Form9!MSFGPromo.Rows - 1
                    Form9!MSFGPromo.Row = i
                    Form9!MSFGPromo.Col = 1
                    Form9!MSFGPromo.CellFontName = "Wingdings"
                    Form9!MSFGPromo.CellFontSize = 14
                    Form9!MSFGPromo.CellAlignment = flexAlignCenterCenter
                    Form9!MSFGPromo.Text = strUnChecked
                    Form9!MSFGPromo.Col = 2
                    Form9!MSFGPromo.TextMatrix(i, 2) = datpromo("Producto.Descripcion")
                    Form9!MSFGPromo.Col = 3
                    Form9!MSFGPromo.TextMatrix(i, 3) = datpromo!modelo
                    Form9!MSFGPromo.Col = 4
                    Form9!MSFGPromo.TextMatrix(i, 4) = datpromo("Promocion.Descripcion")
                    Form9!MSFGPromo.Col = 5
                    Form9!MSFGPromo.TextMatrix(i, 5) = datpromo!descuento
                    datpromo.MoveNext
                    If datpromo.EOF Then
                        Exit For
                    Else
                    End If
            Next
    End If
End Sub


Public Sub mm()
    vtotal = 0
    Form3!MSFG1.Cols = 10
    Form3!MSFG1.Rows = 10
    Form3!MSFG1.ColWidth(0) = 15
    Form3!MSFG1.Col = 1
    Form3!MSFG1.Row = 0
    Form3!MSFG1.Text = "No Serie"
    Form3!MSFG1.Col = 2
    Form3!MSFG1.ColWidth(2) = 2500
    Form3!MSFG1.Text = "Descripcion"
    Form3!MSFG1.Col = 3
    Form3!MSFG1.Text = "Tipo"
    Form3!MSFG1.Col = 4
    Form3!MSFG1.Text = "Color"
    Form3!MSFG1.Col = 5
    Form3!MSFG1.Text = "Talla"
    Form3!MSFG1.Col = 6
    Form3!MSFG1.Text = "Cantidad"
    Form3!MSFG1.Col = 7
    Form3!MSFG1.Text = "Precio"
    Form3!MSFG1.Col = 8
    Form3!MSFG1.Text = "Promocion"
    Form3!MSFG1.Col = 9
    Form3!MSFG1.Text = "Subtotal"
End Sub


Public Sub limpia()
    Form6!MSFG2.Rows = 10
    Form6!MSFG2.Cols = 10
    
    Form6!MSFG2.AllowUserResizing = flexResizeBoth
    Form6!MSFG2.ColWidth(0) = 15
    Form6!MSFG2.Row = 0
    Form6!MSFG2.Col = 2
    Form6!MSFG2.Text = "No Serie"
    Form6!MSFG2.Col = 3
    Form6!MSFG2.ColWidth(2) = 2500
    Form6!MSFG2.Text = "Descripcion"
    Form6!MSFG2.Col = 4
    Form6!MSFG2.Text = "Tipo"
    Form6!MSFG2.Col = 5
    Form6!MSFG2.Text = "Color"
    Form6!MSFG2.Col = 6
    Form6!MSFG2.Text = "Talla"
    Form6!MSFG2.Col = 7
    Form6!MSFG2.Text = "Id_Ventas"
    Form6!MSFG2.Col = 8
    Form6!MSFG2.Text = "Precio"
    Form6!MSFG2.Col = 9
    Form6!MSFG2.Text = "Promocion"
    
    For y = 1 To Form6!MSFG2.Rows - 1
            Form6!MSFG2.Row = y
            Form6!MSFG2.Col = 1
            Form6!MSFG2.CellFontName = "Wingdings"
            Form6!MSFG2.CellFontSize = 14
            Form6!MSFG2.CellAlignment = flexAlignCenterCenter
            Form6!MSFG2.Text = strUnChecked
    Next y
End Sub

Public Sub catempleado()
    MSFGEmp.Cols = 12
    MSFGEmp.Rows = 10
    MSFGEmp.ColWidth(0) = 15
    MSFGEmp.Col = 2
    MSFGEmp.Row = 0
    MSFGEmp.Text = "Id_Empleado"
    MSFGEmp.Col = 3
    MSFGEmp.ColWidth(2) = 1000
    MSFGEmp.Text = "Usuario"
    MSFGEmp.Col = 4
    MSFGEmp.Text = "Password"
    MSFGEmp.Col = 5
    MSFGEmp.Text = "Nombre"
    MSFGEmp.Col = 6
    MSFGEmp.Text = "RFC"
    MSFGEmp.Col = 7
    MSFGEmp.Text = "Direccion"
    MSFGEmp.Col = 8
    MSFGEmp.Text = "Telefono"
    MSFGEmp.Col = 9
    MSFGEmp.Text = "Celular"
    MSFGEmp.Col = 10
    MSFGEmp.Text = "Correo"
    MSFGEmp.Col = 11
    MSFGEmp.Text = "Puesto"
    
    
    SQL = "SELECT  Empleado.id_Empleado,Usuario,Password,Nombre,RFC,Direccion,Telefono,Celular,Correo,Perfil.Tipo from Empleado,Perfil where Empleado.Id_Perfil = Perfil.Id_Perfil"
    Set datemp = consulta(SQL, myConStr)
    'myrs.Open SQL, , adOpenForwardOnly, adLockPessimistic
    If datemp.EOF And datemp.BOF Then
           MsgBox "No hay Registros", vbCritical
    Else
        For i = 1 To MSFGEmp.Rows - 1
                    MSFGEmp.Row = i
                    MSFGEmp.Col = 1
                    MSFGEmp.CellFontName = "Wingdings"
                    MSFGEmp.CellFontSize = 14
                    MSFGEmp.CellAlignment = flexAlignCenterCenter
                    MSFGEmp.Text = strUnChecked
                    MSFGEmp.Col = 2
                    MSFGEmp.TextMatrix(i, 2) = datemp!id_empleado
                    MSFGEmp.Col = 3
                    MSFGEmp.TextMatrix(i, 3) = datemp!Usuario
                    MSFGEmp.Col = 4
                    MSFGEmp.TextMatrix(i, 4) = datemp!Password
                    MSFGEmp.Col = 5
                    MSFGEmp.TextMatrix(i, 5) = datemp!Nombre
                    MSFGEmp.Col = 6
                    MSFGEmp.TextMatrix(i, 6) = datemp!RFC
                    MSFGEmp.Col = 7
                    MSFGEmp.TextMatrix(i, 7) = datemp!direccion
                    MSFGEmp.Col = 8
                    MSFGEmp.TextMatrix(i, 8) = datemp!Telefono
                    MSFGEmp.Col = 9
                    MSFGEmp.TextMatrix(i, 9) = datemp!Celular
                    MSFGEmp.Col = 10
                    MSFGEmp.TextMatrix(i, 10) = datemp!Correo
                    MSFGEmp.Col = 11
                    MSFGEmp.TextMatrix(i, 11) = datemp!Tipo
                    datemp.MoveNext
                    If datemp.EOF Then
                        Exit For
                    Else
                    End If
            Next
    End If
End Sub

Sub Permisos(vpuesto)
    If vpuesto = "Vendedor" Then
        Form3!mnuEmpleados.Enabled = False
        Form6!mnudevEmpleados.Enabled = False
        Form12!mnuprodEmpleados.Enabled = False
        Form9!Command1.Enabled = False
        Form9!Command2.Enabled = False
        Form9!Command3.Enabled = False
        Form12!Command1.Enabled = False
        Form12!Command2.Enabled = False
        Form12!Command3.Enabled = False
        Form3!mnuDevolucion.Enabled = False
        Form9!mnudevoDevoluciones.Enabled = False
        Form9!mnupromEmpleados.Enabled = False
        Form12!mnuprodDevoluciones.Enabled = False
        Form7!Command1.Enabled = False
        Form7!Command2.Enabled = False
        Form7!Command3.Enabled = False
        Form7!mnuDevolucion.Enabled = False
    Else
        If vpuesto = "Gerente" Then
            Form9!Command2.Enabled = True
            Form12!Command2.Enabled = True
        End If
    End If
End Sub
