VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Public idventa As Integer
 
 Function consulta(ByVal qry As Variant, ByVal myconst As Variant) As ADODB.Recordset

    Set mycon = New ADODB.Connection
    Set myrs = New ADODB.Recordset
    mycon.Open myconst
    Set myrs.ActiveConnection = mycon
    myrs.Open qry, , adOpenForwardOnly, adLockPessimistic
    Set consulta = myrs

End Function
Sub guardap(ByVal vpas As String, ByVal vrecor As ADODB.Recordset)
    vrecor!Password = vpas
    vrecor.Update
End Sub


Sub myrsdispo(ByVal vcant As Integer, ByVal myrsrecor As ADODB.Recordset)

    vbcan = myrsrecor!Disponibilidad
    vcant = vbcan - vcant
    myrsrecor!Disponibilidad = vcant
    myrsrecor.Update

End Sub


Public Function myrsventas(ByVal varvidv As Integer, ByVal varvempleado As Integer, ByVal myrsvrecor As ADODB.Recordset) As Integer
    Dim ventas As Integer
    ventas = 0
    
    If myrsvrecor.BOF And myrsvrecor.EOF Then
                vidv = 1
                vord = 1
            Else
                myrsvrecor.MoveLast
                If vord = 0 Then
                    vord = myrsvrecor!orden + 1
                End If
                    vidv = myrsvrecor!id_ventas + 1
            End If
            
            myrsvrecor.AddNew
            myrsvrecor!id_ventas = vidv
            myrsvrecor!orden = vord
            myrsvrecor!fecha = Date
            myrsvrecor!hora = Time
            myrsvrecor!id_empleado = varvempleado
            myrsvrecor!Id_Producto = varvidv
            myrsvrecor.Update
            
            'Set valvid = vidv
            ventas = vidv
            myrsventas = ventas
End Function

    'idventa = vidv

Sub myrspagos(ByVal varvidv As Variant, ByVal varvtipo As Variant, ByVal myrsprecor As ADODB.Recordset)
    If myrsprecor.BOF And myrsprecor.EOF Then
                vidp = 1
            Else
                myrsprecor.MoveLast
                vidp = myrsprecor!id_pagos + 1
            End If
            
            myrsprecor.AddNew
            myrsprecor!id_pagos = vidp
            myrsprecor!Tipo = varvtipo
            myrsprecor!id_venta = Str(varvidv)
            myrsprecor.Update
End Sub



Sub altadevo(ByVal vidv As Integer, ByVal tabventa As ADODB.Recordset) 'As Integer
    
    If tabventa.BOF And tabventa.EOF Then
                        vidd = 1
                    Else
                        tabventa.MoveLast
                        vidd = tabventa!id_devoluciones + 1
                    End If
                    
                    'Form6!MSFG2.Col = 7
                    'vidv = Form6!MSFG2.Text
                    tabventa.AddNew
                    tabventa!id_devoluciones = vidd
                    tabventa!fecha = Date
                    tabventa!hora = Time
                    tabventa!id_ventas = vidv
                    tabventa.Update
                    
                    'altadevo = vidv
    
End Sub


Sub altadevo2(ByVal vidvv As Integer, ByVal devorecor As ADODB.Recordset)
    If devorecor.BOF And devorecor.EOF Then
                        vidd = 1
                    Else
                        devorecor.MoveLast
                        vidd = devorecor!id_devoluciones + 1
                    End If
                    
                    'Form6!MSFG2.Col = 7
                    'vidv = Form6!MSFG2.Text
                    devorecor.AddNew
                    devorecor!id_devoluciones = vidd
                    devorecor!fecha = Date
                    devorecor!hora = Time
                    devorecor!id_ventas = vidvv
                    devorecor.Update

End Sub


'Function agregaemp(ByVal consulta As ADODB.Recordset, ByVal nomusu As String, ByVal nomtrab As String, ByVal varrfc As String, ByVal dir As String, ByVal tel As Integer, ByVal cel As Integer, ByVal email As String) As Integer
 '   Dim puesto As String
  '  puesto = ""
   ' If consulta.BOF And consulta.EOF Then
    '        vid = 1
     '   Else
      '      consulta.MoveLast
       '     vid = consulta!id_empleado + 1
        'End If
         '   consulta.AddNew
          '  consulta!id_empleado = vid
           ' consulta!Usuario = nomusu
            'consulta!Password = puesto
            'consulta!Nombre = nomtrab
            'consulta!RFC = varrfc
            'consulta!Direccion = dir
            'consulta!Telefono = tel
            'consulta!Celular = cel
            'consulta!Correo = email
            'consulta!Id_Perfil = vid
            'consulta.Update
            
        'agregaemp = vid
'End Function



'Sub agregaperf(ByVal perfrecor As ADODB.Recordset, ByVal tipopuesto As String, ByVal idperf As Integer)
 '
  '          If perfrecor.BOF And perfrecor.EOF Then
   '         vid = 1
    '        Else
     '               perfrecor.MoveLast
      '              vid = perfrecor!Id_Perfil + 1
       '     End If
        '            perfrecor.AddNew
         '           perfrecor!Tipo = tipopuesto
          '          perfrecor!Id_Perfil = idperf
           '         perfrecor.Update
'
'End Sub





Sub agregaemp(ByVal vidp As Integer, ByVal emp As String, ByVal conexion As String, ByVal nomtrab As String, ByVal varrfc As String, ByVal dir As String, ByVal tel As Variant, ByVal cel As Variant, ByVal email As String)
    Dim puesto As String
    Set mycon = New ADODB.Connection
    Set myr = New ADODB.Recordset
    mycon.Open conexion
    Set myr.ActiveConnection = mycon
    SQL = "select * from empleado"
    myr.Open SQL, , adOpenForwardOnly, adLockPessimistic
  
    If myr.BOF And myr.EOF Then
            vid = 1
    Else
            myr.MoveLast
           vid = myr!id_empleado + 1
    End If
            myr.AddNew
            myr!id_empleado = vid
            myr!Usuario = emp
            myr!Password = puesto
            myr!Nombre = nomtrab
            myr!RFC = varrfc
            myr!Direccion = dir
            myr!Telefono = tel
            myr!Celular = cel
            myr!Correo = email
            myr!Id_Perfil = vidp
            myr.Update
            myr = Null
End Sub



Sub agregaPromo(ByVal conexion3 As String, ByVal descr As String, ByVal descuento As Variant)
    Set mycon = New ADODB.Connection
    Set myrpro = New ADODB.Recordset
    mycon.Open conexion3
    Set myrpro.ActiveConnection = mycon
    SQL = "select * from Promocion"
    myrpro.Open SQL, , adOpenForwardOnly, adLockPessimistic
  
    If myrpro.BOF And myrpro.EOF Then
            vid = 1
    Else
            myrpro.MoveLast
            vid = myrpro!id_Promocion + 1
    End If
            myrpro.AddNew
            myrpro!id_Promocion = vid
            myrpro!Descripcion = descr
            myrpro!descuento = descuento
            myrpro!Id_Producto = vid
            myrpro.Update
            myrpro = Null
End Sub



Sub agregaprod(ByVal conexion2 As String, ByVal desc As String, ByVal mat As String, ByVal marc As String, ByVal tall As Variant, ByVal pre As Variant, ByVal mode As String, ByVal gen As String, ByVal disp As Variant, ByVal tip As String, ByVal col As String)
    'Dim puesto As String
    Set mycon = New ADODB.Connection
    Set myrp = New ADODB.Recordset
    mycon.Open conexion2
    Set myrp.ActiveConnection = mycon
    SQL = "select * from Producto"
    myrp.Open SQL, , adOpenForwardOnly, adLockPessimistic
  
    If myrp.BOF And myrp.EOF Then
            vid = 1
    Else
            myrp.MoveLast
            vid = myrp!Id_Producto + 1
    End If
            myrp.AddNew
            myrp!Id_Producto = vid
            myrp!Descripcion = desc
            myrp!Material = mat
            myrp!Marca = marc
            myrp!Talla = tall
            myrp!Precio = pre
            myrp!modelo = mode
            myrp!Genero = gen
            myrp!Color = col
            myrp!Disponibilidad = disp
            myrp!Tipo = tip
            myrp.Update
            myrp = Null
End Sub

Sub modemp(ByVal modvidp As Integer, ByVal modemp As String, ByVal modconexion As ADODB.Recordset, ByVal modnomtrab As String, ByVal modvarrfc As String, ByVal moddir As String, ByVal modtel As Variant, ByVal modcel As Variant, ByVal modemail As String)
                modconexion!id_empleado = modconexion!id_empleado
                modconexion!Usuario = modemp
                'modconexion!Password = modconexion!Password
                modconexion!Nombre = modnomtrab
                modconexion!RFC = modvarrfc
                modconexion!Direccion = moddir
                modconexion!Telefono = modtel
                modconexion!Celular = modcel
                modconexion!Correo = modemail
                modconexion!Id_Perfil = modvidp
                modconexion.Update
                'modconexion = Null
End Sub

Sub modpro(ByVal prov2 As String, ByVal moddesc As String, ByVal modmate As String, ByVal modmarca As String, ByVal modtalla As Integer, ByVal modpre As Integer, ByVal mmodelo As String, ByVal modgen As String, ByVal modcol As String, ByVal moddisp As Integer, ByVal modtip As String, ByVal cveemp As Integer)
                Set mycon = New ADODB.Connection
                Set modconexion2 = New ADODB.Recordset
                mycon.Open prov2
                Set modconexion2.ActiveConnection = mycon
                SQLpromod = "select * from Producto  Where Producto.Id_Producto = " + Str(cveemp)
                modconexion2.Open SQLpromod, , adOpenForwardOnly, adLockPessimistic
                'modconexion2!id_Promocion = modconexion2!id_Promocion
                'modconexion2("Promocion.Id_Producto") = cveemp
                'modconexion2!descuento = modconexion2!descuento
                'modconexion2("Promocion.Descripcion") = modconexion2("Promocion.Descripcion")
                'modconexion2("Producto.Id_Producto") = cveemp
                modconexion2("Descripcion") = moddesc
                modconexion2("Material") = modmate
                modconexion2("Marca") = modmarca
                modconexion2("Talla") = modtalla
                modconexion2("Precio") = modpre
                modconexion2("Modelo") = mmodelo
                modconexion2("Genero") = modgen
                modconexion2("Color") = modcol
                modconexion2("Disponibilidad") = moddisp
                modconexion2("Tipo") = modtip
                modconexion2.Update
                'modconexion = Null
End Sub
Sub modpromo(ByVal prov3 As String, ByVal promdesc As String, ByVal promde As Integer, ByVal modelo As String)
                Set mycon = New ADODB.Connection
                Set modconexion3 = New ADODB.Recordset
                mycon.Open prov3
                Set modconexion3.ActiveConnection = mycon
                SQLpromod = "select * from Promocion   INNER JOIN  Producto on(Promocion.Id_Producto = Producto.Id_Producto) where Modelo = '" + cvepromo + "'"
                modconexion3.Open SQLpromod, , adOpenForwardOnly, adLockPessimistic
                modconexion3!id_Promocion = modconexion3!id_Promocion
                modconexion3("Promocion.Descripcion") = promdesc
                modconexion3!descuento = promde
                modconexion3.Update
End Sub

Public Sub borrar(ByVal vid As Integer, ByVal tala As String, ByVal myconst As String)
    Set mycon = New ADODB.Connection
    Set myrs = New ADODB.Recordset
    mycon.Open myconst
    SQL = "delete from " + Trim(tala) + " where id_" + Trim(tala) + "=" + Str(vid)
    Set myrs.ActiveConnection = mycon
    myrs.Open SQL, , adOpenForwardOnly, adLockPessimistic
End Sub
Public Sub borrar2(ByVal model As String, ByVal talapromo As String, ByVal myconst As String)
    Dim idprom As Integer
    
    Set mycon = New ADODB.Connection
    mycon.Open myconst
    Set myrsmod = New ADODB.Recordset
    sqlmod = "select Id_Producto from Producto where Modelo = '" + Trim(model) + "'"
    Set myrsmod.ActiveConnection = mycon
    myrsmod.Open sqlmod, , adOpenForwardOnly, adLockPessimistic
    idprom = Int(myrsmod!Id_Producto)
    Set myrs = New ADODB.Recordset
    Set myrs.ActiveConnection = mycon
    'mycon.Open myconst
    SQL = "delete from " + Trim(talapromo) + " where id_" + Trim(talapromo) + "=" + Str(idprom)
    myrs.Open SQL, , adOpenForwardOnly, adLockPessimistic
End Sub

