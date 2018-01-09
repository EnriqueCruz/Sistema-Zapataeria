# Sistema-Zapataeria
Autor: Luis Enrique Magdaleno de la Cruz
Versión: 1.0

Sistema de Ventas de zapatera que realiza las transacciones más importantes (Consultar, Insertar, Actualizar, Eliminar). Entre las características más destacables se encuentran manejo operaciones de ventas, utilización de catálogos para gestionar personal, productos y sus respectivas promociones. Implementa métodos de pago y una opción de devolución del producto.
Es importante destacar que el proyecto es multiusuario por lo cual se otorgan diferentes accesos dependiendo del perfil, dentro de estos perfiles se tiene: Administrador, Vendedor y Gerente. Cada tipo de usuario tiene diferentes accesos que se traducen en las acciones que pueden realizar dentro del sistema.
Se compone de diversos módulos entre los que se encuentran:
1)	Login
2)	Ventas
3)	Pagos
4)	Devoluciones
5)	Catálogo de Empleados
6)	Catálogo de Promociones
7)	Catálogo de Productos

1.	Login: este  módulo consiste en una ventana la cual está destinando a los usuarios que se registren al sistema utilizando su “Usuario” y “Login” (véase la imagen: login), al momento de registrarse el usuario accede a la página principal (Agregar Producto). Por otra parte este módulo cuenta con una ventana (Contraseña Correcta) la cual registra una contraseña a un usuario sin contraseña, en caso de no poner el mismo valor en los campos de texto, se envía un mensaje advirtiendo de este suceso, una vez añadida la contraseña se envía  la página principal. 

2.	Ventas: este módulo permite a los usuarios seleccionar los productos que compraran (Seleccionar Producto), para ello los usuarios deben de colocar un id de producto en el campo “Numero de Serie” si existe el producto se abrirá una ventana en donde se podrán ver las características del producto seleccionado (Agregar_Producto, Agregar Producto), en esta misma ventana se debe seleccionar la cantidad de producto a comprar, una vez seleccionado se podrá visualizar el producto en la pantalla de venta(Productos_Agregados, Productos_Agregados_2).
En la ventana de vetas existen 2 botones: Limpiar y Quitar. La funcionalidad del primer botón (Botón Limpiar) es la de limpiar la tabla donde se muestran los productos agregados, por su parte el botón de quitar (Botón Quitar) solo quita un producto siempre y cuando se seleccione un producto en cualquiera de las columnas.

3.	Pagos: este módulo se compone de una ventana en donde se elige el tipos de pago: Efectivo o Tarjeta (Tipo Pago Tarjeta, Tipo Pago Efectivo), cada uno de ellos con características distintivas, una vez hecho el pago con cualquiera de los 2 pagos se visualizara una ventana indicando que la venta fue realizada y seguido de eso regresa a la página principal.
4.	Devoluciones: esta parte del sistema se encarga de realizar una devolución de dinero al cliente, para ello se debe se poner el número de orden, seguido de esto marcar el checkbox para seleccionar el producto a devolver, después poner el motivo y al final dar en aceptar. Esta acción quita de la base de datos.

5.	Catalogo Empleados: esta ventana (Catalogo Empleados) se encarga de gestionar el personal de la empresa, entre sus características más importantes se encuentra: consultar, agregar nuevos empleados, modificar información de los ya existentes y eliminarlos. 

6.	Catalogo Promociones: al igual que los empleados contiene funciones de consulta, creación. Modificación y eliminación, todo esto centrado en las promociones respecto a los productos.

7.	Catalogo Productos: siguiendo el mismo patrón que los otros catálogos contiene acciones esenciales para crear, consultar, modificar y eliminar productos.

Requisitos

Microsoft Visual Basic 6.0
Windows XP o superior
Microsoft Access 2003

Instrucciones de uso
1.	Instalar Visual Basic 6.0 
2.	Registrar el archivo “primerdll.dll” ubicado en la carpeta “C:\SISTEMA ZAPATERIA\codigo”. Para registrarlo se tiene dos métodos que dependen del sistema(32 y 64 bits)

      32 bits:
      	Abrir la consola de comandos de Windows en modo administrador.

      	Una vez abierto la consola de comando escribir: “cd C:\SISTEMA ZAPATERIA\codigo”

      	Para registrar el archivo .dll escribir y ejecutar el siguiente comando: regsvr32 “C:\SISTEMA ZAPATERIA\codigo\primerdll.dll”.

      64	bits:
      	Copiar el archivo de librería en el directorio:  “C:\WINDOWS\SysWOW64\”
      	Abrir el Símbolo del sistema como administrador y escribes: “cd C:\WINDOWS\SysWOW64\”
      	Ahora registras la librería: “regsvr32.exe C:\WINDOWS\SysWOW64\NOMBRE_ARCHIVO.DLL”

3.	Colocar la carpeta SISTEMA ZAPATERIA de manera que la ruta que como: “C:\SISTEMA ZAPATERIA”
4.	Abrir VB6 como administrador
5.	Buscar la opción “Abrir”, después ”Abrir Proyecto” y buscar el archivo “Proyecto1.vbp” en la carpeta: “C:\SISTEMA ZAPATERIA\codigo”
6.	Una vez abierto el proyecto ir a la opción “Proyecto” después en “Referencias”.
7.	Cuando se desplegué la venta de “Referencias” buscar la referencia: “Microsoft Activex Data Objects 2.5 Library” y también agregar la referencia: “primerdll”, para agregar esta referencia se debe buscar la opción “Examinar” e ir a la ruta: “C:\SISTEMA ZAPATERIA\codigo”, dentro de la carpeta se deberá apreciar el archivo .dll, se deberá seleccionar el archivo y dar en “Abrir” y después en “Aceptar”
8.	Una vez realizado lo anterior ejecutar el proyecto y probarlo.


Contacto
enrique_magdalenocruz@hotmail.com


