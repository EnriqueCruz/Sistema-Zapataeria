VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H000080FF&
   Caption         =   "Form6"
   ClientHeight    =   6165
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10080
   LinkTopic       =   "Form6"
   ScaleHeight     =   6165
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Motivacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      TabIndex        =   4
      Top             =   4200
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "Desagrado"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Defecto"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFG2 
      Height          =   2175
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Numero de Orden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Devoluciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10095
   End
   Begin VB.Menu mnudevCatalogos 
      Caption         =   "Catalogos"
      Begin VB.Menu mnudevEmpleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu mnudevPromociones 
         Caption         =   "Promociones"
      End
      Begin VB.Menu mnudevProducto 
         Caption         =   "Producto"
      End
   End
   Begin VB.Menu mnudevVentas 
      Caption         =   "Ventas"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Private Sub Form_Load()
    MSFG2.Rows = 10
    MSFG2.Cols = 10
    MSFG2.AllowUserResizing = flexResizeBoth
    MSFG2.ColWidth(0) = 15
    
    MSFG2.Row = 0
    MSFG2.Col = 2
    MSFG2.Text = "No Serie"
    MSFG2.Col = 3
    MSFG2.ColWidth(2) = 2500
    MSFG2.Text = "Descripcion"
    MSFG2.Col = 4
    MSFG2.Text = "Tipo"
    MSFG2.Col = 5
    MSFG2.Text = "Color"
    MSFG2.Col = 6
    MSFG2.Text = "Talla"
    MSFG2.Col = 7
    MSFG2.Text = "Id_Ventas"
    MSFG2.Col = 8
    MSFG2.Text = "Precio"
    MSFG2.Col = 9
    MSFG2.Text = "Promocion"

End Sub

Private Sub validar(iRow As Integer, iCol As Integer)
   
        If MSFG2.TextMatrix(iRow, 1) = strUnChecked Then
            MSFG2.TextMatrix(iRow, 1) = strChecked
        Else
            MSFG2.TextMatrix(iRow, 1) = strUnChecked
        End If
    
End Sub

Private Sub mnudevEmpleados_Click()
    Form7.Show
    Unload Me
End Sub

Private Sub mnudevProducto_Click()
    Form12.Show
    Unload Me
End Sub

Private Sub mnudevPromociones_Click()
    Form9.Show
    Unload Me
End Sub

Private Sub mnudevVentas_Click()
    Form3.Show
    Unload Me
End Sub

Private Sub MSFG2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
            Call validar(MSFG2.Row, MSFG2.Col)
    End If
End Sub

Private Sub MSFG2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    
        If MSFG2.MouseRow <> 0 And MSFG2.MouseCol <> 0 Then
            Call validar(MSFG2.MouseRow, MSFG2.MouseCol)
        End If
    
    End If
End Sub




Private Sub Text1_LostFocus()
    Dim vorden As Integer
    If Text1.Text = "" Then
        MsgBox "Ingresa un Numero de Orden", vbInformation
    Else
       
        SQL = "SELECT Producto.Id_Producto, Producto.Descripcion, Producto.Tipo, Producto.Color, Producto.Descripcion, Producto.Talla, Ventas.Id_Ventas, Producto.Precio, Promocion.Descuento FROM producto, ventas, promocion WHERE Producto.Id_Producto = Ventas.Id_Producto and Producto.Id_Producto = Promocion.Id_Producto  and Ventas.Orden = " + Str(Text1.Text)
        Set info = consulta(SQL, myConStr)
        
        If info.EOF And info.BOF Then
                MsgBox "El numero de serie no es valido", vbCritical, ""
                Text1.Text = ""
                Text1.SetFocus
        Else
        
        MSFG2.Clear
        Call Form_Load
                For i = 1 To MSFG2.Rows - 1
                    MSFG2.Row = i
                    MSFG2.Col = 1
                    MSFG2.CellFontName = "Wingdings"
                    MSFG2.CellFontSize = 14
                    MSFG2.CellAlignment = flexAlignCenterCenter
                    MSFG2.Text = strUnChecked
                    If Trim(MSFG2.Text) <> "" Then
                        MSFG2.TextMatrix(i, 1) = strUnChecked
                        'MSFG2.Col = 2
                        MSFG2.TextMatrix(i, 2) = info!Id_Producto
                        'MSFG2.Col = 3
                        MSFG2.TextMatrix(i, 3) = info!Descripcion
                        'MSFG2.Col = 4
                        MSFG2.TextMatrix(i, 4) = info!Tipo
                        'MSFG2.Col = 5
                        MSFG2.TextMatrix(i, 5) = info!Color
                        'MSFG2.Col = 6
                        MSFG2.TextMatrix(i, 6) = info!Talla
                        'MSFG2.Col = 7
                        MSFG2.TextMatrix(i, 7) = info!id_ventas
                        'MSFG2.Col = 8
                        MSFG2.TextMatrix(i, 8) = info!Precio
                        'MSFG2.Col = 9
                        MSFG2.TextMatrix(i, 9) = info!descuento
                        info.MoveNext
                        If info.EOF Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next i
        End If
    End If
End Sub


Private Sub Command1_Click()
    Dim vidv As Integer
   Dim vidd As Integer
    
    
    For r = 1 To MSFG2.Rows - 1
        MSFG2.Row = r
        MSFG2.Col = 1
        
        If MSFG2.Text <> "" Then
            If MSFG2.Text = strChecked Then
                If Option2.Value = True Then
                
                    MSFG2.Col = 2
                    vid = MSFG2.Text
                    SQL = "select * from producto where id_producto = " + Str(vid)
                    Set info = consulta(SQL, myConStr)
                    'myrs.Open SQL, , adOpenForwardOnly, adLockPessimistic
                    vcan = info!Disponibilidad
                    vcan = vcan + 1
                    info!Disponibilidad = vcan
                    info.Update
                    
                    MSFG2.Col = 7
                    vidv = MSFG2.Text
                    sqld = "select * from devoluciones"
                    Set info2 = consulta(sqld, myConStr)
                    'myrsd.Open sqld, , adOpenForwardOnly, adLockPessimistic
                    Call altadevo(vidv, info2)
                    
                    info = Null
                    info2 = Null
                Else
                    
                    MSFG2.Col = 7
                    vidv = MSFG2.Text
                    sqld = "select * from devoluciones"
                    Set info3 = consulta(sqld, myConStr)
                    Call altadevo2(vidv, info3)
                    'myrsd.Open sqld, , adOpenForwardOnly, adLockPessimistic
                    
                    info3 = Null
                End If
                
                MsgBox "La devolucion se realizo exitosamente", vbInformation, ""
                Text1.Text = ""
                MSFG2.Clear
                Call Form_Load
            Else
                MsgBox "Ningun Producto seleccionado", vbExclamation, ""
            End If
        Else
            Exit For
        End If
    Next
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
