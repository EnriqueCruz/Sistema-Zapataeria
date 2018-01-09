VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form9 
   BackColor       =   &H8000000B&
   Caption         =   "Form9"
   ClientHeight    =   4950
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12495
   LinkTopic       =   "Form9"
   ScaleHeight     =   4950
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   11040
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   11040
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   11040
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   10335
      Begin MSFlexGridLib.MSFlexGrid MSFGPromo 
         Height          =   2535
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4471
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Promociones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12495
   End
   Begin VB.Menu mnupromCatalogos 
      Caption         =   "Catalogos"
      Begin VB.Menu mnupromEmpleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu mnupromProductos 
         Caption         =   "Productos"
      End
   End
   Begin VB.Menu mnupromVentas 
      Caption         =   "Ventas"
   End
   Begin VB.Menu mnudevoDevoluciones 
      Caption         =   "Devoluciones"
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub validar(iRow As Integer, iCol As Integer)
   
        If MSFGPromo.TextMatrix(iRow, 1) = strUnChecked Then
            MSFGPromo.TextMatrix(iRow, 1) = strChecked
        Else
            MSFGPromo.TextMatrix(iRow, 1) = strUnChecked
        End If
    
End Sub

Private Sub Command1_Click()
    Form10.Caption = "Agregar Promocion"
    Form10.Show vbModal
    
    'Unload Me
End Sub



Private Sub Command2_Click()
    Dim c3 As Integer
Dim b3 As Integer
Dim filr13 As Integer
c3 = 0
'filr1 = 0
    For x = 1 To MSFGPromo.Rows - 1
        MSFGPromo.Row = x
        MSFGPromo.Col = 1
        If MSFGPromo.Text = strChecked Then
            c3 = c3 + 1
           Exit For
        End If
    filr13 = x + 1
    Next
    
    If filr13 = 0 Then
        filr13 = filr13 + 1
    End If
    
    For y = 1 To MSFGPromo.Rows - 1
        MSFGPromo.Row = y
        MSFGPromo.Col = 1
        If MSFGPromo.Text = strChecked Then
            b3 = b3 + 1
        End If
    Next
    
            If c3 = 1 And b3 = 1 Then
                
                MSFGPromo.Row = filr13
                MSFGPromo.Col = 3
                cvepromo = MSFGPromo.Text
                MSFGPromo.Col = 4
                Form10!Text1.Text = MSFGPromo.Text
                MSFGPromo.Col = 5
                Form10!Text2.Text = MSFGPromo.Text
                
                Form10.Caption = "Modificar Promocion"
                Form10.Show vbModal
            Else
                If c3 = 1 And b3 >= 2 Then
                    MsgBox "Solo se puede seleccionar un Producto para modificar", vbInformation, ""
                Else
                    If c3 = 0 And b3 = 0 Then
                        MsgBox "Selecciona un Elemento ", vbInformation, ""
                    End If
            End If
        End If

End Sub

Private Sub Command3_Click()
    Dim d As Integer
    Dim tabl As String
    Dim enc As String
    Dim clave As String
    tabl2 = "Promocion"
    
    For z = 1 To MSFGPromo.Rows - 1
        MSFGPromo.Row = z
        MSFGPromo.Col = 1
        If MSFGPromo.Text = strChecked Then
            d = d + 1
        End If
    Next
    
    If d = 0 Then
        MsgBox "Selecciona un Registro Porfavor", vbInformation, ""
    Else
        For a = 1 To MSFGPromo.Rows - 1
            MSFGPromo.Row = a
            MSFGPromo.Col = 1
            If MSFGPromo.Text = strChecked Then
                MSFGPromo.Col = 3
                clave = MSFGPromo.Text
                Call borrar2(clave, tabl2, myConStr)
            End If
        Next
            MsgBox "Registro(s) Eliminado(s)", vbInformation, ""
    End If
Call datprom
End Sub

Private Sub Form_Load()
    Call Permisos(vpuesto)
    Call datprom
End Sub

Private Sub mnudevoDevoluciones_Click()
    Form6.Show
    Unload Me
End Sub

Private Sub mnupromEmpleados_Click()
    Form7.Show
    Unload Me
End Sub

Private Sub mnupromProductos_Click()
    Form7.Show
    Unload Me
End Sub

Private Sub mnupromVentas_Click()
    Form3.Show
    Unload Me
End Sub
Private Sub MSFGPromo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
            Call validar(MSFGPromo.Row, MSFGPromo.Col)
    End If
End Sub
Private Sub MSFGPromo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    
        If MSFGPromo.MouseRow <> 0 And MSFGPromo.MouseCol <> 0 Then
            Call validar(MSFGPromo.MouseRow, MSFGPromo.MouseCol)
        End If
    
    End If
End Sub
