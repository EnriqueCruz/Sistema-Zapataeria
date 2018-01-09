VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   BackColor       =   &H8000000B&
   Caption         =   "Form7"
   ClientHeight    =   5460
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14205
   LinkTopic       =   "Form7"
   ScaleHeight     =   5460
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   12495
      Begin MSFlexGridLib.MSFlexGrid MSFGEmp 
         Height          =   3255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   5741
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   12960
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   12960
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   12960
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Empleados"
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
      Top             =   240
      Width           =   14175
   End
   Begin VB.Menu mnuCatalogo 
      Caption         =   "Catalogo"
      Begin VB.Menu mnuempPromociones 
         Caption         =   "Promociones"
      End
      Begin VB.Menu mnuempProductos 
         Caption         =   "Productos"
      End
   End
   Begin VB.Menu mnuVentas 
      Caption         =   "Ventas"
   End
   Begin VB.Menu mnuDevolucion 
      Caption         =   "Devolucion"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub Command2_Click()
Dim c As Integer
Dim b As Integer
Dim filr1 As Integer
c = 0
filr1 = 0
    For x = 1 To MSFGEmp.Rows - 1
        MSFGEmp.Row = x
        MSFGEmp.Col = 1
        If MSFGEmp.Text = strChecked Then
            c = c + 1
           Exit For
        End If
    filr1 = x + 1
    Next
    
     If filr1 = 0 Then
        filr1 = filr1 + 1
     End If
    
    For y = 1 To MSFGEmp.Rows - 1
        MSFGEmp.Row = y
        MSFGEmp.Col = 1
        If MSFGEmp.Text = strChecked Then
            b = b + 1
        End If
    Next
    
            If c = 1 And b = 1 Then
                
                MSFGEmp.Row = filr1
                MSFGEmp.Col = 2
                cveprod = Int(MSFGEmp.Text)
                MSFGEmp.Col = 3
                Form8!Text1.Text = MSFGEmp.Text
                MSFGEmp.Col = 5
                Form8!Text2.Text = MSFGEmp.Text
                MSFGEmp.Col = 6
                Form8!Text3.Text = MSFGEmp.Text
                MSFGEmp.Col = 7
                Form8!Text4.Text = MSFGEmp.Text
                MSFGEmp.Col = 8
                Form8!Texttel.Text = MSFGEmp.Text
                MSFGEmp.Col = 9
                Form8!Textcel.Text = MSFGEmp.Text
                MSFGEmp.Col = 10
                Form8!Text7.Text = MSFGEmp.Text
                MSFGEmp.Col = 11
                Form8!Combo1.Text = MSFGEmp.Text
                Form8.Caption = "Modificacion"
                Form8.Show vbModal
            Else
                If c = 1 And b >= 2 Then
                    MsgBox "Solo se puede seleccionar un Producto para modificar", vbInformation, ""
                Else
                    If c = 0 And b = 0 Then
                        MsgBox "Selecciona un Elemento", vbInformation, ""
                    End If
            End If
        End If
End Sub

Private Sub Command3_Click()
    Dim d As Integer
    Dim tabl As String
    Dim enc As String
    Dim clave As Integer
    
    tabl = "Empleado"
    enc = "Producto"
    
    For z = 1 To MSFGEmp.Rows - 1
        MSFGEmp.Row = z
        MSFGEmp.Col = 1
        If MSFGEmp.Text = strChecked Then
            d = d + 1
        End If
    Next
    
    
    If d = 0 Then
    MsgBox "Selecciona un Elemento", vbInformation, ""
    
    
    Else
        For a = 1 To MSFGEmp.Rows - 1
            MSFGEmp.Row = a
            MSFGEmp.Col = 1
            If MSFGEmp.Text = strChecked Then
                MSFGEmp.Col = 2
                clave = Int(MSFGEmp.Text)
                Call borrar(clave, tabl, myConStr)
            End If
        Next
            
            MsgBox "Registro(s) Eliminado(s)", vbInformation, ""
    End If
Call dataemp
End Sub

Private Sub Form_Load()
    Call Permisos(vpuesto)
    Call dataemp
End Sub


Private Sub validar(iRow As Integer, iCol As Integer)
   
        If MSFGEmp.TextMatrix(iRow, 1) = strUnChecked Then
            MSFGEmp.TextMatrix(iRow, 1) = strChecked
        Else
            MSFGEmp.TextMatrix(iRow, 1) = strUnChecked
        End If
    
End Sub

Private Sub Command1_Click()
    Form8.Caption = "Alta"
    Form8.Show vbModal
    
End Sub

Private Sub mnuDevolucion_Click()
    Form6.Show
    Unload Me
End Sub

Private Sub mnuempProductos_Click()
    Form12.Show
    Unload Me
End Sub

Private Sub mnuempPromociones_Click()
    Form9.Show
    Unload Me
End Sub

Private Sub mnuVentas_Click()
    Form3.Show
    Unload Me
End Sub


Private Sub MSFGEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
            Call validar(MSFGEmp.Row, MSFGEmp.Col)
    End If
End Sub

Private Sub MSFGEmp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    
        If MSFGEmp.MouseRow <> 0 And MSFGEmp.MouseCol <> 0 Then
            Call validar(MSFGEmp.MouseRow, MSFGEmp.MouseCol)
        End If
    
    End If
End Sub
