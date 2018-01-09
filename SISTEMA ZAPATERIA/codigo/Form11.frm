VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form12 
   BackColor       =   &H8000000B&
   Caption         =   "Form12"
   ClientHeight    =   5625
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11850
   LinkTopic       =   "Form12"
   ScaleHeight     =   5625
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Modificacion"
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   10440
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid MSFGProd 
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5530
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   11895
   End
   Begin VB.Menu mnuprodCatalogos 
      Caption         =   "Catalogos"
      Begin VB.Menu mnuprodEmpleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu mnuprodPromociones 
         Caption         =   "Promociones"
      End
   End
   Begin VB.Menu mnuprodVentas 
      Caption         =   "Ventas"
   End
   Begin VB.Menu mnuprodDevoluciones 
      Caption         =   "Devoluciones"
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub validar(iRow As Integer, iCol As Integer)
   
        If MSFGProd.TextMatrix(iRow, 1) = strUnChecked Then
            MSFGProd.TextMatrix(iRow, 1) = strChecked
        Else
            MSFGProd.TextMatrix(iRow, 1) = strUnChecked
        End If
    
End Sub



Private Sub Command3_Click()
    Dim d As Integer
    Dim tabl As String
    Dim clave2 As Integer
    tabl = "Producto"
        
    For z = 1 To MSFGProd.Rows - 1
        MSFGProd.Row = z
        MSFGProd.Col = 1
        If MSFGProd.Text = strChecked Then
            d = d + 1
        End If
    Next
    
    If d = 0 Then
    MsgBox "Selecciona un Registro Porfavor", vbInformation, ""
    
    Else
        For a = 1 To MSFGProd.Rows - 1
            MSFGProd.Row = a
            MSFGProd.Col = 1
            If MSFGProd.Text = strChecked Then
                MSFGProd.Col = 2
                clave2 = Int(MSFGProd.Text)
                Call borrar(clave2, tabl, myConStr)
            End If
        Next
            MsgBox "Registro(s) Eliminado(s)", vbInformation, ""
    End If
Call datprod
End Sub

Private Sub MSFGProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
            Call validar(MSFGProd.Row, MSFGProd.Col)
    End If
End Sub

Private Sub MSFGProd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    
        If MSFGProd.MouseRow <> 0 And MSFGProd.MouseCol <> 0 Then
            Call validar(MSFGProd.MouseRow, MSFGProd.MouseCol)
        End If
    
    End If
End Sub



Private Sub Command1_Click()
    Form11.Caption = "Agregar Producto"
    Form11.Show vbModal
End Sub

Private Sub Command2_Click()
    
Dim c2 As Integer
Dim b2 As Integer
Dim filr12 As Integer
c2 = 0
'filr1 = 0
    For x = 1 To MSFGProd.Rows - 1
        MSFGProd.Row = x
        MSFGProd.Col = 1
        If MSFGProd.Text = strChecked Then
            c2 = c2 + 1
           Exit For
        End If
    filr12 = x + 1
    Next
    
    If filr12 = 0 Then
        filr12 = filr12 + 1
    End If
    
    For y = 1 To MSFGProd.Rows - 1
        MSFGProd.Row = y
        MSFGProd.Col = 1
        If MSFGProd.Text = strChecked Then
            b2 = b2 + 1
        End If
    Next
    
            If c2 = 1 And b2 = 1 Then
                
                MSFGProd.Row = filr12
                MSFGProd.Col = 2
                cveemp = Int(MSFGProd.Text)
                MSFGProd.Col = 3
                Form11!Text1.Text = MSFGProd.Text
                MSFGProd.Col = 4
                Form11!Text2.Text = MSFGProd.Text
                MSFGProd.Col = 5
                Form11!Text3.Text = MSFGProd.Text
                MSFGProd.Col = 6
                Form11!Text4.Text = MSFGProd.Text
                MSFGProd.Col = 7
                Form11!Text5.Text = MSFGProd.Text
                MSFGProd.Col = 8
                Form11!Text6.Text = MSFGProd.Text
                MSFGProd.Col = 9
                Form11!Combo1.Text = MSFGProd.Text
                MSFGProd.Col = 10
                Form11!Text10.Text = MSFGProd.Text
                MSFGProd.Col = 11
                Form11!Text8.Text = MSFGProd.Text
                MSFGProd.Col = 12
                Form11!Text9.Text = MSFGProd.Text
                Form11.Caption = "Modificacion"
                Form11.Show vbModal
            Else
                If c2 = 1 And b2 >= 2 Then
                    MsgBox "Solo se puede seleccionar un Producto para modificar", vbInformation, ""
                Else
                    If c2 = 0 And b2 = 0 Then
                        MsgBox "Selecciona un Elemento", vbInformation, ""
                    End If
            End If
        End If
End Sub
    
Private Sub Form_Load()
    Call datprod
End Sub

Private Sub mnuprodDevoluciones_Click()
    Form6.Show
    Unload Me
End Sub

Private Sub mnuprodEmpleados_Click()
    Form7.Show
    Unload Me
End Sub

Private Sub mnuprodPromociones_Click()
    Form9.Show
    Unload Me
End Sub

Private Sub mnuprodVentas_Click()
    Form3.Show
    Unload Me
End Sub


