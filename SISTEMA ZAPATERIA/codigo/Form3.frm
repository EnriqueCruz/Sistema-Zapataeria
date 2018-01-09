VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   5970
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14250
   LinkTopic       =   "Form3"
   ScaleHeight     =   5970
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Height          =   735
      Left            =   3840
      TabIndex        =   8
      Top             =   4920
      Width           =   6615
      Begin VB.CommandButton Command3 
         Caption         =   "Limpiar"
         Height          =   435
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   435
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   12735
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   10200
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   11520
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFG1 
         Height          =   1455
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   -2147483637
      End
      Begin VB.Label Label2 
         Caption         =   "Total"
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
         Left            =   9360
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Numero de Serie"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Tipo de Cuenta:"
      Height          =   255
      Left            =   11400
      TabIndex        =   11
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu mnuCatalogos 
      Caption         =   "Catalogos"
      Begin VB.Menu mnuEmpleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu mnuPromociones 
         Caption         =   "Promociones"
      End
      Begin VB.Menu mnuProducto 
         Caption         =   "Producto"
      End
   End
   Begin VB.Menu mnuDevolucion 
      Caption         =   "Devolucion"
   End
   Begin VB.Menu a 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu b 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu c 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu d 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu e 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu f 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu g 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu h 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu i 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu j 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu k 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu l 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu m 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu n 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ñ 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu o 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu p 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu q 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu r 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu s 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu t 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu u 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu v 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu w 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu x 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu y 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu z 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu aa 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ab 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ac 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ad 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ae 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu af 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ag 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ah 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ai 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu aj 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu ak 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu al 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu am 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu an 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu optCerrarSesion 
      Caption         =   "Cerrar Sesion"
      NegotiatePosition=   2  'Middle
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Permisos(vpuesto)
    vtotal = 0
    Label3.Caption = "Tipo de Cuenta:   " + vpuesto
    MSFG1.Cols = 10
    MSFG1.Rows = 10
    MSFG1.ColWidth(0) = 15
    MSFG1.Col = 1
    MSFG1.Row = 0
    MSFG1.Text = "No Serie"
    MSFG1.Col = 2
    MSFG1.ColWidth(2) = 2500
    MSFG1.Text = "Descripcion"
    MSFG1.Col = 3
    MSFG1.Text = "Tipo"
    MSFG1.Col = 4
    MSFG1.Text = "Color"
    MSFG1.Col = 5
    MSFG1.Text = "Talla"
    MSFG1.Col = 6
    MSFG1.Text = "Cantidad"
    MSFG1.Col = 7
    MSFG1.Text = "Precio"
    MSFG1.Col = 8
    MSFG1.Text = "Promocion"
    MSFG1.Col = 9
    MSFG1.Text = "Subtotal"

End Sub

Private Sub Command1_Click()
    On Error GoTo er
    Dim vfila As Variant
    vfila = MSFG1.RowSel
    MSFG1.Row = vfila
    
    For cl = 1 To 8
        MSFG1.Col = cl
        MSFG1.Text = ""
    Next
    
    MSFG1.Col = 9
    Text2.Text = CDbl(Text2.Text) - CDbl(MSFG1.Text)
    MSFG1.Text = ""
    
    For cl = vfila + 1 To MSFG1.Rows - 1
        MSFG1.Row = cl
        If Trim(MSFG1.Text) = "" Then
            Exit For
            
        Else
        
            For ca = 1 To 9
                MSFG1.Col = ca
                vdato = Trim(MSFG1.Text)
                MSFG1.Text = ""
                MSFG1.Row = cl - 1
                MSFG1.Text = vdato
                MSFG1.Row = cl
            Next
            
        End If
    Next
    Exit Sub
er:
    If Err.Number = 13 Then
        Err.Clear
    End If
End Sub


Private Sub Command3_Click()
    Text1.Text = ""
    Text2.Text = ""
    MSFG1.Clear
    Call Form_Load
End Sub

Private Sub Command2_Click()
    If Val(Text2.Text) > 0 Then
        Form5.Show vbModal
    End If
End Sub
Private Sub mnuDevolucion_Click()
    Form6.Show
    Unload Me
End Sub
Private Sub mnuEmpleados_Click()
    Form7.Show
    Unload Me
End Sub
Private Sub mnuProducto_Click()
    Form12.Show
    Unload Me
End Sub
Private Sub mnuPromociones_Click()
    Form9.Show
    Unload Me
End Sub
Private Sub optCerrarSesion_Click()
 vpuesto = ""
 vtotal = 0
 filas = 0
 vempleado = 0
 emple = ""
 cveprod = 0
 cveemp = 0
 cvepromo = ""
 Unload Me
 MsgBox "Hasta Luego", vbOKOnly, "Cerrar Sesión"
 Form1.Show
 
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    End If
End Sub
Private Sub Text1_LostFocus()
    If Text1.Text = "" Then
        MsgBox "Ingresa Numero de Serie", vbInformation, ""
        'Text1.SetFocus
    Else
        a = cambio
        If a = 0 Then
            SQL = "select * from Producto where Id_Producto = " + Trim(Text1.Text)
            Set dato = consulta(SQL, myConStr)
            If dato.EOF And dato.BOF Then
                MsgBox "Numero de Serie no Valido", vbCritical, ""
                Text1.Text = ""
                Text1.SetFocus
                
            Else
                Load Form4
                Form4!Text1.Text = Text1.Text
                Form4.Show vbModal
            End If
        End If
    End If
End Sub
Function cambio() As Boolean
    Dim ap As Boolean
    Dim fl As Integer
    MSFG1.Col = 1
    ap = 0
    For fl = 1 To MSFG1.Rows - 1
        MSFG1.Row = fl
        If Trim(MSFG1.Text) <> "" Then
            If Trim(Text1.Text) = Trim(MSFG1.Text) Then
                ap = 1
                Load Form4
                Form4!Text1.Text = MSFG1.Text
                MSFG1.Col = 2
                Form4!Text2.Text = MSFG1.Text
                MSFG1.Col = 3
                Form4!Text3.Text = MSFG1.Text
                MSFG1.Col = 4
                Form4!Text4.Text = MSFG1.Text
                MSFG1.Col = 5
                Form4!Text5.Text = MSFG1.Text
                MSFG1.Col = 6
                Form4!Combo1.Text = MSFG1.Text
                MSFG1.Col = 7
                Form4!Text6.Text = MSFG1.Text
                MSFG1.Col = 8
                Form4!Text7.Text = MSFG1.Text
                Form4.Caption = "MODIFICAR"
                filas = fl
                Form4.Show vbModal
                Exit For
            End If
        Else
            Exit For
        End If
    Next
    cambio = ap
End Function
Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

