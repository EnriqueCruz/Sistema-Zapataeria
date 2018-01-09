VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "AGREGAR"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4230
   LinkTopic       =   "Form4"
   ScaleHeight     =   5595
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   17
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Informacion del Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   18
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Promocion"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Talla"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Numero de Serie"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
    Dim vsub As Integer
    Dim vtotal_ant As Integer
    Dim vtot As Integer

    If Form4.Caption = "AGREGAR" Then
        vtotal = 0
        If Combo1.Text = "" Then
            MsgBox "Por favor Selecciona la cantidad", vbInformation, ""
            Combo1.SetFocus
        Else
        
            For i = 1 To Form3!MSFG1.Rows - 1
                Form3!MSFG1.Row = i
                If Trim(Form3!MSFG1.Text) = "" Then
                    Form3!MSFG1.Col = 1
                    Form3!MSFG1.Text = Trim(Text1.Text)
                    Form3!MSFG1.Col = 2
                    Form3!MSFG1.Text = Trim(Text2.Text)
                    Form3!MSFG1.Col = 3
                    Form3!MSFG1.Text = Trim(Text3.Text)
                    Form3!MSFG1.Col = 4
                    Form3!MSFG1.Text = Trim(Text4.Text)
                    Form3!MSFG1.Col = 5
                    Form3!MSFG1.Text = Trim(Text5.Text)
                    Form3!MSFG1.Col = 6
                    vcant = Trim(Combo1.Text)
                    Form3!MSFG1.Text = vcant
                    Form3!MSFG1.Col = 7
                    vpre = Trim(Text6.Text)
                    Form3!MSFG1.Text = vpre
                    Form3!MSFG1.Col = 8
                    vpromo = Trim(Text7.Text)
                    Form3!MSFG1.Text = vpromo
                    Form3!MSFG1.Col = 9
                    vsub = (Int(vcant) * Int(vpre)) * ((100 - vpromo) / 100)
                    Form3!MSFG1.Text = vsub
                    vtotal = vtotal + vsub
                    Exit For
                    
                Else
                    Form3!MSFG1.Col = 9
                    vtotal = vtotal + Int(Form3!MSFG1.Text)
                End If
            Next
            Form3!Text1.Text = ""
            Form3!Text2.Text = vtotal
            Unload Me
        End If
    Else
        If Combo1.Text <> "" Then
            Form3!MSFG1.Row = filas
            Form3!MSFG1.Col = 6
            Form3!MSFG1.Text = Combo1.Text
            Form3!MSFG1.Col = 9
            vtotal_ant = Form3!MSFG1.Text
            Form3!MSFG1.Row = filas
            vsub = (Int(Combo1.Text) * Int(Text6.Text)) - ((100 - Int(Text7.Text)) / 100)
            Form3!MSFG1.Text = vsub
            vtot = vsub - vtotal_ant
            Form3!Text2.Text = Int(Form3!Text2.Text) + vtot
            Unload Me
        Else
            MsgBox "Ingresa cantidad", vbInformation, ""
            Combo1.SetFocus
        End If
        Form3!Text1.Text = ""
    End If
End Sub

Private Sub Command2_Click()
    Form3!Text1.Text = ""
    Unload Me
End Sub



Private Sub Form_Load()
    SQL = "Select * from Producto left join Promocion on(Producto.Id_producto = Promocion.Id_producto) where  producto.id_producto = " + Form3!Text1.Text
    
    Set dato = consulta(SQL, myConStr)
    
    'vcant = dato!Disponibilidad
    vcant = dato!Disponibilidad
    If vcant > 0 Then
        For i = 1 To vcant
            Combo1.AddItem i
        Next
        
        Text2.Text = dato("producto.Descripcion")
        Text3.Text = dato!Tipo
        Text4.Text = dato!Color
        Text5.Text = dato!Talla
        Text6.Text = dato!Precio
        If dato!descuento > 0 Then
            Text7.Text = dato!descuento
        Else
            Text7.Text = 0
        End If
        
    Else
        MsgBox "Sin existencia", vbExclamation, ""
        Unload Me
    End If
    
End Sub
