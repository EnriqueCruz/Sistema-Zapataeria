VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000005&
   Caption         =   "Opciones de Pago"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4290
   LinkTopic       =   "Form5"
   ScaleHeight     =   5040
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2400
      TabIndex        =   20
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   435
      Left            =   720
      TabIndex        =   19
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   "Tarjeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   3735
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form5.frx":0000
         Left            =   1680
         List            =   "Form5.frx":000A
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form5.frx":001F
         Left            =   1680
         List            =   "Form5.frx":0038
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Total a Pagar"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "No. Tarjeta"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Tarjeta"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Tarjeta"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Efectivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Cambio"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad Recibida"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Total a Pagar"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3735
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form5.frx":007B
         Left            =   1680
         List            =   "Form5.frx":0085
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Pago"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Porfavor , Selecciona Una Forma de Pago"
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    If Trim(UCase(Combo1.Text)) = "EFECTIVO" Then
        Frame2.Visible = True
        Frame3.Visible = False
        Text1.Text = Form3!Text2.Text
        Text2.SetFocus
    Else
        Frame2.Visible = False
        Frame3.Visible = True
        Text5.Text = Form3!Text2.Text
        Text4.SetFocus
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
    If Trim(Combo1.Text) = "" Then
        MsgBox "Selecciona el tipo de pago...", vbInformation
        Combo1.SetFocus
    Else
        If Trim(UCase(Combo1.Text)) = "EFECTIVO" Then
            If Text2.Text = "" Then
                MsgBox "Introduce el monto de pago...", vbCritical
                Text2.SetFocus
            Else
                Call guarda("E")
            End If
        Else
            If Text4.Text = "" Then
                MsgBox "Introduce el numero de tarjeta", vbCritical
                Text4.SetFocus
            Else
                Call guarda("T")
            End If
        End If
    End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub




Private Sub Text2_LostFocus()
    If Text2.Text = "" Then
        Combo1.SetFocus
    Else
        If Val(Text2.Text) < Val(Text1.Text) Then
            MsgBox "Cantidad incompleta", vbCritical, ""
            Text2.Text = ""
            Text2.SetFocus
        Else
            Text3.Text = (Int(Text2.Text)) - (Int(Text1.Text))
        End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""

End Sub

Private Sub Text4_LostFocus()
    If Combo2.Text = "" Then
        MsgBox "Tarjeta esta vacio", vbInformation, ""
        Combo2.SetFocus
    Else
        If Combo3.Text = "" Then
            MsgBox "Tipo de tarjeta esta vacio", vbInformation, ""
            Combo3.SetFocus
        End If
    End If
End Sub


Sub guarda(vtipo As String)
vord = 0
Dim id_venta As Integer
Dim cls As Integer
aac = 0
    For cls = 1 To Form3!MSFG1.Rows - 1
        Form3!MSFG1.Col = 1
        Form3!MSFG1.Row = cls
        If Form3!MSFG1.Text <> "" Then
            vid = Int(Trim(Form3!MSFG1.Text))
            Form3!MSFG1.Col = 6
            vcan = Int(Trim(Form3!MSFG1.Text))
            
            
            SQL = "select * from Producto where id_producto = " + Str(vid)
            Set dato1 = consulta(SQL, myConStr)
            Call myrsdispo(vcan, dato1)
          
            
            
            sqlv = "select * from ventas"
            Set dato2 = consulta(sqlv, myConStr)
            'Call myrsventas(vid, vempleado, dato2)
            id_venta = myrsventas(vid, vempleado, dato2)
                                    
            sqlp = "select * from pagos"
            Set dato3 = consulta(sqlp, myConStr)
            Call myrspagos(id_venta, vtipo, dato3)
            
            
            
            dato1 = Null
            dato2 = Null
            dato3 = Null
            
            aac = 1
        Else
            Exit For
        End If
    Next
    
    If aac = 1 Then
        MsgBox "Venta Realizada", vbInformation, ""
        Form3!Text1.Text = ""
        Form3!Text2.Text = ""
        Form3!MSFG1.Clear
        Call mm
        Unload Me
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
        KeyAscii = 8
    End If
End Sub

Private Sub Text4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
