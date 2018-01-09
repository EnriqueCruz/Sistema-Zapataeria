VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H8000000B&
   Caption         =   "Form10"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   LinkTopic       =   "Form10"
   ScaleHeight     =   4395
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2880
      TabIndex        =   7
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   435
      Left            =   1440
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informacion de la Promocion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Promocion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
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
      Width           =   5535
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     If Form10.Caption = "Agregar Promocion" Then
        If Text1.Text = "" Or Text2.Text = "" Then
            MsgBox "Faltan Campos por Llenar", vbInformation, ""
        Else
            'vip = Combo1.ListIndex + 1
            Call agregaPromo(myConStr, Trim(Text1.Text), Trim(Text2.Text))
            MsgBox "Promocion Agregada", vbInformation, ""
            Unload Me
            Load Form12
            Call datprom
        End If
    Else
        If Text1.Text = "" Or Text2.Text = "" Then
            MsgBox "Faltan Campos por Llenar", vbInformation, ""
        Else
            'vip = Combo1.ListIndex + 1
            'Set mpro = consulta()
            Call modpromo(myConStr, Trim(Text1.Text), Trim(Text2.Text), cvepromo)
            MsgBox "Promocion Actualizada", vbInformation
            Unload Me
            Load Form12
            Call datprom
        End If
    End If
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
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        KeyAscii = 8
    End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
