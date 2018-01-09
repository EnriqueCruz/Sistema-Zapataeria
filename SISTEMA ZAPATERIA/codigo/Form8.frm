VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form8"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5565
   LinkTopic       =   "Form8"
   ScaleHeight     =   7335
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Informacion del Empleado"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form8.frx":0000
         Left            =   1560
         List            =   "Form8.frx":0002
         TabIndex        =   16
         Top             =   5040
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox Textcel 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Texttel 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000B&
         Caption         =   "Puesto"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000B&
         Caption         =   "Correo"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000B&
         Caption         =   "Celular"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000B&
         Caption         =   "Telefono"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000B&
         Caption         =   "Direccion"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         Caption         =   "RFC"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         Caption         =   "Nombre del Trabajador"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 1 And KeyAscii <= 255) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
    Dim numtel As Variant
    Dim numcel As Variant
    
    numtel = 0
    numcel = 0
        
    numtel = CDbl(Texttel.Text)
    numcel = CDbl(Textcel.Text)
    
    
    If Form8.Caption = "Alta" Then
        If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Texttel.Text = "" Or Textcel.Text = "" Or Text7.Text = "" Or Combo1.Text = "" Then
            MsgBox "Faltan Campos por Llenar", vbInformation
        Else
            vip = Combo1.ListIndex + 1
            Call agregaemp(vip, Text1.Text, myConStr, Text2.Text, Text3.Text, Text4.Text, numtel, numcel, Text7.Text)
            MsgBox "Empleado Agregado"
            Unload Me
            Load Form6
            Call dataemp
        End If
        
    Else
        If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Texttel.Text = "" Or Textcel.Text = "" Or Text7.Text = "" Or Combo1.Text = "" Then
            MsgBox "Faltan Campos por Llenar", vbInformation
        Else
            vip = Combo1.ListIndex + 1
            'Set datemmod = New ADODB.Connection
            'Set datemmod = consulta(SQLempmod, myConStr)
            SQLempmod = "select * from empleado where Id_Empleado = " + Str(cveprod)
            Set tabemp = consulta(SQLempmod, myConStr)
            Call modemp(vip, Text1.Text, tabemp, Text2.Text, Text3.Text, Text4.Text, numtel, numcel, Text7.Text)
            MsgBox "Registro Modificado Satisfactoriamente", vbInformation
            Unload Me
            Load Form6
            Call dataemp
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SQL = "select * from perfil"
    Set puesto = consulta(SQL, myConStr)
    If puesto.EOF And puesto.BOF Then
        MsgBox "No exite el perfil"
    Else
        puesto.MoveFirst
        con = 1
        Do While (Not (puesto.EOF))
            Combo1.AddItem puesto!Tipo
            puesto.MoveNext
            con = con + 1
        Loop
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If (KeyAscii >= 48) And (KeyAscii <= 57) Then
    KeyAscii = 8
 End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub Textcel_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Textcel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub Texttel_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Texttel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
