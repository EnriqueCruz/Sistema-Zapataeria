VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H8000000E&
   Caption         =   "Form11"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   LinkTopic       =   "Form11"
   ScaleHeight     =   8220
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   7680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Informacion del Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4815
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form12.frx":0000
         Left            =   1680
         List            =   "Form12.frx":000A
         TabIndex        =   23
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   4920
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   6120
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000B&
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000B&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000B&
         Caption         =   "Disponibilidad"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000B&
         Caption         =   "Genero"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000B&
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000B&
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         Caption         =   "Talla"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Descripcion 
         BackColor       =   &H8000000B&
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        KeyAscii = 8
    End If
End Sub
Private Sub Command1_Click()
    If Form11.Caption = "Agregar Producto" Then
        If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
            MsgBox "Faltan Campos por Llenar", vbInformation
        Else
            'vip = Combo1.ListIndex + 1
            Call agregaprod(myConStr, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Combo1.Text, Text8.Text, Text9.Text, Text10.Text)
            MsgBox "Producto Agregado", vbInformation, ""
            Unload Me
            Load Form12
            Call datprod
        End If
    Else
        If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
            MsgBox "Faltan Campos por Llenar", vbInformation
        Else
            Call modpro(myConStr, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Combo1.Text, Text10.Text, Text8.Text, Text9.Text, cveemp)
            MsgBox "Informacion del Producto Modificada", vbInformation, ""
            Unload Me
            Load Form12
            Call datprod
        End If
    End If
End Sub
Private Sub Command2_Click()
    Close
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        KeyAscii = 8
    End If
End Sub
Private Sub Text10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
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
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        KeyAscii = 8
    End If
End Sub
Private Sub Text3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
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
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        KeyAscii = 8
    End If
End Sub
Private Sub Text7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Text8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        KeyAscii = 8
    End If
End Sub
Private Sub Text9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub
