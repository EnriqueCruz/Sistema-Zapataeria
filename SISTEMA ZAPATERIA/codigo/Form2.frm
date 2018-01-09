VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0080FF80&
   Caption         =   "Crear Contraseña"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4995
   LinkTopic       =   "Form2"
   ScaleHeight     =   3630
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Confirmar Contraseña"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Text2.Text = "" Then
        MsgBox "Introduce Password"
        Text1.SetFocus
    Else
        If Trim(Text2.Text) <> Trim(Text3.Text) Then
            MsgBox "Password no Coincide", vbExclamation, "Aviso"
            Text2.Text = ""
            Text3.Text = ""
            Text2.SetFocus
        Else
            SQL = "SELECT * From Empleado Where Usuario = '" + Trim(Text1.Text) + "'"
            'Set can = New ADODB.Connection
            Set can = consulta(SQL, myConStr)
            Call guardap(Trim(Text2.Text), can)
            
            sqlperf = "select * from empleado,perfil  where Empleado.Id_Perfil = Perfil.Id_Perfil and empleado.Id_Empleado = " + Str(vempleado)
            Set datoperf = consulta(sqlperf, myConStr)
            vpuesto = datoperf!Tipo
            Load Form3
            Form3.Show
            Unload Me
            
        End If
    End If
End Sub

