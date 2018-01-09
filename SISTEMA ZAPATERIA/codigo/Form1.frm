VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2910
   ClientLeft      =   9150
   ClientTop       =   5235
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3885
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Sub Command1_Click()

    If Text1.Text = "" Then
        MsgBox "Introduce el Usuario"
        Text1.SetFocus
    Else
        SQL = " select * from Empleado where Usuario = '" + Trim(Text1.Text) + "'"
        Set dato = New ADODB.Connection
        Set dato = consulta(SQL, myConStr)
        
        
        If dato.EOF And dato.BOF Then
            MsgBox "Usuario no Valido"
            Text1.Text = ""
            Text1.SetFocus
        Else
            'emple = dato!
            vempleado = dato!id_empleado
            If IsNull(dato!Password) Or dato!Password = "" Then
                Load Form2
                Form2!Text1.Text = Text1.Text
                Form2.Show
                Unload Me
            Else
            vempleado = dato!id_empleado
               vps = dato!Password
               sqlperf = "select * from empleado,perfil  where Empleado.Id_Perfil = Perfil.Id_Perfil and empleado.Id_Empleado = " + Str(vempleado)
               Set datoperf = consulta(sqlperf, myConStr)
               vpuesto = datoperf!Tipo
               If Trim(Text2.Text) = Trim(vps) Then
               
                'MsgBox "Bienvenido", vbApplicationModal
                
                Load Form3
                Form3.Show
                Unload Me
            Else
                MsgBox "Password Incorrecto"
                Text2.Text = ""
                Text2.SetFocus
            End If
        End If
        End If
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub


