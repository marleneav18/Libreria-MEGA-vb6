VERSION 5.00
Begin VB.Form FormLibros 
   Caption         =   "Agrega un libro"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   8640
      Width           =   1695
   End
   Begin VB.TextBox txtPrestadoA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   7320
      Width           =   3855
   End
   Begin VB.CheckBox chkPrestado 
      Caption         =   "Prestado?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   6720
      Width           =   3855
   End
   Begin VB.CheckBox chkRecomendado 
      Caption         =   "Recomendado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   5520
      Width           =   3855
   End
   Begin VB.CheckBox chkQuiero 
      Caption         =   "Quiero leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   4680
      Width           =   3855
   End
   Begin VB.CheckBox chkLeido 
      Caption         =   "Ya leido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3840
      Width           =   3855
   End
   Begin VB.TextBox txtCalif 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox cboGenero 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txt_autor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txt_titulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Prestamo"
      Height          =   1575
      Left            =   480
      TabIndex        =   11
      Top             =   6600
      Width           =   6615
      Begin VB.Label Label5 
         Caption         =   "Prestado a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Calificacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "FormLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditandoID As Integer

Private Sub chkLeido_Click()
    If chkLeido.Value = 1 Then
        chkQuiero.Value = 0
        txtCalif.Enabled = True
    Else
        txtCalif.Enabled = False
    End If
End Sub

Private Sub chkQuiero_Click()
    If chkQuiero.Value = 1 Then
        chkLeido.Value = 0
    End If
End Sub

Private Sub chkPrestado_Click()
   If chkPrestado.Value = 1 Then
        txtPrestadoA.Enabled = True
    Else
        txtPrestadoA.Enabled = False
        txtPrestadoA.Text = ""
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If Trim(txt_titulo.Text) = "" Or Trim(txt_autor.Text) = "" Then
        MsgBox "El titulo y el autor son obligatorios", vbExclamation, "Datos Incompletos"
        Exit Sub
    End If
    
    If cboGenero.ListIndex = -1 Then
        MsgBox "Seleccione un genero", vbExclamation, "Datos Incompletos"
        Exit Sub
    End If
    
    If chkLeido.Value = 1 And Trim(txtCalif.Text) = "" Then
        MsgBox "Ingresa una calificacion del 1 al 5", vbInformation
    End If
    
    Dim calif As Variant
    If Trim(txtCalif.Text) <> "" Then
        calif = Val(txtCalif.Text)
        If (calif < 1 Or calif > 5) Then
            MsgBox "Calificacion debe ser entre 1 y 5", vbExclamation
            Exit Sub
        End If
    Else
        calif = "NULL"
    End If
    
    ' Preparar los datos
    Dim titulo As String, autor As String, generoID As Long
    titulo = Replace(txt_titulo.Text, "'", "''")
    autor = Replace(txt_autor.Text, "'", "''")
    generoID = cboGenero.ItemData(cboGenero.ListIndex)
    
    Dim leido As Integer, porLeer As Integer, recom As Integer, prestado As Integer
    leido = IIf(chkLeido.Value = 1, 1, 0)
    porLeer = IIf(chkQuiero.Value = 1, 1, 0)
    recom = IIf(chkRecomendado.Value = 1, 1, 0)
    prestado = IIf(chkPrestado.Value = 1, 1, 0)
    
    Dim prestadoA As String, fechaPrestamo As String
    
    If prestado = 1 Then
        prestadoA = Replace(txtPrestadoA.Text, "'", "''")
        fechaPrestamo = Format$(Now, "yyyy-mm-dd")
    Else
        prestadoA = ""
        fechaPrestamo = ""
    End If
    
    On Error GoTo ErrSave
    
    Dim sqlInsert As String
    sqlInsert = "INSERT INTO Libros (Titulo, Autor, GeneroID, Calificacion, Leido, PorLeer, Recomendado, Prestado, PrestadoA, FechaPrestamo) VALUES ('" & titulo & "', '" & autor & "', " & CStr(generoID) & ", "
    
    If calif = "NULL" Then
        sqlInsert = sqlInsert & "NULL"
    Else
        sqlInsert = sqlInsert & CStr(calif)
    End If
    
    sqlInsert = sqlInsert & ", " & CStr(leido) & ", " & CStr(porLeer) & ", " & CStr(recom) & ", " & CStr(prestado)
    
    If prestado = 1 Then
    sqlInsert = sqlInsert & ", '" & prestadoA & "', '" & fechaPrestamo & "')"
Else
    sqlInsert = sqlInsert & ", NULL, NULL)"
End If
    
    conn.Execute sqlInsert
    MsgBox "Libro agregado exitosamente", vbInformation
    
Exit Sub

ErrSave:
    MsgBox "Ocurrio un error al guardar: " & Err.Description, vbCritical

    
        
    
End Sub

Private Sub Form_Load()
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    rsG.Open "Select GeneroID, Nombre FROM Generos ORDER BY Nombre", conn, adOpenStatic, adLockReadOnly
    cboGenero.Clear
    Do Until rsG.EOF
        cboGenero.AddItem rsG!Nombre
        cboGenero.ItemData(cboGenero.NewIndex) = rsG!generoID
        rsG.MoveNext
    Loop
    
    rsG.Close: Set rsG = Nothing
    
    If EditandoID = 0 Then
        ' Modo agregar, limpiar campos
        
        txt_titulo.Text = ""
        txt_autor.Text = ""
        cboGenero.ListIndex = -1 ' NO HAY NADA SELECCIONADO
        txtCalif = ""
        chkLeido.Value = 0
        txtPrestadoA.Enabled = False
        Me.Caption = "Agregar Libro"
        
    Else
        
    End If
End Sub
