VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_eliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   8280
      TabIndex        =   10
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton btn_modificar 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton libros_fav 
      Caption         =   "Libros favoritos"
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton generos_fav 
      Caption         =   "Generos favoritos"
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton no_gustaron 
      Caption         =   "No me gustaron"
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton quiero_leer 
      Caption         =   "Quiero Leer"
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton btn_agregar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin MSComctlLib.ListView list_libros 
      Height          =   4335
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btn_leiste 
      Caption         =   "Ya leiste"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton btn_catalogo 
      Caption         =   "Catalogo MEGA"
      Height          =   735
      Left            =   600
      Picture         =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarLibros(filtroSQL As String)
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT L.LibroID, L.Titulo, L.Autor, G.Nombre AS Genero, L.Calificacion, L.Prestado, L.PrestadoA FROM Libros L INNER JOIN Generos G ON L.GeneroID = G.GeneroID"
    
    If filtroSQL <> "" Then
        sql = sql & " WHERE " & filtroSQL
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    list_libros.ListItems.Clear
     
    If Not rs.EOF Then
        Dim item As ListItem
        Do Until rs.EOF
            Set item = list_libros.ListItems.Add(, , rs!titulo)
            item.SubItems(1) = rs!autor
            item.SubItems(2) = rs!Genero
            item.SubItems(3) = IIf(IsNull(rs!Calificacion), "", rs!Calificacion)
            
            If rs!prestado = True Then
                item.SubItems(4) = rs!prestadoA
            Else
                item.SubItems(4) = ""
            End If
            
            item.Tag = rs!libroID
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close: Set rs = Nothing
    
End Sub

Private Sub btn_agregar_Click()
    FormLibros.EditandoID = 0
    FormLibros.Show vbModal
End Sub

Private Sub btn_catalogo_Click()
    CargarLibros ""
End Sub

Private Sub btn_eliminar_Click()
    Dim item As ListItem
    Set item = list_libros.SelectedItem
    
    If item Is Nothing Then
        MsgBox "Selecciona el libro a eliminar", vbExclamation
        Exit Sub
    End If
    
    Dim titulo As String
    titulo = item.Text
    Dim resp As Integer
    resp = MsgBox("Estas seguro de eliminar el libro '" & titulo & "'?", vbYesNo + vbQuestion, "Confirmar eliminacion")
    
    If resp = vbYes Then
        Dim libroID As Long
        libroID = item.Tag
        On Error GoTo ErrorDelete
        conn.Execute "DELETE FROM Libros WHERE LibroID=" & CStr(libroID)
        MsgBox "Libro Eliminado", vbInformation
        CargarLibros ""
    End If

    Exit Sub
    
ErrorDelete:
    MsgBox "Error eliminando libro" & Err.Description, vbCritical

    
End Sub

Private Sub btn_leiste_Click()
    CargarLibros "L.Leido = 1"
End Sub

Private Sub btn_modificar_Click()
    FormLibros.EditandoID = list_libros.SelectedItem.Tag
    FormLibros.Show vbModal
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    
    Dim connString As String
    connString = "Provider=SQLOLEDB.1;DATA Source=Quiroz;Initial Catalog=LibreriaMega;Integrated Security=SSPI;"
        
    conn.Open connString
    
    With list_libros
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Titulo", 2000
        .ColumnHeaders.Add , , "Autor", 1500
        .ColumnHeaders.Add , , "Género", 1000
        .ColumnHeaders.Add , , "Calificación", 1100
        .ColumnHeaders.Add , , "Prestado a", 1500
        
    End With
    CargarLibros ""
End Sub

Private Sub generos_fav_Click()
    CargarLibros "G.EsFavorito = 1"
End Sub

Private Sub libros_fav_Click()
    CargarLibros "L.Recomendado = 1"
End Sub

Private Sub no_gustaron_Click()
    CargarLibros "L.Leido = 1 AND L.Calificacion <= 2"
End Sub

Private Sub quiero_leer_Click()
    CargarLibros "L.PorLeer = 1"
End Sub
