VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3540
   ClientLeft      =   5310
   ClientTop       =   1770
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4575
   Begin VB.TextBox txtUF 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1155
      MaxLength       =   2
      TabIndex        =   13
      Top             =   2295
      Width           =   450
   End
   Begin VB.TextBox txtCidade 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1155
      TabIndex        =   12
      Top             =   1845
      Width           =   3120
   End
   Begin VB.TextBox txtBairro 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1155
      TabIndex        =   11
      Top             =   1395
      Width           =   3120
   End
   Begin VB.TextBox txtEndereco 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1155
      TabIndex        =   10
      Top             =   945
      Width           =   3120
   End
   Begin MSAdodcLib.Adodc adoPostgree 
      Height          =   330
      Left            =   1740
      Top             =   2385
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PostgreSQL35W"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PostgreSQL35W"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from ""cep_endereco"""
      Caption         =   "adoPostgree"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btnExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3405
      TabIndex        =   9
      Top             =   3060
      Width           =   915
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1845
      TabIndex        =   8
      Top             =   3075
      Width           =   915
   End
   Begin VB.CommandButton btnInserir 
      Caption         =   "Inserir"
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   3075
      Width           =   915
   End
   Begin VB.TextBox txtcep 
      Height          =   285
      Left            =   660
      MaxLength       =   8
      TabIndex        =   1
      Top             =   300
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   255
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "UF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   2415
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   465
      TabIndex        =   5
      Top             =   1935
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   555
      TabIndex        =   4
      Top             =   1500
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   1065
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CEP:"
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public temNobanco As Boolean

Public Function VerificaCepNoBanco(CEP As String) As Boolean
    adoPostgree.RecordSource = "Select * from cep_endereco where cep = '" + CEP + "'"
    adoPostgree.Refresh
    
    If adoPostgree.Recordset.RecordCount > 0 Then
        VerificaCepNoBanco = True
        temNobanco = True
    Else
        VerificaCepNoBanco = False
        temNobanco = False
    End If
    
End Function

Private Sub btnEditar_Click()
    
    If btnEditar.Caption = "Editar" Then
        txtEndereco.Enabled = True
        txtBairro.Enabled = True
        txtCidade.Enabled = True
        txtUF.Enabled = True
        
        txtEndereco.SetFocus
        btnEditar.Caption = "Salvar"
    Else
        adoPostgree.Recordset.fields("logradouro") = txtEndereco.Text
        adoPostgree.Recordset.Update
        
        MsgBox "CEP editado com sucesso!"
    
        btnEditar.Caption = "Editar"
    End If
    
    
End Sub

Private Sub btnExcluir_Click()
    adoPostgree.Recordset.Delete
    
    MsgBox "CEP excluído do banco de dados!"
    
    txtcep.Text = ""
    txtEndereco.Text = ""
    txtBairro.Text = ""
    txtCidade.Text = ""
    txtUF.Text = ""
    
End Sub

Private Sub btnInserir_Click()

    If VerificaCepNoBanco(txtcep.Text) Then 'Verifica se o CEP já existe no banco
        MsgBox "O banco de dados já possui as informações do CEP " + txtcep.Text
    Else
        Call insereCep
        MsgBox "O CEP " + txtcep.Text + " foi inserido no banco de dados!"
        txtEndereco.Text = ""
        txtBairro.Text = ""
        txtCidade.Text = ""
        txtUF.Text = ""
        
    End If
    
    btnInserir.Enabled = False

End Sub

Private Sub Command1_Click()
    
    'Valida preenchimento do campo CEP
    If txtcep.Text = "" Then
        MsgBox "Informe um CEP"
        txtcep.SetFocus
        Exit Sub
    End If
    
    If VerificaCepNoBanco(txtcep.Text) Then 'Verifica se o CEP informado já está cadastrado no banco de dados
        txtEndereco.Text = adoPostgree.Recordset.fields("logradouro")
        txtBairro.Text = adoPostgree.Recordset.fields("bairro")
        txtCidade.Text = adoPostgree.Recordset.fields("cidade")
        txtUF.Text = adoPostgree.Recordset.fields("uf")
    Else 'O CEP não está no banco de dados, então busco via API
        txtEndereco.Text = ConsultaCep(txtcep.Text, "logradouro")
        txtBairro.Text = ConsultaCep(txtcep.Text, "bairro")
        txtCidade.Text = ConsultaCep(txtcep.Text, "localidade")
        txtUF.Text = ConsultaCep(txtcep.Text, "uf")
    End If
    
    btnInserir.Enabled = txtEndereco.Text <> ""
    
    If temNobanco = True Then
        btnEditar.Enabled = True
        btnExcluir.Enabled = True
    Else
        btnEditar.Enabled = False
        btnExcluir.Enabled = False
    End If
        
    
    
End Sub

Private Sub insereCep()
    adoPostgree.Refresh
    adoPostgree.Recordset.AddNew
    adoPostgree.Recordset.fields("cep") = txtcep.Text
    adoPostgree.Recordset.fields("logradouro") = txtEndereco.Text
    adoPostgree.Recordset.fields("bairro") = txtBairro.Text
    adoPostgree.Recordset.fields("cidade") = txtcep.Text
    adoPostgree.Recordset.fields("uf") = txtUF.Text
    
    adoPostgree.Recordset.Update
End Sub





Private Sub Form_Activate()
    txtcep.SetFocus
End Sub

