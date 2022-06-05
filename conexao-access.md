Option Explicit

Private Banco As New ADODB.Connection
Private Tabela As New ADODB.Recordset

Public SQL As String
Public TotalRegistros As Integer
Public TotalColunas As Integer
Public resultado As Variant
Public NomeColuna(100) As Variant

Function Conectar() As Boolean

On Error GoTo Erro

    Set Banco = New ADODB.Connection
    
    Dim CaminhoBanco As String
    CaminhoBanco = bd.Range("A1").value
    Dim senha As String
    senha = bd.Range("B1").value
    senha = Criptografia(senha)
    
    With Banco
         .Provider = "Microsoft.ACE.OLEDB.12.0"
         .ConnectionString = "Data Source=" & CaminhoBanco & "; Jet OLEDB:Database Password=" & senha
         .Mode = adModeReadWrite
    
         .Open
    End With
    
Erro:

    If Banco.State = adStateOpen Then
       Conectar = True
    Else
       Conectar = False
       Cod.Erro ("Não foi possível conectar ao banco de dados")
    End If
    
End Function

Sub Desconectar()

    If Tabela.State = adStateOpen Then Tabela.Close
    If Banco.State = adStateOpen Then Banco.Close
    
    Set Tabela = Nothing
    Set Banco = Nothing
    
End Sub

Function ExecuteSQL(Optional Msg As String = Empty) As Boolean
    
On Error GoTo Erro
    
    ExecuteSQL = False
    
    If Conectar = False Then Exit Function
    
    Banco.Execute (SQL)
    
    Call Desconectar
    
    If Msg <> Empty Then Cod.OK (Msg)
    
     ExecuteSQL = True
    
    Exit Function
Erro:

    Call Desconectar
    Cod.Erro (Err.Description)
    
End Function

Public Function Buscar(Optional Aba As String = Empty, _
           Optional ColI As String = Empty, _
           Optional LinI As Integer = Empty, _
           Optional ColF As String = Empty, _
           Optional Msg As Boolean = False) As Boolean

On Error GoTo Erro

Dim i As Long

   Buscar = False

    Set Tabela = New ADODB.Recordset
    
    Set resultado = Nothing
    
    Call Conectar
    
    Tabela.CursorLocation = adUseClient
    
    Tabela.Open SQL, Banco, adOpenStatic
    
    TotalRegistros = Tabela.RecordCount
    TotalColunas = Tabela.Fields.Count - 1
    
    For i = 0 To TotalColunas
       NomeColuna(i) = Tabela.Fields.Item(i).Name
    Next i
    
    If TotalRegistros = 0 Then
    
        TotalRegistros = -1
        
        If Aba <> Empty Then
           Call Cod.Borda(Aba, ColI, LinI, ColF, 0)
           Call Cod.LimparPlanilha(Aba, ColI, LinI, ColF)
        End If
        
        Buscar = False
        
        GoTo Erro
    End If
    
    TotalRegistros = TotalRegistros - 1
    
    Buscar = True
    
    If Aba = Empty Then
        resultado = Tabela.GetRows
    Else
    
        Call Cod.Borda(Aba, ColI, LinI, ColF, 0)
        
        Call Cod.LimparPlanilha(Aba, ColI, LinI, ColF)
        
        Sheets(Aba).Range(ColI & LinI).CopyFromRecordset Tabela
        
        Call Cod.Borda(Aba, ColI, LinI, ColF, 1)
        
    End If
    
Erro:

    Call Desconectar
    
    If Err Then
       Cod.Erro (Err.Description)
    End If
    
End Function

Sub ValidarAcessoUsuario()
   
   Dim senha As String
   
   With PlanLogin
   
   If .txtUsuario = Empty Or .txtSenha = Empty Then
      Call Cod.Erro("Favor preencher todos os campos!")
      Exit Sub
   End If
   
   senha = Criptografia(.txtSenha)
   
      SQL = "SELECT LOGIN FROM USUARIO WHERE LOGIN = '" & .txtUsuario & "'"
      If Buscar = False Then Exit Sub
   
      If TotalRegistros <> 0 Then
         Call Cod.Erro("Usuário não cadastrado!")
         Exit Sub
      Else
         
         SQL = "SELECT SENHA FROM USUARIO WHERE SENHA = '" & senha & "' AND LOGIN = '" & .txtUsuario & "'"
         If Buscar = False Then Exit Sub
         
         If TotalRegistros <> 0 Then
            Call Cod.Erro("Senha incorreta!")
            Exit Sub
         End If
      
      Call Cod.OK("Login realizado com sucesso!")
      
      .txtUsuario = Empty: .txtSenha = Empty
      
      PlanMenu.Visible = xlSheetVisible
      PlanMenu.Activate
      PlanLogin.Visible = xlSheetHidden
      Call Cod.TelaNormal
      
      End If
   
   End With
   
End Sub

Sub ValidarBanco()

   If ValidarCampos(frmConfigurarBD.Controls) = True Then Exit Sub
   
   Dim senha As String
   senha = Criptografia(frmConfigurarBD.txtSenha)
   
   With bd
        .Range("A1").value = frmConfigurarBD.txtArquivo
        .Range("B1").value = senha
   End With
   
   If Conectar = True Then
      Call Cod.OK("Configurações validadas com sucesso...")
      Call Desconectar
      Unload frmConfigurarBD
   End If

End Sub

Private Function Criptografia(TextoSenha As String) As String

    Dim CaractereTemporario As String
    Dim i As Integer
    
    For i = 1 To Len(TextoSenha)

        If VBA.Asc(VBA.Mid$(TextoSenha, i, 1)) < 128 Then
           CaractereTemporario = VBA.Asc(VBA.Mid$(TextoSenha, i, 1)) + 128
        ElseIf VBA.Asc(VBA.Mid$(TextoSenha, i, 1)) > 128 Then
           CaractereTemporario = VBA.Asc(VBA.Mid$(TextoSenha, i, 1)) - 128
        End If

        Mid$(TextoSenha, i, 1) = VBA.Chr(CaractereTemporario)

    Next i

    Criptografia = TextoSenha

End Function