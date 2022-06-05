Option Explicit

#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Global Resposta As Boolean
Global TextoPesquisa As String

Function MascaraDecimal(ByVal KeyAscii As MSForms.ReturnInteger, ByVal txt As MSForms.TextBox) As String

Dim n As Double, n1 As Long, n2 As Long, nFor As Double
    
    txt.Locked = True
    
    If KeyAscii = 9 Or KeyAscii = 13 Or KeyAscii = 16 Then
        MascaraDecimal = txt
        Exit Function
    End If
    
    If VBA.Len(txt) > 9 Then Exit Function
    
    If KeyAscii < 96 Or KeyAscii > 105 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
                Exit Function
            End If
        End If
    End If
    
    If txt = Empty Or VBA.IsNumeric(txt) = False Then
        n = 0
    Else
        n = txt * 100
    End If
    
    If KeyAscii = 8 Then
        If VBA.Len(n) = 1 Then
            n2 = 0
        Else
            n2 = VBA.Left(n, VBA.Len(n) - 1)
        End If
    Else
        If KeyAscii < 48 Or KeyAscii > 57 Then
            n1 = VBA.Chr(KeyAscii - 48)
        Else
            n1 = VBA.Chr(KeyAscii)
        End If
        n2 = n & n1
    End If
    
    nFor = n2 / 100
    
    MascaraDecimal = VBA.FormatNumber(nFor, 2)

End Function

Sub Rotina(ByVal tempo As Variant, ByVal ScriptExec As String)
    Application.OnTime Now + VBA.TimeValue(tempo), ScriptExec
End Sub

Sub OcultarSaldos()

    With Plan1
        
        If .Range("A1") < 1 Then
            .Range("A1") = 1
            Call AtualizarQuery
            .Shapes("Total").Visible = True
        Else
            .Range("A1") = 0
            .Range("C8:C5000").ClearContents
            .Shapes("Total").Visible = False
        End If
        
    End With

End Sub

Sub AtualizarQuery()

Dim i As Integer

    SQL = "SELECT NOME, SALDO FROM CONTAS WHERE SALDO <> 0 AND SALDO_ATIVO = 1 ORDER BY SALDO DESC"
    If Buscar = False Then Exit Sub
    
    With Plan1
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        .Range("B8:C100").ClearContents
        
        For i = 0 To TotalRegistros
            .Cells(i + 8, 2) = resultado(0, i)
            .Cells(i + 8, 3) = resultado(1, i)
        Next i
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
    End With
    
    'Call Cod.AtualizarBase

End Sub

Sub AtualizarBase()

Dim i As Long

    SQL = "SELECT A.CÓDIGO, A.CLASSE, B.NOME, A.DATA, A.VALOR, A.DESCRICAO, FORMAT(A.DATA, 'MMM'), YEAR(A.DATA), "
    SQL = SQL & "FORMAT(A.DATA, 'DDDD') FROM BD A INNER JOIN CONTAS B ON A.CONTA = B.CÓDIGO"
    If Buscar = False Then Exit Sub
    Call Buscar("BD", "A", 2, "I")
    
End Sub

Public Function FormataData(ByVal valor As Variant) As String
 
Dim data As String
Dim mes As String
mes = VBA.Format(Now(), "mm")
Dim ano As String
ano = VBA.Format(Now(), "yyyy")
   
    If valor = Empty Then Exit Function
    
    data = VBA.Replace(valor, "/", "")
    data = VBA.Replace(data, ".", "")
    data = VBA.Replace(data, "-", "")
    
    If VBA.Len(data) = 2 Then
       data = data & "/" & mes & "/" & ano
    ElseIf VBA.Len(data) = 4 Then
       data = VBA.Mid(data, 1, 2) & "/" & VBA.Mid(data, 3, 2) & "/" & ano
    ElseIf VBA.Len(data) = 8 Then
       data = VBA.Mid(data, 1, 2) & "/" & VBA.Mid(data, 3, 2) & "/" & VBA.Mid(data, 5, 4)
    End If
    
    If VBA.IsDate(data) = False Then
        Erro ("Data inválida!")
        FormataData = Empty
        Exit Function
    End If
    
    FormataData = VBA.Format(data, "dd/mm/yyyy")
    
End Function

Sub Erro(Erro As String)
    On Error Resume Next
    Application.StatusBar = Erro
    Beep
    Call Rotina("00:00:05", "LimparMensagem")
End Sub

Sub Pesquisa(Por As String)
   FrmPesquisa.lblTitulo.Caption = Por
   FrmPesquisa.Show
End Sub

Sub FormataMoeda(controle As Control)
    
On Error GoTo Erro
    
    If controle.Text = Empty Then Exit Sub
    
    If VBA.IsNumeric(controle.Text) = False Then
        Cod.Erro ("Valor inválido!")
        controle.Text = Empty
        Exit Sub
    End If
    
    controle.Text = VBA.Format(controle.Text, "###,##0.00")
    
Erro:
    
End Sub

Public Function DATAAbreviada(ByVal data As String) As Date

   If data = Empty Then Exit Function
   
   DATAAbreviada = VBA.Format(data, "dd/mm/yyyy")
   
End Function

Sub ValidarNumero(ByVal TeclaPressionada As MSForms.ReturnInteger, controle As Control)

    Select Case TeclaPressionada
        
        Case 48 To 57, 44
            
            If TeclaPressionada = 44 Then
                
                If VBA.InStr(1, controle.Text, ",") > 0 Then TeclaPressionada = 0
            
            End If
        
        Case Else
            TeclaPressionada = 0
        
    End Select

End Sub

'FIM FW 04 =======================================


'INICIO FW 05 ====================================

Sub FormatarCEP(ByVal TeclaPressionada As MSForms.ReturnInteger, controle As Control)

    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim CEP As String
            CEP = controle.Text
            
            Dim X As Integer
            X = VBA.Len(CEP)
            
            If X = 10 Then TeclaPressionada = 0
            
            If X = 2 Then CEP = CEP & "."
            If X = 6 Then CEP = CEP & "-"
            
            controle.Text = CEP
            
        Case Else
            TeclaPressionada = 0
    End Select
    
End Sub

'FIM FW 05    ====================================


'INICIO FW 06 ====================================
'Macro para formatar campos do tipo CPF
Sub FormatarCPF(ByVal TeclaPressionada As MSForms.ReturnInteger, controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim CPF As String
            CPF = controle.Text
            
            Dim X As Integer
            X = VBA.Len(CPF)
            
            If X = 14 Then TeclaPressionada = 0
            
            If X = 3 Or X = 7 Then CPF = CPF & "."
            If X = 11 Then CPF = CPF & "-"
            
            controle.Text = CPF
            
        Case Else
            TeclaPressionada = 0
    End Select
    
End Sub

'FIM FW 06    ====================================


'INICIO FW 07 ====================================
'Macro para formatar campos do tipo CNPJ
Sub FormatarCNPJ(ByVal TeclaPressionada As MSForms.ReturnInteger, controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim CNPJ As String
            CNPJ = controle.Text
            
            Dim X As Integer
            X = VBA.Len(CNPJ)
            
            If X = 18 Then TeclaPressionada = 0
            
            If X = 2 Or X = 6 Then CNPJ = CNPJ & "."
            If X = 10 Then CNPJ = CNPJ & "/"
            If X = 15 Then CNPJ = CNPJ & "-"
            
            controle.Text = CNPJ
            
        Case Else
            TeclaPressionada = 0
    End Select
    
End Sub
'FIM FW 07    ====================================


'INICIO FW 08 ====================================
'Macro para formatar campos do tipo CELULAR
Sub FormatarCelular(ByVal TeclaPressionada As MSForms.ReturnInteger, controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim Celular As String
            Celular = controle.Text
            
            Dim X As Integer
            X = VBA.Len(Celular)
            
            If X = 16 Then TeclaPressionada = 0
            
            If X = 0 Then Celular = Celular & "("
            If X = 3 Then Celular = Celular & ") "
            If X = 6 Then Celular = Celular & " "
            If X = 11 Then Celular = Celular & "-"
            
            controle.Text = Celular
            
        Case Else
            
            TeclaPressionada = 0
    
    End Select
    
End Sub
'FIM FW 08    ====================================


'INICIO FW 09 ====================================
'Macro para formatar campos do tipo Telefone no formato abaixo
'(00) 0000-0000
Sub FormatarTelefone(ByVal TeclaPressionada As MSForms.ReturnInteger, controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57

            Dim Telefone As String
            Telefone = controle.Text
            
            Dim X As Integer
            X = VBA.Len(Telefone)
            
            If X = 14 Then TeclaPressionada = 0
            
            If X = 0 Then Telefone = Telefone & "("
            If X = 3 Then Telefone = Telefone & ") "
            If X = 9 Then Telefone = Telefone & "-"
            
            controle.Text = Telefone
            
        Case Else
            TeclaPressionada = 0
    
    End Select
    
End Sub
'FIM FW 09    ====================================


'INICIO FW 10    ====================================
'Macro para desabilitar alguns recursos do Excel deixando-o com uma cara de executável
Sub TelaMenu()
    
    With Application
        
        .ScreenUpdating = False
        .EnableEvents = False
        
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .Caption = "Programe aqui"
        
        With ActiveWindow
            
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayHeadings = False
            .DisplayWorkbookTabs = False
            .DisplayGridlines = False
            
        End With
        
        .ScreenUpdating = True
        .EnableEvents = True
        
    End With
    
End Sub
'FIM FW 10    ====================================


'INICIO FW 11    ====================================
'Macro para voltar o Excel as configurações padrões de exibição
Sub TelaNormal()
    
    With Application
        
        .ScreenUpdating = False
        .EnableEvents = False
        
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .Caption = Empty
        
        With ActiveWindow
            
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayHeadings = True
            .DisplayWorkbookTabs = True
            .DisplayGridlines = True
            
        End With
        
        .ScreenUpdating = True
        .EnableEvents = True
        
    End With
    
End Sub
'FIM FW 11    ====================================


'INICIO FW 12    =======================================
'Macro para verificar campos vazios em formulários
Function ValidarCampos(Controles As Controls) As Boolean
    
On Error Resume Next
    
    Dim controle As Control
    
    For Each controle In Controles
        
        If controle.Tag <> Empty Then
        
            If controle.Text = Empty Then
                
                Erro ("O campo " & controle.Tag & " é obrigatório!")
                controle.SetFocus
                ValidarCampos = True
                Exit Function
                
            End If

        End If
        
    Next controle
    
    ValidarCampos = False
    
End Function
'FIM FW 12        =======================================


Sub LimparMensagem()
    Application.StatusBar = Empty
End Sub

'INICIO FW 13     =======================================
Sub OK(Msg As String)
    Application.StatusBar = Msg
    Call Rotina("00:00:05", "LimparMensagem")
End Sub
'FIM FW 13        =======================================


'INICIO FW 14     =======================================
Sub Perguntar(Msg As String)
    
    Resposta = False
    
    FrmPergunta.LblMensagem.Caption = Msg
    FrmPergunta.Show
    
End Sub
'FIM FW 14        =======================================


'INICIO FW 15     =======================================
Sub SalvarPDF(plan As String, caminho As String, nomeARQUIVO As String)
    
On Error GoTo Erro
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ChDir caminho
    
    Sheets(plan).Visible = True
    
    Sheets(plan).ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        caminho & "\" & nomeARQUIVO & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
     
    Sheets(plan).Visible = False
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Exit Sub
Erro:
    
    Sheets(plan).Visible = False
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Cod.Erro ("Não foi possível salvar o arquivo, verifique se o mesmo já está aberto!")
    
End Sub
'FIM FW 15        =======================================


'INICIO FW 16     =======================================
'Função para abrir a caixa de seleção de arquivos
Function SelecionaArquivo(Optional FILTRO As String = "", Optional Extensao As String = "", _
Optional Titulo As String = "", Optional Email As Boolean = False) As String
    
    Dim Caixa As FileDialog
    
    Set Caixa = Application.FileDialog(msoFileDialogOpen)
    
    With Caixa
        
        .InitialView = msoFileDialogViewDetails
        
        .InitialFileName = "C:\"
        
        .AllowMultiSelect = Email
        
        .Title = Titulo
        
        If FILTRO <> Empty Then
            .Filters.Clear
            .Filters.Add FILTRO, Extensao
        End If
        
    End With
    
    Caixa.Show
    
    SelecionaArquivo = ""
    
    On Error Resume Next
        SelecionaArquivo = Caixa.SelectedItems(1)
    
End Function
'FIM FW 16        =======================================


'INICIO FW 17     =======================================
Sub CriarPasta(pasta)
    
    If VBA.Dir(pasta, vbDirectory) = "" Then
        
        Shell ("cmd /c mkdir """ & pasta & """")
        
    End If
    
End Sub
'FIM    FW 17     =======================================


'INICIO FW 18     =======================================
Sub LimparPlanilha(Aba As String, ColI As String, LinI As Integer, ColF As String)
    
    With Sheets(Aba)
        
        If .Range(ColI & LinI).value = Empty Then Exit Sub
        
        Dim lin As Integer
        lin = .Range(ColI & ":" & ColI).Find("", .Range(ColI & LinI)).Row
        
        .Range(ColI & LinI & ":" & ColF & lin - 1).ClearContents
    
    End With
    
End Sub
'FIM    FW 18     =======================================


'INICIO FW 19     =======================================
Sub Borda(Aba As String, ColI As String, LinI As Integer, ColF As String, Estilo As Integer)
    
    With Sheets(Aba)
            
        If .Range(ColI & LinI).value = "" Then Exit Sub
        
        Dim lin As Integer
        lin = .Range(ColI & ":" & ColI).Find("", .Range(ColI & LinI)).Row
        
        .Range(ColI & (LinI - 1) & ":" & ColF & lin - 1).Borders.LineStyle = Estilo
        
    End With
    
End Sub
'FIM    FW 19     =======================================

Sub LimparCampos(frm As UserForm, Optional NoClear As Control, Optional Foco As Control)

   Dim campo As Variant
   
   For Each campo In frm.Controls
   
      If TypeName(campo) = "TextBox" Or TypeName(campo) = "ComboBox" Then
         
         If NoClear Is Nothing Then
            campo.value = Empty
         Else
    
            If campo.value <> NoClear Then
               campo.value = Empty
            End If
         
         End If
         
      End If
   
   Next campo
   
   If Not Foco Is Nothing Then
      Foco.SetFocus
   End If

End Sub

Sub TirarBorda(frm As UserForm)

Dim controle As Control

   For Each controle In frm.Controls
      If TypeName(controle) = "Label" Then
         If controle.Name Like "*btn*" Then
            controle.BorderStyle = fmBorderStyleNone
         End If
      End If
   Next

End Sub

Public Function BuscarMes(mes As Variant) As Variant

   If VBA.IsNumeric(mes) = True Then
      SQL = "SELECT MES FROM DATAS WHERE CÓDIGO = " & mes
   Else
      SQL = "SELECT CÓDIGO FROM DATAS WHERE MES = '" & mes & "'"
   End If

   If Buscar = False Then Exit Function
   
   BuscarMes = resultado(0, 0)

End Function

Sub AtualizaTelaLogin()
   
   If ActiveSheet.Name <> "Login" Then Exit Sub
   
   Dim LarguraExcel As Single
   LarguraExcel = Application.Width
   
   Columns("A:A").ColumnWidth = ((LarguraExcel / 2) * 0.1805499276411) - 27

End Sub

Sub OrganizarColunaListView(ByVal frm As UserForm, ByVal col As Integer)

   With frm.lista
      
      .SortKey = col - 1
      If .SortOrder = lvwAscending Then
         .SortOrder = lvwDescending
      Else
         .SortOrder = lvwAscending
      End If
      .Sorted = True
      
   End With

End Sub


