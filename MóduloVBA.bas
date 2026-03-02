Attribute VB_Name = "ReportGenerator"

' =================================================================
' 1. MACROS DE INTERFACE (BOTÕES)
' =================================================================
Sub Macro_Generate_Individual()
    Dim FundoParaGerar As String
    FundoParaGerar = InputBox("Digite o nome da aba do fundo que deseja gerar:", "Gerar Individual")
    If FundoParaGerar <> "" Then Call Motor_Gerador_Relatorio(FundoParaGerar)
End Sub

Sub Macro_Generate_All_In_Mass()
    Dim wbMain As Workbook: Dim ws As Worksheet: Dim Caminho As String
    Caminho = ThisWorkbook.Path & "\"
    
    On Error Resume Next
    ' Nome alterado para Main_Database.xlsm para o GitHub
    Set wbMain = Workbooks("Main_Database.xlsm")
    If wbMain Is Nothing Then Set wbMain = Workbooks.Open(Caminho & "Main_Database.xlsm", ReadOnly:=True)
    On Error GoTo 0
    
    For Each ws In wbMain.Worksheets
        ' Ignora abas de sistema/configuração
        If ws.Name <> "Capa" And ws.Name <> "Config" And ws.Name <> "Instruções" Then
            Call Motor_Gerador_Relatorio(ws.Name)
            DoEvents
        End If
    Next ws
    MsgBox "Processamento concluído!", vbInformation
End Sub

' =================================================================
' 2. MOTOR PRINCIPAL
' =================================================================
Sub Motor_Gerador_Relatorio(ByVal FundoNome As String)
    Dim WordApp As Object, WordDoc As Object, Caminho As String, CNPJ As String, PeriodoRelatorio As String
    Dim wbApont As Workbook, wbEstrut As Workbook, wbMain As Workbook
    Dim wsFundoDados As Worksheet
    Dim resumoTexto As String, NomeArquivoSaida As String
    
    Caminho = ThisWorkbook.Path & "\"
    PeriodoRelatorio = "PERIODO_RELATORIO" ' Alterar para o trimestre/ano desejado
    
    On Error Resume Next
    ' Carregamento das bases com nomes genéricos
    Set wbApont = Workbooks("Findings_Data.xlsm"): If wbApont Is Nothing Then Set wbApont = Workbooks.Open(Caminho & "Findings_Data.xlsm", ReadOnly:=True)
    Set wbEstrut = Workbooks("Structural_Data.xlsm"): If wbEstrut Is Nothing Then Set wbEstrut = Workbooks.Open(Caminho & "Structural_Data.xlsm", ReadOnly:=True)
    Set wbMain = Workbooks("Main_Database.xlsm"): If wbMain Is Nothing Then Set wbMain = Workbooks.Open(Caminho & "Main_Database.xlsm", ReadOnly:=True)
    Set wsFundoDados = wbMain.Sheets(FundoNome)
    On Error GoTo 0
    
    If wsFundoDados Is Nothing Then Exit Sub
    CNPJ = wsFundoDados.Range("A2").Value

    ' Inicia o Word e abre o Template genérico
    Set WordApp = CreateObject("Word.Application"): WordApp.Visible = False
    Set WordDoc = WordApp.Documents.Add(Template:=Caminho & "Report_Template.docx")

    ' --- Preenchimento de Tags Globais ---
    ColocarTextoDireto WordDoc, "[QUADROS_NOME]", UCase(wsFundoDados.Range("A1").Value)
    ColocarTextoDireto WordDoc, "[QUADROS_CNPJ]", CNPJ
    ColocarTextoDireto WordDoc, "[QUADROS_TRIMESTRE]", PeriodoRelatorio
    ColocarTextoDireto WordDoc, "[QUADROS_TIPO]", UCase(wsFundoDados.Range("C2").Value), True
    
    ColocarTextoDireto WordDoc, "[QUADROS_INVEST]", wsFundoDados.Range("D2").Value
    resumoTexto = wsFundoDados.Range("A6").Value
    ColocarTextoDireto WordDoc, "[QUADROS_RESUMO]", resumoTexto
    ColocarTextoDireto WordDoc, "[QUADROS_EXTENSO]", ExtrairApenasParte1(resumoTexto)
    
    ' Data base genérica para o GitHub
    ColocarTextoDireto WordDoc, "[DATA_BASE]", "DD/MM/YYYY"

    ' --- Extrações de Bases Externas ---
    ColocarTextoDireto WordDoc, "[ESTRUT_3.2]", BuscarDados(wbEstrut.Sheets("3.2"), CNPJ, 2, 14)
    ColocarTextoDireto WordDoc, "[ESTRUT_3.3.2]", BuscarDados(wbEstrut.Sheets("3.3.2"), CNPJ, 4, 14)

    ' --- Lógica de Apontamentos com Verificação de Conteúdo ---
    Dim txt33Ap As String: txt33Ap = BuscarDados(wbApont.Sheets("3.3"), CNPJ, 1, 10, False)
    If InStr(1, txt33Ap, "não foram identificados", vbTextCompare) > 0 Then
        ColocarTextoDireto WordDoc, "[APONT_3.3_FRASE]", "não constatamos divergências."
        ApagarBlocoSeguro WordDoc, "[APONT_3.3]", "Estes itens citados estão dispostos no Anexo I."
    Else
        ColocarTextoDireto WordDoc, "[APONT_3.3_FRASE]", "apresentamos o quadro a seguir:"
        ColocarTextoDireto WordDoc, "[APONT_3.3]", txt33Ap
    End If

    ' --- Lógica de Seção 3.4/3.5 ---
    Dim txt34 As String: txt34 = BuscarDados(wbApont.Sheets("3.4"), CNPJ, 2, 8)
    If InStr(1, txt34, "não foram identificados", vbTextCompare) > 0 Then
        ApagarBlocoSeguro WordDoc, "Constatações:", "relacionados."
    Else
        ColocarTextoDireto WordDoc, "[APONT_3.5]", txt34
    End If

    ' --- Inserção e Formatação de Tabelas ---
    ' Estoque: Centralizado
    ColarTabelaWord WordDoc, "[QUADROS_TABELA_ESTOQUE]", wsFundoDados.Range("A20:D28"), True
    
    ' Amostra: Alinhamento misto (Esquerda/Centro)
    ColarTabelaWord WordDoc, "[QUADROS_TABELA_AMOSTRA]", wsFundoDados.Range("A12:D17"), False
    
    ' Carteira: Verificação de valores zerados antes da inserção
    If Application.WorksheetFunction.CountA(wsFundoDados.Range("B32:B34")) = 0 Or _
       (Val(wsFundoDados.Range("B32").Value) = 0 And Val(wsFundoDados.Range("B33").Value) = 0 And Val(wsFundoDados.Range("B34").Value) = 0) Then
        ApagarBlocoSeguro WordDoc, "Posição Carteira", "[QUADROS_TABELA_CARTEIRA]"
    Else
        ColarTabelaWord WordDoc, "[QUADROS_TABELA_CARTEIRA]", wsFundoDados.Range("A31:C35"), False
    End If

    ' --- Salvamento e Finalização ---
    NomeArquivoSaida = Caminho & PeriodoRelatorio & " - " & FundoNome & ".docx"
    
    On Error Resume Next
    WordDoc.SaveAs2 NomeArquivoSaida
    DoEvents
    Application.Wait (Now + TimeValue("00:00:02")) ' Pausa de segurança para sincronização
    
    WordDoc.Close SaveChanges:=True
    WordApp.Quit
    
    Set WordDoc = Nothing: Set WordApp = Nothing
    Application.CutCopyMode = False
End Sub

' =================================================================
' 3. FUNÇÕES DE APOIO
' =================================================================

Sub ColocarTextoDireto(Doc As Object, Tag As String, Texto As String, Optional ForcarNegrito As Boolean = False)
    Dim Rng As Object: Set Rng = Doc.Content
    With Rng.Find
        .Text = Tag: .Forward = True: .Wrap = 1
        Do While .Execute
            If Texto = "" Then
                Rng.Delete
            Else
                Rng.Text = Texto
                If ForcarNegrito Then Rng.Font.Bold = True
                If InStr(1, Texto, Chr(149)) > 0 Then
                    With Rng.ParagraphFormat
                        .LeftIndent = Doc.Application.CentimetersToPoints(1.27)
                        .FirstLineIndent = Doc.Application.CentimetersToPoints(-0.64)
                    End With
                End If
            End If
            Rng.Collapse 0
        Loop
    End With
End Sub

Sub ColarTabelaWord(Doc As Object, Tag As String, RngExcel As Range, TudoCentralizado As Boolean)
    Dim RngWord As Object: Set RngWord = Doc.Content
    With RngWord.Find
        If .Execute(FindText:=Tag) Then
            RngExcel.Copy: Application.Wait (Now + TimeValue("00:00:01"))
            RngWord.PasteExcelTable False, False, False
            
            If Doc.Tables.Count > 0 Then
                Dim Tbl As Object: Set Tbl = Doc.Tables(Doc.Tables.Count)
                Dim r As Long, TextoB As String, TextoC As String
                
                ' Deleta linhas zeradas (Amostra e Carteira)
                If Tag = "[QUADROS_TABELA_AMOSTRA]" Or Tag = "[QUADROS_TABELA_CARTEIRA]" Then
                    For r = Tbl.Rows.Count - 1 To 2 Step -1
                        TextoB = CleanWordCell(Tbl.Cell(r, 2).Range.Text)
                        TextoC = CleanWordCell(Tbl.Cell(r, 3).Range.Text)
                        If (TextoB = "" Or TextoB = "0" Or TextoB = "-") And _
                           (TextoC = "" Or TextoC = "0" Or TextoC = "-") Then
                            Tbl.Rows(r).Delete
                        End If
                    Next r
                End If
                
                ' Ajustes estéticos e de alinhamento
                Dim TotalLinhas As Long: TotalLinhas = Tbl.Rows.Count
                With Tbl
                    .AllowAutoFit = True
                    .AutoFitBehavior (2) ' Autoajuste à janela
                    .Range.Font.Size = 9
                    
                    ' Negritos em cabeçalhos e totais
                    .Rows(1).Range.Font.Bold = True
                    .Rows(TotalLinhas).Range.Font.Bold = True
                    
                    If TudoCentralizado Then
                        .Range.ParagraphFormat.Alignment = 1
                    Else
                        Dim c As Long
                        For c = 1 To .Columns.Count
                            If c = 1 Then
                                .Columns(c).Select
                                Doc.Application.Selection.ParagraphFormat.Alignment = 0 ' Esquerda para descrições
                            Else
                                .Columns(c).Select
                                Doc.Application.Selection.ParagraphFormat.Alignment = 1 ' Centro para números
                            End If
                        Next c
                    End If
                End With
            End If
        End If
    End With
End Sub

Function CleanWordCell(ByVal txt As String) As String
    ' Remove caracteres de controle de tabela do Word
    txt = Replace(txt, Chr(7), "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(10), "")
    txt = Replace(txt, Chr(160), "")
    txt = Trim(txt)
    If txt = " - " Then txt = "-"
    CleanWordCell = txt
End Function

Function BuscarDados(Aba As Worksheet, ValorBusca As String, ColID As Integer, ColSaida As Integer, Optional ComBullet As Boolean = False) As String
    If Aba Is Nothing Then Exit Function
    Dim i As Long, Res As String, Encontrou As Boolean, Contador As Long: Contador = 0
    Dim ID As String: ID = Trim(Replace(Replace(ValorBusca, ".", ""), "-", ""))
    Dim ValText As String
    
    For i = 2 To Aba.Cells(Aba.Rows.Count, ColID).End(xlUp).Row
        If Trim(Replace(Replace(CStr(Aba.Cells(i, ColID).Value), ".", ""), "-", "")) = ID Then
            ValText = Trim(CStr(Aba.Cells(i, ColSaida).Value))
            If Len(ValText) > 3 Then
                Contador = Contador + 1: Encontrou = True
                If Contador > 1 Then Res = Res & vbCrLf & (IIf(ComBullet, Chr(149) & " ", "")) & ValText Else Res = ValText
            End If
        End If
    Next i
    If Encontrou Then BuscarDados = Res Else BuscarDados = "não foram identificados apontamentos para este item."
End Function

Sub ApagarBlocoSeguro(Doc As Object, Inicio As String, Fim As String)
    ' Apaga blocos de texto entre dois marcadores com segurança de parágrafo
    Dim Rng As Object: Set Rng = Doc.Content: Dim StartPos As Long, EndPos As Long
    If Rng.Find.Execute(FindText:=Inicio) Then
        StartPos = Rng.Start: Set Rng = Doc.Content
        If Rng.Find.Execute(FindText:=Fim) Then
            EndPos = Rng.End
            Do While EndPos < Doc.Range.End
                If Doc.Range(EndPos, EndPos + 1).Text = vbCr Or Doc.Range(EndPos, EndPos + 1).Text = Chr(13) Or Doc.Range(EndPos, EndPos + 1).Text = " " Then
                    EndPos = EndPos + 1
                Else: Exit Do: End If
            Loop
            Doc.Range(StartPos, EndPos).Delete
        End If
    End If
End Sub

Function ExtrairApenasParte1(ByVal Texto As String) As String
    Dim s As String: s = Trim(Texto)
    If InStr(1, s, "não foram", vbTextCompare) > 0 Or s = "" Then Exit Function
    
    ' Remove prefixos comuns de amostragem
    If InStr(1, s, "A amostra total de ", vbTextCompare) > 0 Then
        s = Mid(s, InStr(1, s, "A amostra total de ", vbTextCompare) + Len("A amostra total de "))
    End If
    
    If InStr(1, s, "amostra total de ", vbTextCompare) > 0 Then
        s = Mid(s, InStr(1, s, "amostra total de ", vbTextCompare) + Len("amostra total de "))
    End If

    Dim p As Long: p = InStr(1, s, ")")
    If p > 0 Then ExtrairApenasParte1 = Trim(Left(s, p)) Else ExtrairApenasParte1 = Trim(s)
End Function
