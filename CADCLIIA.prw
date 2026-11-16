#include "totvs.ch"

/*/{Protheus.doc} CADCLIIA
Importacao e cadastro em lote de clientes (SA1 / AI0) a partir de arquivo CSV,
utilizando a rotina automatica MVC CRMA980 na operacao de inclusao (3), com
mensagens de validacao amigaveis e tela de resumo estruturada em grid.

@type    User Function
@author  Diego Alves Santana
@since   16/11/2026
@version 12.1.26.10
@see     https://tdn.totvs.com/display/public/PROT/DT+Novo+Fonte+de+Cadastro+de+Clientes+em+MVC
/*/
User Function CADCLIIA()
    Local cArqCsv    := ""
    Local aLinhas    := {}
    Local aMapaCol   := {}
    Local aDicCampos := {}
    Local aResultado := {}

    cArqCsv := SelecionarArquivo()
    If Empty(cArqCsv)
        Return Nil
    EndIf

    aLinhas := LerArquivoCsv(cArqCsv)
    If Len(aLinhas) < 2
        MsgAlert("O arquivo selecionado nao possui cabecalho e registros validos para importacao.", "CADCLIIA")
        Return Nil
    EndIf

    aMapaCol := MapearCabecalho(aLinhas[1])
    If Len(aMapaCol) == 0
        MsgAlert("Nao foi possivel identificar colunas validas da SA1 no cabecalho do arquivo.", "CADCLIIA")
        Return Nil
    EndIf

    aDicCampos := ObterDicionarioCampos()

    FWMsgRun(, {|| aResultado := ProcessarLote(aLinhas, aMapaCol, aDicCampos)}, "Importacao de Clientes", "Processando registros via CRMA980...")

    ExibirResumo(aResultado)

Return Nil

/*/{Protheus.doc} ObterPastaDocumentos
Retorna o caminho da pasta Documentos do usuario logado no SmartClient,
com fallback seguro para o diretorio raiz.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@return  Character, Caminho do diretorio terminado com barra
/*/
Static Function ObterPastaDocumentos()
    Local cDirInicial := ""
    Local cUserProf   := AllTrim(GetEnv("USERPROFILE"))

    If !Empty(cUserProf)
        cDirInicial := cUserProf + "\Documents\"
        If !ExistDir(cDirInicial)
            cDirInicial := cUserProf + "\"
        EndIf
    EndIf

    If Empty(cDirInicial) .Or. !ExistDir(cDirInicial)
        cDirInicial := "C:\"
    EndIf

Return cDirInicial

/*/{Protheus.doc} SelecionarArquivo
Abre janela de selecao de arquivo CSV no SmartClient iniciando na pasta
Documentos do usuario local.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@return  Character, Caminho completo do arquivo selecionado ou string vazia
/*/
Static Function SelecionarArquivo()
    Local cMascara    := "Arquivos CSV (*.csv)|*.csv|Todos os Arquivos (*.*)|*.*"
    Local cDirInicial := ObterPastaDocumentos()
    Local cArquivo    := ""

    cArquivo := cGetFile(cMascara, "Selecione o arquivo CSV de clientes", 1, cDirInicial, .T., GETF_LOCALHARD + GETF_NETWORKDRIVE, .F.)
    If !Empty(cArquivo)
        cArquivo := AllTrim(cArquivo)
    EndIf

Return cArquivo

/*/{Protheus.doc} LerArquivoCsv
Realiza a leitura das linhas do arquivo CSV utilizando FWFileReader,
ignorando linhas completamente vazias.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cArqCsv, Character, Caminho completo do arquivo CSV
@return  Array, Lista de linhas nao vazias lidas do arquivo
/*/
Static Function LerArquivoCsv(cArqCsv)
    Local oFile   := Nil
    Local aLinhas := {}
    Local cLinha  := ""

    oFile := FWFileReader():New(cArqCsv)
    If oFile:Open()
        While oFile:HasLine()
            cLinha := oFile:GetLine()
            If !Empty(AllTrim(cLinha))
                AAdd(aLinhas, cLinha)
            EndIf
        EndDo
        oFile:Close()
    Else
        MsgStop("Falha ao abrir o arquivo: " + cArqCsv, "CADCLIIA")
    EndIf

    FreeObj(oFile)
    oFile := Nil

Return aLinhas

/*/{Protheus.doc} MapearCabecalho
Mapeia as posicoes das colunas permitidas da SA1 a partir da primeira linha do CSV.
Descarta explicitamente a coluna A1_COD caso esteja presente no arquivo.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cCabecalho, Character, Primeira linha do arquivo CSV
@return  Array, Array bidimensional no formato {{cNomeCampo, nPosicaoColuna}, ...}
/*/
Static Function MapearCabecalho(cCabecalho)
    Local aColunas    := {}
    Local aMapa       := {}
    Local aPermitidos := {"A1_LOJA", "A1_NOME", "A1_NREDUZ", "A1_TIPO", "A1_PESSOA", "A1_EST", "A1_MUN", "A1_CGC", "A1_END", "A1_BAIRRO", "A1_CEP", "A1_EMAIL"}
    Local nIdx        := 0
    Local cCampo      := ""

    // Remove possivel BOM UTF-8 no inicio do arquivo
    If Left(cCabecalho, 3) == Chr(239) + Chr(187) + Chr(191)
        cCabecalho := SubStr(cCabecalho, 4)
    EndIf

    aColunas := StrTokArr2(cCabecalho, ";", .T.)

    For nIdx := 1 To Len(aColunas)
        cCampo := Upper(AllTrim(aColunas[nIdx]))

        // Descarta A1_COD pois a numeracao e automatica
        If cCampo == "A1_COD"
            Loop
        EndIf

        If AScan(aPermitidos, {|x| x == cCampo}) > 0
            AAdd(aMapa, {cCampo, nIdx})
        EndIf
    Next nIdx

Return aMapa

/*/{Protheus.doc} ObterDicionarioCampos
Consulta o dicionario de dados (SX3) para obter tamanho maximo (TamSX3)
e titulo amigavel (RetTitle) de cada campo da SA1 utilizado na importacao.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@return  Array, Array no formato {{cNomeCampo, nTamanhoMaximo, cRotuloAmigavel}, ...}
/*/
Static Function ObterDicionarioCampos()
    Local aCampos  := {"A1_LOJA", "A1_NOME", "A1_NREDUZ", "A1_TIPO", "A1_PESSOA", "A1_EST", "A1_MUN", "A1_CGC", "A1_END", "A1_BAIRRO", "A1_CEP", "A1_EMAIL"}
    Local aDic     := {}
    Local aTamSX3  := {}
    Local nIdx     := 0
    Local nTam     := 0
    Local cTitulo  := ""

    For nIdx := 1 To Len(aCampos)
        aTamSX3 := TamSX3(aCampos[nIdx])
        nTam    := 0
        If Len(aTamSX3) > 0
            nTam := aTamSX3[1]
        EndIf

        cTitulo := AllTrim(RetTitle(aCampos[nIdx]))
        If Empty(cTitulo)
            cTitulo := aCampos[nIdx]
        Else
            cTitulo := cTitulo + " (" + aCampos[nIdx] + ")"
        EndIf

        AAdd(aDic, {aCampos[nIdx], nTam, cTitulo})
    Next nIdx

Return aDic

/*/{Protheus.doc} ObterNomeAmigavel
Retorna o nome amigavel do campo (ex.: "Razao Social (A1_NOME)") a partir do
dicionario carregado em memoria ou consulta dinamica na SX3.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cCampo, Character, Nome tecnico do campo (ex.: A1_CGC)
@param   aDicCampos, Array, Dicionario de campos carregado previamente
@return  Character, Descricao amigavel acompanhada do nome do campo
/*/
Static Function ObterNomeAmigavel(cCampo, aDicCampos)
    Local cNomeFmt := Upper(AllTrim(cCampo))
    Local nPos     := 0
    Local cTitSX3  := ""

    nPos := AScan(aDicCampos, {|x| x[1] == cNomeFmt})
    If nPos > 0
        Return aDicCampos[nPos][3]
    EndIf

    cTitSX3 := AllTrim(RetTitle(cNomeFmt))
    If !Empty(cTitSX3)
        cNomeFmt := cTitSX3 + " (" + cNomeFmt + ")"
    EndIf

Return cNomeFmt

/*/{Protheus.doc} ProcessarLote
Itera sobre as linhas de dados do CSV, realiza a pre-validacao e executa o MsExecAuto
da rotina CRMA980 (operacao 3 - inclusao) com complemento AI0 (AI0_SALDO = 0).

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aLinhas, Array, Linhas lidas do arquivo CSV (incluindo cabecalho na posicao 1)
@param   aMapaCol, Array, Mapeamento de colunas gerado por MapearCabecalho
@param   aDicCampos, Array, Tamanhos maximos e titulos amigaveis dos campos
@return  Array, Estrutura {nIncluidos, nErros, aDetalhes}
/*/
Static Function ProcessarLote(aLinhas, aMapaCol, aDicCampos)
    Local nLinha      := 0
    Local nCol        := 0
    Local nPosCol     := 0
    Local nIncluidos  := 0
    Local nErros      := 0
    Local aColunas    := {}
    Local aDadosSA1   := {}
    Local aDadosAI0   := {}
    Local aDetalhes   := {}
    Local cCampo      := ""
    Local cValor      := ""
    Local cErroValid  := ""
    Local cErroExec   := ""
    Local cFilSA1     := FWxFilial("SA1")
    Local cFilAI0     := FWxFilial("AI0")
    Private lMsErroAuto    := .F.
    Private lMsHelpAuto    := .T.
    Private lAutoErrNoFile := .T.

    For nLinha := 2 To Len(aLinhas)
        aColunas  := StrTokArr2(aLinhas[nLinha], ";", .T.)
        aDadosSA1 := {}
        aDadosAI0 := {}

        AAdd(aDadosSA1, {"A1_FILIAL", cFilSA1, Nil})

        For nCol := 1 To Len(aMapaCol)
            cCampo  := aMapaCol[nCol][1]
            nPosCol := aMapaCol[nCol][2]
            cValor  := ""

            If nPosCol <= Len(aColunas)
                cValor := AllTrim(aColunas[nPosCol])
            EndIf

            If cCampo $ "A1_TIPO|A1_PESSOA|A1_EST"
                cValor := Upper(cValor)
            EndIf

            If cCampo $ "A1_CGC|A1_CEP"
                cValor := StrTran(cValor, ".", "")
                cValor := StrTran(cValor, "-", "")
                cValor := StrTran(cValor, "/", "")
            EndIf

            AAdd(aDadosSA1, {cCampo, cValor, Nil})
        Next nCol

        cErroValid := ValidarRegistro(aDadosSA1, aDicCampos)
        If !Empty(cErroValid)
            nErros++
            AAdd(aDetalhes, {.F., nLinha, "ERRO VALIDACAO", cErroValid})
            Loop
        EndIf

        // Complemento do cliente (AI0) com saldo zerado conforme documentacao CRMA980
        AAdd(aDadosAI0, {"AI0_FILIAL", cFilAI0, Nil})
        AAdd(aDadosAI0, {"AI0_SALDO",  0,       Nil})

        lMsErroAuto := .F.
        MsExecAuto({|a, b, c| CRMA980(a, b, c)}, aDadosSA1, 3, aDadosAI0)

        If lMsErroAuto
            nErros++
            cErroExec := FormatarErroExecAuto(aDicCampos)
            FWLogMsg("WARN", , "CADCLIIA", "ProcessarLote", "", "01", "Linha " + cValToChar(nLinha) + ": " + cErroExec, 0, 0, {})
            AAdd(aDetalhes, {.F., nLinha, "ERRO CADASTRO", cErroExec})
        Else
            nIncluidos++
            AAdd(aDetalhes, {.T., nLinha, "INCLUIDO", "Cliente cadastrado com sucesso (Codigo: " + SA1->A1_COD + " / Loja: " + SA1->A1_LOJA + " - " + AllTrim(SA1->A1_NOME) + ")."})
        EndIf
    Next nLinha

Return {nIncluidos, nErros, aDetalhes}

/*/{Protheus.doc} ValidarRegistro
Valida campos obrigatorios, dominio de A1_TIPO e A1_PESSOA e tamanho maximo (SX3)
antes da chamada ao ExecAuto, gerando mensagens amigaveis com o rotulo do campo.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aDadosSA1, Array, Array de campos e valores montado para a SA1
@param   aDicCampos, Array, Tamanhos maximos e titulos amigaveis obtidos da SX3
@return  Character, Mensagem consolidada de inconsistencias ou string vazia se valido
/*/
Static Function ValidarRegistro(aDadosSA1, aDicCampos)
    Local aObrig    := {"A1_LOJA", "A1_NOME", "A1_NREDUZ", "A1_TIPO", "A1_PESSOA", "A1_EST", "A1_MUN", "A1_CGC"}
    Local aMensag   := {}
    Local nIdx      := 0
    Local nPos      := 0
    Local nTamMax   := 0
    Local cCampo    := ""
    Local cRotulo   := ""
    Local cValor    := ""
    Local cRetorno  := ""

    // 1. Validacao de obrigatoriedade
    For nIdx := 1 To Len(aObrig)
        cCampo := aObrig[nIdx]
        nPos   := AScan(aDadosSA1, {|x| x[1] == cCampo})
        If nPos == 0 .Or. Empty(aDadosSA1[nPos][2])
            cRotulo := ObterNomeAmigavel(cCampo, aDicCampos)
            AAdd(aMensag, "O campo " + cRotulo + " e obrigatorio e nao foi preenchido.")
        EndIf
    Next nIdx

    // 2. Validacao de dominio A1_TIPO (F, L, R, S, X)
    nPos := AScan(aDadosSA1, {|x| x[1] == "A1_TIPO"})
    If nPos > 0 .And. !Empty(aDadosSA1[nPos][2])
        cValor := aDadosSA1[nPos][2]
        If !(cValor $ "F|L|R|S|X")
            cRotulo := ObterNomeAmigavel("A1_TIPO", aDicCampos)
            AAdd(aMensag, "Valor invalido ('" + cValor + "') no campo " + cRotulo + ". Escolha: F (Cons. Final), L (Prod. Rural), R (Revendedor), S (Solidario) ou X (Exportacao).")
        EndIf
    EndIf

    // 3. Validacao de dominio A1_PESSOA (J, F)
    nPos := AScan(aDadosSA1, {|x| x[1] == "A1_PESSOA"})
    If nPos > 0 .And. !Empty(aDadosSA1[nPos][2])
        cValor := aDadosSA1[nPos][2]
        If !(cValor $ "J|F")
            cRotulo := ObterNomeAmigavel("A1_PESSOA", aDicCampos)
            AAdd(aMensag, "Valor invalido ('" + cValor + "') no campo " + cRotulo + ". Escolha: J (Juridica) ou F (Fisica).")
        EndIf
    EndIf

    // 4. Validacao de tamanho maximo conforme SX3
    For nIdx := 1 To Len(aDadosSA1)
        cCampo := aDadosSA1[nIdx][1]
        If cCampo == "A1_FILIAL"
            Loop
        EndIf

        cValor := aDadosSA1[nIdx][2]
        nPos   := AScan(aDicCampos, {|x| x[1] == cCampo})
        If nPos > 0
            nTamMax := aDicCampos[nPos][2]
            cRotulo := aDicCampos[nPos][3]
            If nTamMax > 0 .And. Len(cValor) > nTamMax
                AAdd(aMensag, "O campo " + cRotulo + " ultrapassou o limite de " + cValToChar(nTamMax) + " caracteres (foram informados " + cValToChar(Len(cValor)) + ").")
            EndIf
        EndIf
    Next nIdx

    For nIdx := 1 To Len(aMensag)
        If nIdx > 1
            cRetorno += CRLF + "- "
        Else
            cRetorno += "- "
        EndIf
        cRetorno += aMensag[nIdx]
    Next nIdx

Return cRetorno

/*/{Protheus.doc} FormatarErroExecAuto
Captura o log de erro gerado em memoria pelo ExecAuto MVC via GetAutoGRLog(),
extrai o campo responsavel, a mensagem de erro e a solucao sugerida,
traduzindo o identificador tecnico para o titulo amigavel do campo.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aDicCampos, Array, Dicionario de campos para traducoes amigaveis
@return  Character, Mensagem de erro amigavel para exibicao ao usuario
/*/
Static Function FormatarErroExecAuto(aDicCampos)
    Local aLog      := {}
    Local nIdx      := 0
    Local cLinha    := ""
    Local cCampoId  := ""
    Local cMsgErro  := ""
    Local cSolucao  := ""
    Local cFallback := ""
    Local cErroFmt  := ""

    aLog := GetAutoGRLog()

    For nIdx := 1 To Len(aLog)
        cLinha := AllTrim(aLog[nIdx])
        If Empty(cLinha)
            Loop
        EndIf

        If At("ID DO CAMPO:", Upper(cLinha)) > 0
            cCampoId := AllTrim(SubStr(cLinha, At(":", cLinha) + 1))
        ElseIf At("MENSAGEM DO ERRO:", Upper(cLinha)) > 0
            cMsgErro := AllTrim(SubStr(cLinha, At(":", cLinha) + 1))
        ElseIf At("MENSAGEM DA SOLUCAO:", Upper(cLinha)) > 0 .Or. At("MENSAGEM DA SOLUÇÃO:", Upper(cLinha)) > 0
            cSolucao := AllTrim(SubStr(cLinha, At(":", cLinha) + 1))
        ElseIf !(Upper(Left(cLinha, 16)) == "ID DO SUBMODELO:") .And. ;
               !(Upper(Left(cLinha, 15)) == "ID DO FORMULARIO:") .And. ;
               !(Upper(Left(cLinha, 11)) == "ID DO ERRO:") .And. ;
               !(Upper(Left(cLinha, 17)) == "VALOR ATRIBUIDO:") .And. ;
               !(Upper(Left(cLinha, 15)) == "VALOR ANTERIOR:") .And. ;
               !(Left(cLinha, 3) == "---")
            If !Empty(cFallback)
                cFallback += " "
            EndIf
            cFallback += cLinha
        EndIf
    Next nIdx

    If !Empty(cCampoId)
        cErroFmt += "Campo: " + ObterNomeAmigavel(cCampoId, aDicCampos) + " - "
    EndIf

    If !Empty(cMsgErro)
        cErroFmt += cMsgErro
        If !Empty(cSolucao)
            cErroFmt += " (Sugestao: " + cSolucao + ")"
        EndIf
    ElseIf !Empty(cFallback)
        cErroFmt += cFallback
    Else
        cErroFmt := "Nao foi possivel concluir o cadastro deste cliente. Verifique os dados informados na linha."
    EndIf

Return "- " + cErroFmt

/*/{Protheus.doc} ExibirResumo
Apresenta uma interface visual estruturada com indicadores de totais,
grid zebrado com alto contraste e semaforo por linha, painel de detalhe
completo do registro selecionado e opcao para exportar o relatorio.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aResultado, Array, Estrutura {nIncluidos, nErros, aDetalhes}
@return  Nil
/*/
Static Function ExibirResumo(aResultado)
    Local oDlg        := Nil
    Local oBrowse     := Nil
    Local oDetalhe    := Nil
    Local oBtnExport  := Nil
    Local oBtnFechar  := Nil
    Local oFontBold   := TFont():New("Arial", , -14, , .T.)
    Local oFontNorm   := TFont():New("Arial", , -12, , .F.)
    Local oVerde      := LoadBitmap(GetResources(), "BR_VERDE")
    Local oVermelho   := LoadBitmap(GetResources(), "BR_VERMELHO")
    Local nIncluidos  := aResultado[1]
    Local nErros      := aResultado[2]
    Local aDetalhes   := aResultado[3]
    Local cDetAtual   := ""

    If Len(aDetalhes) > 0
        cDetAtual := "Linha " + StrZero(aDetalhes[1][2], 4) + " [" + aDetalhes[1][3] + "]:" + CRLF + aDetalhes[1][4]
    EndIf

    Define MsDialog oDlg Title "Resultado da Importacao de Clientes - CADCLIIA" From 000, 000 To 560, 920 Pixel

    // Painel superior com totalizadores visuais
    @ 008, 010 To 042, 150 Label " Total Processado " Of oDlg Pixel
    @ 020, 020 Say cValToChar(nIncluidos + nErros) + " registro(s)" Size 120, 015 Font oFontBold Color CLR_BLACK Of oDlg Pixel

    @ 008, 160 To 042, 305 Label " Incluidos com Sucesso " Of oDlg Pixel
    @ 020, 170 Say cValToChar(nIncluidos) + " cliente(s)" Size 125, 015 Font oFontBold Color CLR_GREEN Of oDlg Pixel

    @ 008, 315 To 042, 450 Label " Registros com Erro " Of oDlg Pixel
    @ 020, 325 Say cValToChar(nErros) + " erro(s)" Size 115, 015 Font oFontBold Color CLR_RED Of oDlg Pixel

    // Grid principal orientado a objetos (TWBrowse) com alto contraste
    oBrowse := TWBrowse():New(048, 010, 440, 145, , {"", "Linha CSV", "Situacao", "Resumo da Ocorrencia"}, {18, 42, 75, 295}, oDlg, , , , , , , , , , , , .F., , .T., , .F., , , )
    oBrowse:SetArray(aDetalhes)
    oBrowse:bLine := {|| { ;
        ObterIconeStatus(aDetalhes[oBrowse:nAt][1], oVerde, oVermelho), ;
        "Linha " + StrZero(aDetalhes[oBrowse:nAt][2], 4), ;
        aDetalhes[oBrowse:nAt][3], ;
        StrTran(aDetalhes[oBrowse:nAt][4], CRLF, " ") ;
    }}
    oBrowse:lUseDefaultColors := .F.
    oBrowse:SetBlkColor({|| RGB(20, 20, 20)})
    oBrowse:SetBlkBackColor({|| ObterCorLinha(oBrowse:nAt)})
    oBrowse:bChange := {|| ;
        cDetAtual := "Linha " + StrZero(aDetalhes[oBrowse:nAt][2], 4) + " [" + aDetalhes[oBrowse:nAt][3] + "]:" + CRLF + aDetalhes[oBrowse:nAt][4], ;
        oDetalhe:Refresh() ;
    }

    // Painel inferior com o detalhe completo da linha selecionada (quebra de linha limpa)
    @ 198, 010 Say "Detalhe completo da linha selecionada:" Size 200, 010 Font oFontNorm Of oDlg Pixel
    @ 208, 010 Get oDetalhe Var cDetAtual Memo Size 440, 048 Of oDlg Pixel ReadOnly

    // Botoes de acao
    @ 262, 305 Button oBtnExport Prompt "Exportar Log (.txt)" Size 070, 015 Of oDlg Pixel Action ExportarLogTxt(nIncluidos, nErros, aDetalhes)
    @ 262, 380 Button oBtnFechar Prompt "Fechar" Size 070, 015 Of oDlg Pixel Action oDlg:End()

    Activate MsDialog oDlg Centered

    oBrowse:Destroy()
    oFontBold:Destroy()
    oFontNorm:Destroy()

Return Nil

/*/{Protheus.doc} ObterIconeStatus
Retorna o bitmap de semaforo (verde ou vermelho) conforme o sucesso da linha.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   lSucesso, Logical, Indica se a linha foi incluida com sucesso
@param   oVerde, Object, Bitmap verde carregado
@param   oVermelho, Object, Bitmap vermelho carregado
@return  Object, Bitmap correspondente ao status
/*/
Static Function ObterIconeStatus(lSucesso, oVerde, oVermelho)
    Local oRet := oVermelho

    If lSucesso
        oRet := oVerde
    EndIf

Return oRet

/*/{Protheus.doc} ObterCorLinha
Alterna a cor de fundo das linhas do grid (efeito zebrado suave) para alto contraste.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   nLinha, Numeric, Indice atual da linha no grid
@return  Numeric, Codigo RGB da cor de fundo
/*/
Static Function ObterCorLinha(nLinha)
    Local nCor := RGB(255, 255, 255)

    If Mod(nLinha, 2) == 0
        nCor := RGB(238, 243, 250)
    EndIf

Return nCor

/*/{Protheus.doc} ExportarLogTxt
Exporta o resumo completo e detalhado para um arquivo texto na pasta Documentos.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   nIncluidos, Numeric, Total de clientes incluidos
@param   nErros, Numeric, Total de registros com erro
@param   aDetalhes, Array, Lista detalhada de ocorrencias por linha
@return  Nil
/*/
Static Function ExportarLogTxt(nIncluidos, nErros, aDetalhes)
    Local cDirDest  := ObterPastaDocumentos()
    Local cArqLog   := cDirDest + "log_importacao_clientes_" + DToS(Date()) + "_" + StrTran(Time(), ":", "") + ".txt"
    Local oWriter   := Nil
    Local nIdx      := 0
    Local cTexto    := ""

    cTexto += "==================================================================================" + CRLF
    cTexto += " RESUMO DA IMPORTACAO DE CLIENTES (CRMA980)" + CRLF
    cTexto += " Data/Hora: " + DToC(Date()) + " " + Time() + CRLF
    cTexto += "==================================================================================" + CRLF
    cTexto += " Total de registros processados : " + cValToChar(nIncluidos + nErros) + CRLF
    cTexto += " Incluidos com sucesso          : " + cValToChar(nIncluidos) + CRLF
    cTexto += " Registros com erro             : " + cValToChar(nErros) + CRLF
    cTexto += "==================================================================================" + CRLF + CRLF

    For nIdx := 1 To Len(aDetalhes)
        cTexto += "----------------------------------------------------------------------------------" + CRLF
        cTexto += "Linha " + StrZero(aDetalhes[nIdx][2], 4) + " | Situacao: " + aDetalhes[nIdx][3] + CRLF
        cTexto += aDetalhes[nIdx][4] + CRLF
    Next nIdx

    oWriter := FWFileWriter():New(cArqLog, .T.)
    If oWriter:Create()
        oWriter:Write(cTexto)
        oWriter:Close()
        MsgInfo("Log exportado com sucesso em:" + CRLF + cArqLog, "Log Exportado")
    Else
        MsgStop("Nao foi possivel salvar o arquivo de log em: " + cArqLog, "CADCLIIA")
    EndIf

    FreeObj(oWriter)
    oWriter := Nil

Return Nil
