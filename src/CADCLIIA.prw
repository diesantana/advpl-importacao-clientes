#include "totvs.ch"

/*/{Protheus.doc} CADCLIIA
Importacao e cadastro em lote de clientes (SA1 / AI0) diretamente a partir
de planilha Excel (.xlsx), utilizando a rotina automatica MVC CRMA980 na
operacao de inclusao (3).

@type    User Function
@author  Diego Alves Santana
@since   16/11/2026
@version 12.1.26.10
@see     https://tdn.totvs.com/display/public/PROT/DT+Novo+Fonte+de+Cadastro+de+Clientes+em+MVC
/*/
User Function CADCLIIA()
    Local cArqXlsx   := ""
    Local aLinhas    := {}
    Local aMapaCol   := {}
    Local aDicCampos := {}
    Local aResultado := {}

    cArqXlsx := SelecionarArquivo()
    If Empty(cArqXlsx)
        Return Nil
    EndIf

    FWMsgRun(, {|| aLinhas := LerArquivoExcel(cArqXlsx)}, "Leitura da Planilha", "Lendo dados do arquivo Excel...")

    If Len(aLinhas) < 2
        MsgAlert("A planilha selecionada nao possui cabecalho e registros validos para importacao.", "CADCLIIA")
        Return Nil
    EndIf

    aMapaCol := MapearCabecalho(aLinhas[1])
    If Len(aMapaCol) == 0
        MsgAlert("Nao foi possivel identificar colunas validas da SA1 no cabecalho da planilha.", "CADCLIIA")
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
Abre janela de selecao de planilha Excel (.xlsx) no SmartClient iniciando
na pasta Documentos do usuario local.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@return  Character, Caminho completo do arquivo selecionado ou string vazia
/*/
Static Function SelecionarArquivo()
    Local cMascara    := "Planilhas Excel (*.xlsx)|*.xlsx|Todos os Arquivos (*.*)|*.*"
    Local cDirInicial := ObterPastaDocumentos()
    Local cArquivo    := ""

    cArquivo := cGetFile(cMascara, "Selecione a planilha Excel de clientes", 1, cDirInicial, .T., GETF_LOCALHARD + GETF_NETWORKDRIVE, .F.)
    If !Empty(cArquivo)
        cArquivo := AllTrim(cArquivo)
    EndIf

Return cArquivo

/*/{Protheus.doc} LerArquivoExcel
Transfere temporariamente o arquivo .xlsx do SmartClient para o servidor,
descompacta via FUnZip() e realiza a leitura direta dos XMLs internos
(sharedStrings.xml e sheet1.xml), garantindo extracao exata de todas as celulas.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cArqLocal, Character, Caminho completo do arquivo .xlsx na estacao cliente
@return  Array, Matriz bidimensional {{nNumLinhaExcel, aValoresColunas}, ...}
/*/
Static Function LerArquivoExcel(cArqLocal)
    Local aLinhas     := {}
    Local aSharedStr  := {}
    Local aRow        := {}
    Local cIdUnico    := FWUUIDV4()
    Local cDirBase    := "\temp_cadcliia\"
    Local cDirExt     := cDirBase + cIdUnico + "\"
    Local cArqSrv     := cDirBase + "imp_" + cIdUnico + ".xlsx"
    Local cXmlSheet   := ""
    Local cConteudo   := ""
    Local cBlocoRow   := ""
    Local cTagRow     := ""
    Local cBlocoCell  := ""
    Local cTagCell    := ""
    Local cRefCelula  := ""
    Local cTipoCelula := ""
    Local cValorBruto := ""
    Local cValorFinal := ""
    Local nPosRowIni  := 0
    Local nPosRowFim  := 0
    Local nPosTagFim  := 0
    Local nPosCellIni := 0
    Local nPosCellFim := 0
    Local nNumLinha   := 0
    Local nColIdx     := 0
    Local nStrIdx     := 0
    Local lTemDados   := .F.

    If !ExistDir(cDirBase)
        MakeDir(cDirBase)
    EndIf

    If !CpyT2S(cArqLocal, cDirBase, .F.)
        MsgStop("Nao foi possivel transferir a planilha para o servidor Protheus.", "CADCLIIA")
        Return aLinhas
    EndIf

    If FRename(cDirBase + ExtrairNomeArquivo(cArqLocal), cArqSrv) <> 0
        cArqSrv := cDirBase + ExtrairNomeArquivo(cArqLocal)
    EndIf

    MakeDir(cDirExt)

    If FUnZip(cArqSrv, cDirExt) <> 0
        MsgStop("Falha ao descompactar a planilha Excel (.xlsx). Verifique se o arquivo nao esta corrompido.", "CADCLIIA")
        LimparTemporarios(cArqSrv, cDirExt)
        Return aLinhas
    EndIf

    aSharedStr := CarregarSharedStrings(cDirExt + "xl\sharedStrings.xml")
    cXmlSheet  := LocalizarXmlAba(cDirExt, "CLIENTES")

    If Empty(cXmlSheet) .Or. !File(cXmlSheet)
        MsgStop("Nao foi possivel localizar os dados da aba de clientes dentro do arquivo .xlsx.", "CADCLIIA")
        LimparTemporarios(cArqSrv, cDirExt)
        Return aLinhas
    EndIf

    cConteudo := LerTextoArquivo(cXmlSheet)

    While .T.
        nPosRowIni := At("<row ", cConteudo)
        If nPosRowIni == 0
            nPosRowIni := At("<row>", cConteudo)
        EndIf
        If nPosRowIni == 0
            Exit
        EndIf

        nPosRowFim := At("</row>", SubStr(cConteudo, nPosRowIni))
        If nPosRowFim == 0
            Exit
        EndIf

        cBlocoRow  := SubStr(cConteudo, nPosRowIni, nPosRowFim + 5)
        cConteudo  := SubStr(cConteudo, nPosRowIni + nPosRowFim + 5)

        nPosTagFim := At(">", cBlocoRow)
        cTagRow    := Left(cBlocoRow, nPosTagFim)
        nNumLinha  := Val(ExtrairAtributoXml(cTagRow, "r"))

        aRow      := Array(15)
        AFill(aRow, "")
        lTemDados := .F.

        While .T.
            nPosCellIni := At("<c ", cBlocoRow)
            If nPosCellIni == 0
                nPosCellIni := At("<c>", cBlocoRow)
            EndIf
            If nPosCellIni == 0
                Exit
            EndIf

            nPosTagFim := At(">", SubStr(cBlocoRow, nPosCellIni))
            If nPosTagFim == 0
                Exit
            EndIf

            cTagCell := SubStr(cBlocoRow, nPosCellIni, nPosTagFim)

            // Ignora tags auto-fechadas sem valor (ex: <c r="H3" s="4"/>)
            If Right(AllTrim(cTagCell), 2) == "/>"
                cBlocoRow := SubStr(cBlocoRow, nPosCellIni + nPosTagFim)
                Loop
            EndIf

            nPosCellFim := At("</c>", SubStr(cBlocoRow, nPosCellIni))
            If nPosCellFim == 0
                Exit
            EndIf

            cBlocoCell := SubStr(cBlocoRow, nPosCellIni, nPosCellFim + 3)
            cBlocoRow  := SubStr(cBlocoRow, nPosCellIni + nPosCellFim + 3)

            cRefCelula  := ExtrairAtributoXml(cTagCell, "r")
            cTipoCelula := Lower(AllTrim(ExtrairAtributoXml(cTagCell, "t")))
            nColIdx     := ColunaParaIndice(cRefCelula)
            cValorFinal := ""

            If cTipoCelula == "s"
                cValorBruto := AllTrim(ExtrairValorTagXml(cBlocoCell, "v"))
                If !Empty(cValorBruto)
                    nStrIdx := Val(cValorBruto) + 1
                    If nStrIdx >= 1 .And. nStrIdx <= Len(aSharedStr)
                        cValorFinal := aSharedStr[nStrIdx]
                    EndIf
                EndIf
            ElseIf cTipoCelula == "inlinestr"
                cValorFinal := ExtrairTodosTextosTagT(cBlocoCell)
            Else
                cValorFinal := ExtrairValorTagXml(cBlocoCell, "v")
                If Empty(cValorFinal)
                    cValorFinal := ExtrairTodosTextosTagT(cBlocoCell)
                EndIf
            EndIf

            cValorFinal := DecodificarEntidadesXml(cValorFinal)

            If nColIdx > Len(aRow)
                ASize(aRow, nColIdx)
            EndIf

            If nColIdx >= 1
                aRow[nColIdx] := cValorFinal
                If !Empty(AllTrim(cValorFinal))
                    lTemDados := .T.
                EndIf
            EndIf
        EndDo

        If lTemDados
            AAdd(aLinhas, {nNumLinha, aRow})
        EndIf
    EndDo

    LimparTemporarios(cArqSrv, cDirExt)

Return aLinhas

/*/{Protheus.doc} CarregarSharedStrings
Le o arquivo xl/sharedStrings.xml diretamente como texto e extrai todos
os itens <si> na ordem exata de indice.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cArqShared, Character, Caminho completo do sharedStrings.xml
@return  Array, Lista de strings em ordem de indice (base 1 no AdvPL)
/*/
Static Function CarregarSharedStrings(cArqShared)
    Local aStrings  := {}
    Local cConteudo := ""
    Local cBlocoSi  := ""
    Local cTexto    := ""
    Local nPosIni   := 0
    Local nPosFim   := 0

    If !File(cArqShared)
        Return aStrings
    EndIf

    cConteudo := LerTextoArquivo(cArqShared)

    While .T.
        nPosIni := At("<si>", cConteudo)
        If nPosIni == 0
            nPosIni := At("<si ", cConteudo)
        EndIf
        If nPosIni == 0
            Exit
        EndIf

        nPosFim := At("</si>", SubStr(cConteudo, nPosIni))
        If nPosFim == 0
            Exit
        EndIf

        cBlocoSi  := SubStr(cConteudo, nPosIni, nPosFim + 4)
        cConteudo := SubStr(cConteudo, nPosIni + nPosFim + 4)

        cTexto := ExtrairTodosTextosTagT(cBlocoSi)
        cTexto := DecodificarEntidadesXml(cTexto)

        AAdd(aStrings, cTexto)
    EndDo

Return aStrings

/*/{Protheus.doc} LocalizarXmlAba
Identifica o arquivo XML correspondente a aba desejada (ex: 'Clientes')
lendo xl/workbook.xml, com fallback para sheet1.xml.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cDirExt, Character, Diretorio raiz onde o .xlsx foi extraido
@param   cNomeAba, Character, Nome da aba procurada (em maiusculo)
@return  Character, Caminho completo do arquivo sheetN.xml
/*/
Static Function LocalizarXmlAba(cDirExt, cNomeAba)
    Local cWorkbook := cDirExt + "xl\workbook.xml"
    Local cSheetXml := cDirExt + "xl\worksheets\sheet1.xml"
    Local cConteudo := ""
    Local cTagSheet := ""
    Local cNome     := ""
    Local nPosIni   := 0
    Local nPosFim   := 0
    Local nIdx      := 0

    If File(cWorkbook)
        cConteudo := LerTextoArquivo(cWorkbook)
        While .T.
            nPosIni := At("<sheet ", cConteudo)
            If nPosIni == 0
                Exit
            EndIf

            nPosFim := At(">", SubStr(cConteudo, nPosIni))
            If nPosFim == 0
                Exit
            EndIf

            nIdx++
            cTagSheet := SubStr(cConteudo, nPosIni, nPosFim)
            cConteudo := SubStr(cConteudo, nPosIni + nPosFim)

            cNome := Upper(AllTrim(ExtrairAtributoXml(cTagSheet, "name")))
            If cNome == cNomeAba
                If File(cDirExt + "xl\worksheets\sheet" + cValToChar(nIdx) + ".xml")
                    cSheetXml := cDirExt + "xl\worksheets\sheet" + cValToChar(nIdx) + ".xml"
                EndIf
                Exit
            EndIf
        EndDo
    EndIf

Return cSheetXml

/*/{Protheus.doc} LerTextoArquivo
Le o conteudo completo de um arquivo XML no servidor por blocos binarios
(FOpen/FRead), evitando truncamento em arquivos XML minificados de linha unica.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cArquivo, Character, Caminho do arquivo no servidor
@return  Character, Conteudo completo do arquivo
/*/
Static Function LerTextoArquivo(cArquivo)
    Local nHandle   := -1
    Local nTamArq   := 0
    Local nLidos    := 0
    Local nBytes    := 0
    Local cBuffer   := ""
    Local cConteudo := ""

    nHandle := FOpen(cArquivo, 0)
    If nHandle >= 0
        nTamArq := FSeek(nHandle, 0, 2)
        FSeek(nHandle, 0, 0)

        While nLidos < nTamArq
            cBuffer := ""
            nBytes  := FRead(nHandle, @cBuffer, 32768)
            If nBytes <= 0
                Exit
            EndIf
            cConteudo += cBuffer
            nLidos    += nBytes
        EndDo

        FClose(nHandle)
    EndIf

Return cConteudo

/*/{Protheus.doc} ExtrairAtributoXml
Extrai o valor de um atributo XML (ex: r="F3") dentro de uma tag de abertura.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cTag, Character, Tag XML de abertura
@param   cAtributo, Character, Nome do atributo
@return  Character, Valor do atributo
/*/
Static Function ExtrairAtributoXml(cTag, cAtributo)
    Local cBusca := " " + cAtributo + '="'
    Local nPos   := At(cBusca, cTag)
    Local nFim   := 0
    Local cValor := ""

    If nPos > 0
        cValor := SubStr(cTag, nPos + Len(cBusca))
        nFim   := At('"', cValor)
        If nFim > 0
            cValor := Left(cValor, nFim - 1)
        EndIf
    EndIf

Return cValor

/*/{Protheus.doc} ExtrairValorTagXml
Extrai o conteudo interno de uma tag XML simples (ex: <v>103</v>).

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cBloco, Character, Trecho XML
@param   cNomeTag, Character, Nome da tag (ex: "v")
@return  Character, Conteudo interno da tag
/*/
Static Function ExtrairValorTagXml(cBloco, cNomeTag)
    Local nIni   := At("<" + cNomeTag + ">", cBloco)
    Local nFim   := 0
    Local cValor := ""

    If nIni > 0
        cValor := SubStr(cBloco, nIni + Len(cNomeTag) + 2)
        nFim   := At("</" + cNomeTag + ">", cValor)
        If nFim > 0
            cValor := Left(cValor, nFim - 1)
        EndIf
    EndIf

Return cValor

/*/{Protheus.doc} ExtrairTodosTextosTagT
Concatena o conteudo de todas as ocorrencias de <t> ou <t ...> dentro de um bloco XML.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cBloco, Character, Trecho XML (<si> ou <is>)
@return  Character, Texto concatenado
/*/
Static Function ExtrairTodosTextosTagT(cBloco)
    Local cTexto  := ""
    Local cResto  := cBloco
    Local nIni    := 0
    Local nTagFim := 0
    Local nFim    := 0

    While .T.
        nIni := At("<t>", cResto)
        If nIni == 0
            nIni := At("<t ", cResto)
        EndIf
        If nIni == 0
            Exit
        EndIf

        nTagFim := At(">", SubStr(cResto, nIni))
        If nTagFim == 0
            Exit
        EndIf

        cResto := SubStr(cResto, nIni + nTagFim)
        nFim   := At("</t>", cResto)
        If nFim == 0
            Exit
        EndIf

        cTexto += Left(cResto, nFim - 1)
        cResto := SubStr(cResto, nFim + 4)
    EndDo

Return cTexto

/*/{Protheus.doc} DecodificarEntidadesXml
Converte entidades XML padrao (&amp;, &lt;, &gt;, &quot;, &apos;) e caracteres UTF-8
para texto CP-1252, tratando com seguranca retornos Nil de DecodeUTF8().

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cTexto, Character, Texto extraido do XML
@return  Character, Texto decodificado
/*/
Static Function DecodificarEntidadesXml(cTexto)
    Local xDec := Nil
    Local cRet := cTexto

    If !Empty(cTexto)
        xDec := DecodeUTF8(cTexto)
        If ValType(xDec) == "C" .And. !Empty(xDec)
            cRet := xDec
        EndIf
    EndIf

    cRet := StrTran(cRet, "&amp;",  "&")
    cRet := StrTran(cRet, "&lt;",   "<")
    cRet := StrTran(cRet, "&gt;",   ">")
    cRet := StrTran(cRet, "&quot;", '"')
    cRet := StrTran(cRet, "&apos;", "'")

Return cRet

/*/{Protheus.doc} ColunaParaIndice
Converte a referencia de celula do Excel (ex: "A2", "D10", "AA5") no indice
numerico da coluna (A=1, B=2, ..., Z=26, AA=27).

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cRef, Character, Referencia alfanumerica da celula
@return  Numeric, Indice numerico da coluna (base 1)
/*/
Static Function ColunaParaIndice(cRef)
    Local cLetras := ""
    Local nIdx    := 0
    Local nCol    := 0
    Local cChar   := ""

    cRef := Upper(AllTrim(cRef))
    For nIdx := 1 To Len(cRef)
        cChar := SubStr(cRef, nIdx, 1)
        If cChar >= "A" .And. cChar <= "Z"
            cLetras += cChar
        Else
            Exit
        EndIf
    Next nIdx

    For nIdx := 1 To Len(cLetras)
        nCol := (nCol * 26) + (Asc(SubStr(cLetras, nIdx, 1)) - 64)
    Next nIdx

Return nCol

/*/{Protheus.doc} LimparTemporarios
Remove o arquivo .xlsx temporario e a pasta descompactada no servidor.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cArqSrv, Character, Caminho do arquivo .xlsx temporario
@param   cDirExt, Character, Caminho da pasta extraida
@return  Nil
/*/
Static Function LimparTemporarios(cArqSrv, cDirExt)
    Local aSubDirs := {"_rels\", "docProps\", "xl\_rels\", "xl\theme\", "xl\worksheets\_rels\", "xl\worksheets\", "xl\", ""}
    Local aFiles   := {}
    Local nDir     := 0
    Local nFile    := 0
    Local cPasta   := ""

    If File(cArqSrv)
        FErase(cArqSrv)
    EndIf

    If ExistDir(cDirExt)
        For nDir := 1 To Len(aSubDirs)
            cPasta := cDirExt + aSubDirs[nDir]
            If ExistDir(cPasta)
                aFiles := Directory(cPasta + "*.*")
                For nFile := 1 To Len(aFiles)
                    If aFiles[nFile][1] <> "." .And. aFiles[nFile][1] <> ".."
                        FErase(cPasta + aFiles[nFile][1])
                    EndIf
                Next nFile
                If !Empty(aSubDirs[nDir])
                    DirRemove(cPasta)
                EndIf
            EndIf
        Next nDir
        DirRemove(cDirExt)
    EndIf

Return Nil

/*/{Protheus.doc} ExtrairNomeArquivo
Extrai apenas o nome do arquivo com extensao a partir de um caminho completo.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   cCaminho, Character, Caminho completo do arquivo
@return  Character, Nome do arquivo
/*/
Static Function ExtrairNomeArquivo(cCaminho)
    Local nPosBarra := RAt("\", cCaminho)
    Local cNome     := cCaminho

    If nPosBarra > 0
        cNome := SubStr(cCaminho, nPosBarra + 1)
    EndIf

Return cNome

/*/{Protheus.doc} MapearCabecalho
Mapeia as posicoes das colunas permitidas da SA1 a partir da primeira linha
da planilha, suportando tanto "A1_CAMPO" quanto "Descricao (UF) * (A1_CAMPO)".
Descarta explicitamente a coluna A1_COD caso esteja presente.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aLinhaCab, Array, Estrutura {nLinhaExcel, aColunasCabecalho}
@return  Array, Array bidimensional no formato {{cNomeCampo, nPosicaoColuna}, ...}
/*/
Static Function MapearCabecalho(aLinhaCab)
    Local aColunas    := aLinhaCab[2]
    Local aMapa       := {}
    Local aPermitidos := {"A1_LOJA", "A1_NOME", "A1_NREDUZ", "A1_TIPO", "A1_PESSOA", "A1_EST", "A1_MUN", "A1_CGC", "A1_END", "A1_BAIRRO", "A1_CEP", "A1_EMAIL"}
    Local nIdx        := 0
    Local nPosAbre    := 0
    Local nPosFecha   := 0
    Local cCampo      := ""

    For nIdx := 1 To Len(aColunas)
        cCampo := Upper(AllTrim(aColunas[nIdx]))

        // Usa RAt() para extrair o ULTIMO parentese (ex: "ESTADO (UF) * (A1_EST)" -> "A1_EST")
        nPosAbre  := RAt("(", cCampo)
        nPosFecha := RAt(")", cCampo)
        If nPosAbre > 0 .And. nPosFecha > nPosAbre
            cCampo := AllTrim(SubStr(cCampo, nPosAbre + 1, nPosFecha - nPosAbre - 1))
        EndIf

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
Consulta o dicionario de dados (SX3) para obter tamanho maximo (TamSX3),
titulo amigavel (RetTitle) e obrigatoriedade real da base (X3Obrigat)
de cada campo da SA1 utilizado na importacao.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@return  Array, Array no formato {{cNomeCampo, nTamanhoMaximo, cRotuloAmigavel, lObrigatorio}, ...}
/*/
Static Function ObterDicionarioCampos()
    Local aCampos  := {"A1_LOJA", "A1_NOME", "A1_NREDUZ", "A1_TIPO", "A1_PESSOA", "A1_EST", "A1_MUN", "A1_CGC", "A1_END", "A1_BAIRRO", "A1_CEP", "A1_EMAIL"}
    Local aDic     := {}
    Local aTamSX3  := {}
    Local nIdx     := 0
    Local nTam     := 0
    Local cTitulo  := ""
    Local lObrig   := .F.

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

        lObrig := X3Obrigat(aCampos[nIdx])

        AAdd(aDic, {aCampos[nIdx], nTam, cTitulo, lObrig})
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
Itera sobre as linhas de dados da planilha, realiza a pre-validacao, ordena
os campos conforme precedencia de gatilhos e executa o MsExecAuto da rotina
CRMA980 (operacao 3 - inclusao) com complemento AI0.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aLinhas, Array, Linhas lidas da planilha (incluindo cabecalho na posicao 1)
@param   aMapaCol, Array, Mapeamento de colunas gerado por MapearCabecalho
@param   aDicCampos, Array, Tamanhos maximos, titulos amigaveis e obrigatoriedade SX3
@return  Array, Estrutura {nIncluidos, nErros, aDetalhes}
/*/
Static Function ProcessarLote(aLinhas, aMapaCol, aDicCampos)
    Local nIdxLinha   := 0
    Local nNumLinha   := 0
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

    For nIdxLinha := 2 To Len(aLinhas)
        nNumLinha := aLinhas[nIdxLinha][1]
        aColunas  := aLinhas[nIdxLinha][2]
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

            // Extrai apenas o codigo inicial quando selecionado na lista descritiva (ex: "F - FISICO" -> "F")
            If cCampo $ "A1_TIPO|A1_PESSOA"
                cValor := Upper(Left(cValor, 1))
            ElseIf cCampo == "A1_EST"
                cValor := Upper(Left(cValor, 2))
            EndIf

            If cCampo $ "A1_CGC|A1_CEP"
                cValor := StrTran(cValor, ".", "")
                cValor := StrTran(cValor, "-", "")
                cValor := StrTran(cValor, "/", "")
            EndIf

            // Envia ao ExecAuto apenas campos preenchidos para nao sobrescrever inicializadores padrao
            If !Empty(cValor)
                AAdd(aDadosSA1, {cCampo, cValor, Nil})
            EndIf
        Next nCol

        cErroValid := ValidarRegistro(aDadosSA1, aDicCampos)
        If !Empty(cErroValid)
            nErros++
            AAdd(aDetalhes, {.F., nNumLinha, "ERRO VALIDACAO", cErroValid})
            Loop
        EndIf

        // Ordena os campos para respeitar gatilhos e validacoes encadeadas do CRMA980
        OrdenarCamposSA1(@aDadosSA1)

        // Complemento do cliente (AI0) com saldo zerado conforme documentacao CRMA980
        AAdd(aDadosAI0, {"AI0_FILIAL", cFilAI0, Nil})
        AAdd(aDadosAI0, {"AI0_SALDO",  0,       Nil})

        lMsErroAuto := .F.
        MsExecAuto({|a, b, c| CRMA980(a, b, c)}, aDadosSA1, 3, aDadosAI0)

        If lMsErroAuto
            nErros++
            cErroExec := FormatarErroExecAuto(aDicCampos)
            FWLogMsg("WARN", , "CADCLIIA", "ProcessarLote", "", "01", "Linha " + cValToChar(nNumLinha) + ": " + cErroExec, 0, 0, {})
            AAdd(aDetalhes, {.F., nNumLinha, "ERRO CADASTRO", cErroExec})
        Else
            nIncluidos++
            AAdd(aDetalhes, {.T., nNumLinha, "INCLUIDO", "Cliente cadastrado com sucesso (Codigo: " + SA1->A1_COD + " / Loja: " + SA1->A1_LOJA + " - " + AllTrim(SA1->A1_NOME) + ")."})
        EndIf
    Next nIdxLinha

Return {nIncluidos, nErros, aDetalhes}

/*/{Protheus.doc} OrdenarCamposSA1
Reordena o array aDadosSA1 conforme a precedencia de gatilhos (SX7) e
validacoes (SX3) do modelo MVC CRMA980 (ex: A1_PESSOA antes de A1_CGC,
A1_EST antes de A1_MUN).

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aDadosSA1, Array, Array de campos e valores da SA1 (passado por referencia)
@return  Nil
/*/
Static Function OrdenarCamposSA1(aDadosSA1)
    Local aOrdem    := {"A1_FILIAL", "A1_LOJA", "A1_PESSOA", "A1_TIPO", "A1_CGC", "A1_NOME", "A1_NREDUZ", "A1_EST", "A1_MUN", "A1_END", "A1_BAIRRO", "A1_CEP", "A1_EMAIL"}
    Local aOrdenado := {}
    Local nIdx      := 0
    Local nPos      := 0

    For nIdx := 1 To Len(aOrdem)
        nPos := AScan(aDadosSA1, {|x| x[1] == aOrdem[nIdx]})
        If nPos > 0
            AAdd(aOrdenado, aDadosSA1[nPos])
        EndIf
    Next nIdx

    For nIdx := 1 To Len(aDadosSA1)
        If AScan(aOrdenado, {|x| x[1] == aDadosSA1[nIdx][1]}) == 0
            AAdd(aOrdenado, aDadosSA1[nIdx])
        EndIf
    Next nIdx

    aDadosSA1 := aOrdenado

Return Nil

/*/{Protheus.doc} ValidarRegistro
Valida obrigatoriedade real conforme X3Obrigat() da base, dominio de A1_TIPO
e A1_PESSOA e tamanho maximo (SX3) antes da chamada ao ExecAuto.

@type    Static Function
@author  Diego Alves Santana
@since   16/11/2026
@param   aDadosSA1, Array, Array de campos e valores montado para a SA1
@param   aDicCampos, Array, Tamanhos maximos, titulos amigaveis e obrigatoriedade SX3
@return  Character, Mensagem consolidada de inconsistencias ou string vazia se valido
/*/
Static Function ValidarRegistro(aDadosSA1, aDicCampos)
    Local aMensag   := {}
    Local nIdx      := 0
    Local nPos      := 0
    Local nTamMax   := 0
    Local cCampo    := ""
    Local cRotulo   := ""
    Local cValor    := ""
    Local cRetorno  := ""

    // 1. Validacao de obrigatoriedade dinamica conforme X3Obrigat() da base
    For nIdx := 1 To Len(aDicCampos)
        If aDicCampos[nIdx][4]
            cCampo  := aDicCampos[nIdx][1]
            cRotulo := aDicCampos[nIdx][3]
            nPos    := AScan(aDadosSA1, {|x| x[1] == cCampo})
            If nPos == 0 .Or. Empty(aDadosSA1[nPos][2])
                AAdd(aMensag, "O campo " + cRotulo + " e obrigatorio na base e nao foi preenchido.")
            EndIf
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
Apresenta uma interface visual estruturada com indicadores de totais compactos
e centralizados, grid zebrado via CSS nativo de alto contraste e semaforo por
linha, painel de detalhe completo e opcao para exportar o relatorio.

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
    Local oFontBold   := TFont():New("Arial", , -13, , .T.)
    Local oFontNorm   := TFont():New("Arial", , -12, , .F.)
    Local oVerde      := LoadBitmap(GetResources(), "BR_VERDE")
    Local oVermelho   := LoadBitmap(GetResources(), "BR_VERMELHO")
    Local nIncluidos  := aResultado[1]
    Local nErros      := aResultado[2]
    Local aDetalhes   := aResultado[3]
    Local cDetAtual   := ""
    Local cCssGrid    := ""

    If Len(aDetalhes) > 0
        cDetAtual := "Linha " + StrZero(aDetalhes[1][2], 4) + " [" + aDetalhes[1][3] + "]:" + CRLF + aDetalhes[1][4]
    EndIf

    cCssGrid := "QTableView {" + ;
                "  background-color: #FFFFFF;" + ;
                "  alternate-background-color: #F2F6FC;" + ;
                "  color: #111111;" + ;
                "  gridline-color: #D9E1EC;" + ;
                "  selection-background-color: #1E88E5;" + ;
                "  selection-color: #FFFFFF;" + ;
                "}"

    Define MsDialog oDlg Title "Resultado da Importacao de Clientes - CADCLIIA" From 000, 000 To 550, 920 Pixel

    // Painel superior com cards compactos e centralizados
    @ 008, 047 To 038, 162 Label " Total Processado " Of oDlg Pixel
    @ 019, 057 Say cValToChar(nIncluidos + nErros) + " registro(s)" Size 095, 012 Font oFontBold Color CLR_BLACK Of oDlg Pixel

    @ 008, 172 To 038, 287 Label " Incluidos com Sucesso " Of oDlg Pixel
    @ 019, 182 Say cValToChar(nIncluidos) + " cliente(s)" Size 095, 012 Font oFontBold Color CLR_GREEN Of oDlg Pixel

    @ 008, 297 To 038, 412 Label " Registros com Erro " Of oDlg Pixel
    @ 019, 307 Say cValToChar(nErros) + " erro(s)" Size 095, 012 Font oFontBold Color CLR_RED Of oDlg Pixel

    // Grid principal orientado a objetos (TWBrowse) estilizado via CSS
    oBrowse := TWBrowse():New(045, 010, 440, 148, , {"", "Linha Excel", "Situacao", "Resumo da Ocorrencia"}, {18, 45, 75, 292}, oDlg, , , , , , , , , , , , .F., , .T., , .F., , , )
    oBrowse:SetArray(aDetalhes)
    oBrowse:bLine := {|| { ;
        ObterIconeStatus(aDetalhes[oBrowse:nAt][1], oVerde, oVermelho), ;
        "Linha " + StrZero(aDetalhes[oBrowse:nAt][2], 4), ;
        aDetalhes[oBrowse:nAt][3], ;
        StrTran(aDetalhes[oBrowse:nAt][4], CRLF, " ") ;
    }}
    oBrowse:SetCSS(cCssGrid)
    oBrowse:bChange := {|| ;
        cDetAtual := "Linha " + StrZero(aDetalhes[oBrowse:nAt][2], 4) + " [" + aDetalhes[oBrowse:nAt][3] + "]:" + CRLF + aDetalhes[oBrowse:nAt][4], ;
        oDetalhe:Refresh() ;
    }

    // Painel inferior com o detalhe completo da linha selecionada
    @ 198, 010 Say "Detalhe completo da linha selecionada:" Size 200, 010 Font oFontNorm Of oDlg Pixel
    @ 208, 010 Get oDetalhe Var cDetAtual Memo Size 440, 046 Of oDlg Pixel ReadOnly

    // Botoes de acao
    @ 259, 305 Button oBtnExport Prompt "Exportar Log (.txt)" Size 070, 015 Of oDlg Pixel Action ExportarLogTxt(nIncluidos, nErros, aDetalhes)
    @ 259, 380 Button oBtnFechar Prompt "Fechar" Size 070, 015 Of oDlg Pixel Action oDlg:End()

    Activate MsDialog oDlg Centered

    FreeObj(oFontBold)
    FreeObj(oFontNorm)
    oFontBold := Nil
    oFontNorm := Nil
    oBrowse   := Nil

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
