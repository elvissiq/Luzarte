#INCLUDE "totvs.ch"
#INCLUDE "fwprintsetup.ch"
#INCLUDE "rptdef.ch" 
#INCLUDE "TBICONN.CH"
#INCLUDE "topconn.ch"

#DEFINE ENTER Chr(13)+Chr(10)
#define PAD_LEFT		0
#define PAD_RIGHT		1
#define PAD_CENTER   	2

/*/{Protheus.doc} LOMSR002
Relatório Fluxo Financeiro de Viagem Romaneio Luzarte
@type function
@version
@author TOTVS Nordeste
@since 22/11/2023
@return 
/*/
User Function LOMSR002()
    Local aArea    := FWGetArea()
    
    Private cBarra   := IIF(GetRemoteType() == 1,"\","/")
    Private cNomArq  := ""
    Private cDirArq  := ""
    Private oProcess
    
    If Pergunte( "LOMSR002" , .T. , "Perguntas - Relatório Fluxo Financeiro de Viagem Romaneio" )
        cDirArq := TFileDialog("Arquivos Adobe PDF (*.pdf)",'Informe onde será gravado o arquivo.',,,.T.,/*GETF_MULTISELECT*/)
        cNomArq := SubSTR(cDirArq,RAT(cBarra,cDirArq)+1,Rat(".",SubSTR(cDirArq,RAT(cBarra,cDirArq)+1))-1)
        cDirArq := SubSTR(cDirArq,1,RAT(cBarra,cDirArq))
        
        If !Empty(cDirArq) .AND. !Empty(cNomArq)
            oProcess := MsNewProcess():New({|| fLOMSR002()}, "Processando...", "Aguarde...", .T.)
            oProcess:Activate()
        EndIf 
    EndIF 
    
    FWRestArea(aArea)

Return

/*---------------------------------------------------------------------*
 | Func:  fLOMSR002                                                    |
 | Desc:  Impressão do relatório                                       |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function fLOMSR002()
    
    Local nAtual := 0
    Local nTotal := 0
    Local cQry   := ""

    Private oPrint
    Private nLin      := 0
    Private nCol      := 0
    Private nPagina   := 0
    Private oFont10   := TFont():New( "Arial",,10,,.F.,,,,,.F. )
    Private oFont10B  := TFont():New( "Arial",,10,,.T.,,,,,.F. )
    Private oFont11   := TFont():New( "Arial",,11,,.F.,,,,,.F. )
    Private oFont11B  := TFont():New( "Arial",,11,,.T.,,,,,.F. )
    Private oFont12   := TFont():New( "Arial",,12,,.F.,,,,,.F. )
    Private oFont12B  := TFont():New( "Arial",,12,,.T.,,,,,.F. )
    Private oFont14   := TFont():New( "Arial",,14,,.F.,,,,,.F. )
    Private oFont14B  := TFont():New( "Arial",,14,,.T.,,,,,.F. )
    Private oFont16   := TFont():New( "Arial",,16,,.F.,,,,,.F. )
    Private oFont16B  := TFont():New( "Arial",,16,,.T.,,,,,.F. )
    Private cAliasDAK := GetNextAlias()

    cQry := " SELECT * "
    cQry += " FROM " + RetSqlName("DAK") + " DAK " 
    cQry += "   WHERE DAK.D_E_L_E_T_ <> '*'"
    cQry += "     AND DAK.DAK_FILIAL  = '" + xFilial("DAK") + "'"
    cQry += "     AND DAK.DAK_COD  = '" + MV_PAR01 + "' "
    cQry := ChangeQuery(cQry)
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),cAliasDAK,.F.,.T.)

    Count To nTotal
    oProcess:SetRegua1(nTotal)
    
    DBSelectArea(cAliasDAK)
    (cAliasDAK)->(DbGoTop())

    If !(cAliasDAK)->(Eof())
        
        oPrint:=FWMSPrinter():New(cNomArq,IMP_PDF, .F., , .T.)
        oPrint:SetResolution(72)
        oPrint:SetLandscape()
        oPrint:SetPaperSize(DMPAPER_A4)
        oPrint:SetMargin(60,60,60,60)
        oPrint:cPathPDF := cDirArq
        
        While !(cAliasDAK)->(Eof()) 
            
            nAtual++
            oProcess:IncRegua1("Carga " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

            fCabec() //Imprime o Cabeçalho
            fNFiscais((cAliasDAK)->(DAK_COD)) //Imprime as Notas Fiscais da Carga
            
        (cAliasDAK)->(DBSkip())
        EndDo 
        
        oPrint:Preview()
        ms_flush()

    Else
        FWAlertWarning("Nenhuma informação encontrada com os parâmetros informados.",'Parâmetros da pergunta "LOMSR002"')
    EndIF

    If Select((cAliasDAK)) <> 0
		(cAliasDAK)->(DbCloseArea())
	Endif

Return

/*---------------------------------------------------------------------*
 | Func:  fCabec                                                       |
 | Desc:  Imprime o cabeçalho do relatório                             |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function fCabec()
    
    Local cNomeEmp  := Upper(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_NOMECOM" } )[1][2]))

    nLin := 35
    nCol := 5

    nPagina += 1

    oPrint:StartPage()
    oPrint:SayAlign(nLin, nCol, cNomeEmp, oFont12, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+710, "Emissão: "+DToC(dDataBase), oFont12, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
	nLin += 10
    oPrint:SayAlign(nLin, nCol+710, "Página: "+cValTochar(nPagina), oFont12, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    nLin += 10
    oPrint:SayAlign(nLin, nCol+710, "Hora...: "+Time(), oFont12, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
	nLin += 30
    oPrint:SayAlign(nLin, nCol+250, "Fluxo Financeiro de Viagem Romaneio", oFont16B, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
	nLin += 40
    oPrint:SayAlign(nLin, nCol, "Motorista: "+Alltrim(Posicione("DA4",1,FWxFilial("DA4")+(cAliasDAK)->DAK_MOTORI,"DA4_NOME")), oFont12B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+450, "Numero do Lacre: ", oFont12B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    nLin += 10
    oPrint:SayAlign(nLin, nCol, "Placa: "+Alltrim(Posicione("DA3",1,FWxFilial("DA3")+(cAliasDAK)->DAK_CAMINH,"DA3_PLACA")), oFont12B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+450, "Num. Manifesto...: ", oFont12B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    nLin += 30
    oPrint:SayAlign(nLin, nCol    , "NF"          , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+045, "DATA"        , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+090, "CNPJ/CPF"    , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+170, "RAZÃO SOCIAL", oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+300, "CIDADE"      , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+370, "UF"          , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+385, "VALOR"       , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+450, "IPI"         , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+480, "SUBSTIT."    , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+530, "RECEBIMENTO" , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+600, "DEVOLUÇÃO"   , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+680, "ACRÉSCIMO"   , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+740, "OUTROS"      , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)

    nLin += 15

Return

/*---------------------------------------------------------------------*
 | Func:  fNFiscais                                                    |
 | Desc:  Imprime as notas fiscais do relatório                        |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function fNFiscais(pCodCarga)
    Local cAliasSF2 := GetNextAlias()
    Local cQry      := ""
    Local cPictFat  := PesqPict("SF2","F2_VALFAT")
    Local cPictIPI  := PesqPict("SF2","F2_VALIPI")
    Local cPictSub  := PesqPict("SF2","F2_BSFCPST")
    Local cPictCNPJ := "@R 99.999.999/9999-99"
    Local cPictCPF  := "@R 999.999.999-99"
    Local nValTot   := 0
    Local nValIPI   := 0
    Local nValSub   := 0
    Local nTotal    := 0
    Local nAtual    := 0

    cQry := " SELECT * "
    cQry += " FROM " + RetSqlName("SF2") + " SF2 " 
    cQry += "   WHERE SF2.D_E_L_E_T_ <> '*'"
    cQry += "     AND SF2.F2_FILIAL  = '" + xFilial("SE1") + "'"
    cQry += "     AND SF2.F2_CARGA  = '" + pCodCarga + "' "
    cQry := ChangeQuery(cQry)
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),cAliasSF2,.F.,.T.)

    Count To nTotal
    oProcess:SetRegua2(nTotal)
    
    DBSelectArea(cAliasSF2)
    (cAliasSF2)->(DbGoTop())

    DBSelectArea("SA1")
    SA1->(DBSetOrder(1))

    DBSelectArea("SA2")
    SA2->(DBSetOrder(1))

    While !(cAliasSF2)->(Eof())

        nAtual++
        oProcess:IncRegua2("Nota Fiscal " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

        If nLin >= 600
            fCabec() //Imprime o Cabeçalho
        EndIf 

        If SA1->(MSSeek(FWxFilial("SA1")+(cAliasSF2)->F2_CLIENTE+(cAliasSF2)->F2_LOJA))

            oPrint:SayAlign(nLin, nCol    , (cAliasSF2)->F2_DOC                                     , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+045, DToC(SToD((cAliasSF2)->F2_EMISSAO))                     , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+090, Transform(Alltrim(SA1->A1_CGC),IIF(SA1->A1_PESSOA == 'J',cPictCNPJ,cPictCPF)) , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+170, Alltrim(SA1->A1_NOME)                                   , oFont10, 130, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+300, Alltrim(SA1->A1_MUN)                                    , oFont10, 070, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+370, Alltrim(SA1->A1_EST)                                    , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+385, Alltrim(AllToChar((cAliasSF2)->F2_VALFAT, cPictFat ))   , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+450, Alltrim(AllToChar((cAliasSF2)->F2_VALIPI, cPictIPI ))   , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+490, Alltrim(AllToChar((cAliasSF2)->F2_BSFCPST, cPictSub ))  , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+530,nLin+10,nCol+600)
            oPrint:Box(nLin,nCol+600,nLin+10,nCol+680)
            oPrint:Box(nLin,nCol+680,nLin+10,nCol+740)
            oPrint:Box(nLin,nCol+740,nLin+10,nCol+800)

            nValTot += (cAliasSF2)->F2_VALFAT
            nValIPI += (cAliasSF2)->F2_VALIPI
            nValSub += (cAliasSF2)->F2_BSFCPST

            nLin += 10
        ElseIF SA2->(MSSeek(FWxFilial("SA2")+(cAliasSF2)->F2_CLIENTE+(cAliasSF2)->F2_LOJA))  
            
            oPrint:SayAlign(nLin, nCol    , (cAliasSF2)->F2_DOC                                     , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+045, DToC(SToD((cAliasSF2)->F2_EMISSAO))                     , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+090, Transform(Alltrim(SA2->A2_CGC),IIF(SA2->A2_PESSOA == 'J',cPictCNPJ,cPictCPF)), oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+170, Alltrim(SA2->A2_NOME)                                   , oFont10, 130, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+300, Alltrim(SA2->A2_MUN)                                    , oFont10, 070, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+370, Alltrim(SA2->A2_EST)                                    , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+385, Alltrim(AllToChar((cAliasSF2)->F2_VALFAT, cPictFat ))   , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+450, Alltrim(AllToChar((cAliasSF2)->F2_VALIPI, cPictIPI ))   , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+490, Alltrim(AllToChar((cAliasSF2)->F2_BSFCPST, cPictSub ))  , oFont10, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+530,nLin+10,nCol+600)
            oPrint:Box(nLin,nCol+600,nLin+10,nCol+680)
            oPrint:Box(nLin,nCol+680,nLin+10,nCol+740)
            oPrint:Box(nLin,nCol+740,nLin+10,nCol+800)

            nValTot += (cAliasSF2)->F2_VALFAT
            nValIPI += (cAliasSF2)->F2_VALIPI
            nValSub += (cAliasSF2)->F2_BSFCPST

            nLin += 10
        EndIf

        (cAliasSF2)->(DBSkip())
    EndDo
    
    nLin += 15

    If nLin >= 600
        fCabec() //Imprime o Cabeçalho
    EndIf

    oPrint:SayAlign(nLin, nCol+310, "TOTAL EMPRESA: ", oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+385, Alltrim(AllToChar(nValTot, cPictFat )) , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+450, Alltrim(AllToChar(nValIPI, cPictIPI )) , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
    oPrint:SayAlign(nLin, nCol+490, Alltrim(AllToChar(nValSub, cPictSub )) , oFont10B, 280, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)

    If Select((cAliasDAK)) <> 0
		(cAliasSF2)->(DbCloseArea())
	Endif      

    oPrint:EndPage()

Return
