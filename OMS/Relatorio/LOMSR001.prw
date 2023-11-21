#INCLUDE "totvs.ch"
#INCLUDE "fwprintsetup.ch"
#INCLUDE "rptdef.ch" 
#INCLUDE "TBICONN.CH"
#INCLUDE "topconn.ch"

#DEFINE ENTER Chr(13)+Chr(10)
#define PAD_LEFT		0
#define PAD_RIGHT		1
#define PAD_CENTER   	2

/*/{Protheus.doc} LOMSR001
Relatório Comprovante de duplicatas Luzarte
@type function
@version
@author TOTVS Nordeste
@since 17/11/2023
@return 
/*/
User Function LOMSR001()
    Local aArea    := FWGetArea()
    
    Private cBarra   := IIF(GetRemoteType() == 1,"\","/")
    Private cNomArq  := ""
    Private cDirArq  := ""
    
    If Pergunte( "LOMSR001" , .T. , "Perguntas - Relatório Comprovante de Duplicatas" )
        cDirArq := TFileDialog("Arquivos Adobe PDF (*.pdf)",'Informe onde será gravado o arquivo.',,,.T.,/*GETF_MULTISELECT*/)
        cNomArq := SubSTR(cDirArq,RAT(cBarra,cDirArq)+1,Rat(".",SubSTR(cDirArq,RAT(cBarra,cDirArq)+1))-1)
        cDirArq := SubSTR(cDirArq,1,RAT(cBarra,cDirArq))
        
        If !Empty(cDirArq) .AND. !Empty(cNomArq)
            Processa({|| fLOMSR001()}, "Processando...")
        EndIf 
    EndIF 
    
    FWRestArea(aArea)

Return

Static Function fLOMSR001()
    
    Local aTitulos := {}
    Local nLin     := 0
    Local nCol     := 0
    Local nAtual   := 0
    Local nTotal   := 0
    Local cNomeEmp := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_NOMECOM" } )[1][2]))
    Local cCNPJ    := Transform(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_CGC" } )[1][2]),"@R 99.999.999/9999-99")
    Local cEndere  := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_ENDENT" } )[1][2]))
    Local cComple  := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_COMPENT" } )[1][2]))
    Local cBairro  := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_BAIRENT" } )[1][2]))
    Local cCidade  := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_CIDENT" } )[1][2]))
    Local cEstado  := Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_ESTENT" } )[1][2])
    Local cTelefo  := Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_TEL" } )[1][2])
    Local cEmail   := ""
    Local _cAlias  := GetNextAlias()
    Local cQry     := ""
    Local cNomVend := ""
    Local cPicture := PesqPict("SF2","F2_VALFAT")
    Local nY

    oFont12  := TFont():New( "Times New Roman",,12,,.F.,,,,,.F. )
    oFont12B := TFont():New( "Times New Roman",,12,,.T.,,,,,.F. )
    oFont14  := TFont():New( "Times New Roman",,14,,.F.,,,,,.F. )
    oFont14B := TFont():New( "Times New Roman",,14,,.T.,,,,,.F. )
    oFont16  := TFont():New( "Times New Roman",,16,,.F.,,,,,.F. )
    oFont16B := TFont():New( "Times New Roman",,16,,.T.,,,,,.F. )

    cQry := " SELECT * "
    cQry += " FROM " + RetSqlName("SF2") + " SF2 " 
    cQry += "   WHERE SF2.D_E_L_E_T_ <> '*'"
    cQry += "     AND SF2.F2_FILIAL  = '" + xFilial("SF2") + "'"
    cQry += "     AND SF2.F2_CARGA    BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "' "
    cQry += "     AND SF2.F2_CLIENTE  BETWEEN '" + MV_PAR03 + "' AND '" + MV_PAR04 + "' "
    cQry += "     AND SF2.F2_LOJA     BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "' "
    cQry += "     AND SF2.F2_DOC      BETWEEN '" + MV_PAR07 + "' AND '" + MV_PAR08 + "' "
    cQry += "     AND SF2.F2_SERIE    BETWEEN '" + MV_PAR09 + "' AND '" + MV_PAR10 + "' "
    cQry := ChangeQuery(cQry)
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),_cAlias,.F.,.T.)

    Count To nTotal
    ProcRegua(nTotal)
    
    DBSelectArea(_cAlias)
    (_cAlias)->(DbGoTop())

    If !(_cAlias)->(Eof())
        
        oPrint:=FWMSPrinter():New(cNomArq,IMP_PDF, .F., , .T.)
        oPrint:SetResolution(72)
        oPrint:SetPortrait()
        oPrint:SetPaperSize(DMPAPER_A4)
        oPrint:SetMargin(60,60,60,60) // nEsquerda, nSuperior, nDireita, nInferior
        oPrint:cPathPDF := cDirArq
        
        While !(_cAlias)->(Eof())
            
            nLin     := 50
            nCol     := 13
            aTitulos := fTitulos((_cAlias)->F2_DOC,(_cAlias)->F2_SERIE,(_cAlias)->F2_CLIENTE,(_cAlias)->F2_LOJA,(_cAlias)->F2_COND) //Busca os títulos da NF
            cNomVend := Upper(Alltrim(Posicione("SA3",1,FWxFilial("SA3")+(_cAlias)->F2_VEND1,"A3_NOME")))

            nAtual++
            IncProc("Processando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

            oPrint:StartPage()
            oPrint:SayAlign(nLin, nCol+97, cNomeEmp +" - "+ cCNPJ, oFont12, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
			nLin += 10
            oPrint:SayAlign(nLin, nCol+97, cEndere+", "+cComple+", "+cBairro+", "+cCidade+"/"+cEstado, oFont12, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
			nLin += 10
            oPrint:SayAlign(nLin, nCol+97, cTelefo+" - "+cEmail, oFont12, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
			nLin += 15
            oPrint:SayAlign(nLin, nCol+97, "Comprovante de Duplicatas", oFont16B, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
			nLin += 15
            oPrint:SayAlign(nLin, nCol+97, "Nº Nota fiscal: "+(_cAlias)->F2_DOC+" - "+DToC(SToD((_cAlias)->F2_EMISSAO)), oFont16B, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
			nLin += 15
            oPrint:SayAlign(nLin, nCol+97, "DADOS DO REMETENTE", oFont14B, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
			nLin += 20
            oPrint:Box(nLin,nCol,nLin+20,nCol+350)
            oPrint:SayAlign(nLin+5, nCol+5, "Responsável: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+60, cNomeEmp, oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+350,nLin+20,nCol+530)
            oPrint:SayAlign(nLin+5, nCol+355, "Fone: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+380, cTelefo, oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:SayAlign(nLin, nCol+97, "DADOS DO CLIENTE", oFont14B, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
            nLin += 20
            oPrint:Box(nLin,nCol,nLin+20,nCol+350)
            oPrint:SayAlign(nLin+5, nCol+5, "Cliente: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+40, Upper(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_NOME"))), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+350,nLin+20,nCol+530)
            oPrint:SayAlign(nLin+5, nCol+355, "Fone: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+380, "("+Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_DDD"))+")", oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+398, Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_TEL")), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 25
            oPrint:Box(nLin,nCol,nLin+20,nCol+170)
            oPrint:SayAlign(nLin+5, nCol+5, "CNPJ/CPF: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+55, Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_CGC"), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+170,nLin+20,nCol+320)
            oPrint:SayAlign(nLin+5, nCol+175, "I.E: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+193, Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_INSCR")), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+320,nLin+20,nCol+530)
            oPrint:SayAlign(nLin+5, nCol+325, "E-mail: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+360, LOWER(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_EMAIL"))), oFont12, 300, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:SayAlign(nLin, nCol+10, "ENDEREÇO DO CLIENTE", oFont14B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:Box(nLin,nCol,nLin+20,nCol+530)
            oPrint:SayAlign(nLin+5, nCol+5, "Endereço: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+50, Upper(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_END"))), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:Box(nLin,nCol,nLin+20,nCol+530)
            nLin += 20
            oPrint:Box(nLin,nCol,nLin+20,nCol+100)
            oPrint:SayAlign(nLin+5, nCol+5, "CEP: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+30, Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_CEP"), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+100,nLin+20,nCol+250)
            oPrint:SayAlign(nLin+5, nCol+105, "Bairro: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+138, Upper(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_BAIRRO"))), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+250,nLin+20,nCol+450)
            oPrint:SayAlign(nLin+5, nCol+255, "Cidade: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+290, Upper(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_MUN"))), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+450,nLin+20,nCol+530)
            oPrint:SayAlign(nLin+5, nCol+455, "Estado: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+488, Upper(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_EST")), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:SayAlign(nLin, nCol+10, "DADOS PARA ENTREGA", oFont14B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:Box(nLin,nCol,nLin+20,nCol+530)
            oPrint:SayAlign(nLin+5, nCol+5, "Endereço: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+50, Upper(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_ENDENT"))), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:Box(nLin,nCol,nLin+20,nCol+100)
            oPrint:SayAlign(nLin+5, nCol+5, "CEP: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+30, Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_CEPE"), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+100,nLin+20,nCol+250)
            oPrint:SayAlign(nLin+5, nCol+105, "Bairro: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+138, Upper(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_BAIRROE"))), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+250,nLin+20,nCol+450)
            oPrint:SayAlign(nLin+5, nCol+255, "Cidade: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+290, Upper(Alltrim(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_MUNE"))), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:Box(nLin,nCol+450,nLin+20,nCol+530)
            oPrint:SayAlign(nLin+5, nCol+455, "Estado: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+5, nCol+488, Upper(Posicione("SA1",1,FWxFilial("SA1")+(_cAlias)->F2_CLIENTE+(_cAlias)->F2_LOJA,"A1_ESTE")), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 50
            oPrint:SayAlign(nLin, nCol+270, "Valor Total da Compra: R$ ", oFont14B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+405, Alltrim(AllToChar((_cAlias)->F2_VALFAT, cPicture )), oFont14B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            oPrint:Box(nLin,nCol,nLin+100,nCol+200)
            oPrint:Box(nLin+100,nCol,nLin+200,nCol+400)
            oPrint:Box(nLin,nCol+200,nLin+200,nCol+530)
            oPrint:SayAlign(nLin+10, nCol+35, "CONDIÇÕES DE PAGAMENTO", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+30, nCol+60, (_cAlias)->F2_COND + " - " + Alltrim(Posicione("SE4",1,FWxFilial("SE4")+(_cAlias)->F2_COND,"E4_DESCRI")), oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+10, nCol+205, "PARCELA", oFont12B, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+10, nCol+270, "VENCTO", oFont12B, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+10, nCol+340, "VALOR", oFont12B, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+10, nCol+420, "MODO DE PAGAMENTO", oFont12B, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 20
            For nY := 1 To Len(aTitulos)
                nLin += 10
                oPrint:SayAlign(nLin, nCol+205, aTitulos[nY,1], oFont12, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
                oPrint:SayAlign(nLin, nCol+270, aTitulos[nY,2], oFont12, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
                oPrint:SayAlign(nLin, nCol+340, "R$ "+aTitulos[nY,3], oFont12, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
                oPrint:SayAlign(nLin, nCol+420, aTitulos[nY,4], oFont12, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            Next
            nLin := 410
            nLin += 100
            oPrint:SayAlign(nLin+10, nCol+5, "Representante: ", oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin+30, nCol+20, cNomVend, oFont12B, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 110
            oPrint:Box(nLin,nCol,nLin+200,nCol+530)
            nLin += 20
            oPrint:SayAlign(nLin, nCol+97, "RECEBIMENTO", oFont14B, 280, /*nHeigth*/, CLR_BLACK , PAD_CENTER, PAD_CENTER)
            nLin += 30
            oPrint:SayAlign(nLin, nCol+5, "CLIENTE RECEBEDOR: ", oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+110, REPLICATE("_", 83), oFont12, 900, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 30
            oPrint:SayAlign(nLin, nCol+5, "DATA: ", oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+40, "__________/__________/____________________", oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            nLin += 30
            oPrint:SayAlign(nLin, nCol+5, "OBS: ", oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)
            oPrint:SayAlign(nLin, nCol+35, REPLICATE("_", 100), oFont12, 500, /*nHeigth*/, CLR_BLACK , PAD_LEFT, PAD_LEFT)

            oPrint:EndPage()

            (_cAlias)->(DBSkip())
        EndDo

        (_cAlias)->(DbCloseArea())
        oPrint:Preview()
        ms_flush() 
    Else
        FWAlertWarning("Nenhuma informação encontrada com os parâmetros informados.",'Parâmetros da pergunta "LOMSR001"')
    EndIF

Return

Static Function fTitulos(pDoc,pSerie,pCliente,pLoja,pCond)
    Local aTitulos := {}
    Local _cAlias  := GetNextAlias()
    Local cQry     := ""
    Local cPicture := PesqPict("SE1","E1_SALDO")
    Local cCondPg  := Alltrim(FWGetSX5('24',Alltrim(Posicione("SE4",1,FwxFilial("SE4")+pCond,"E4_FORMA")),'pt-br')[1][4])

    cQry := " SELECT * "
    cQry += " FROM " + RetSqlName("SE1") + " SE1 " 
    cQry += "   WHERE SE1.D_E_L_E_T_ <> '*'"
    cQry += "     AND SE1.E1_FILIAL  = '" + xFilial("SE1") + "'"
    cQry += "     AND SE1.E1_PREFIXO = '" + pSerie + "' "
    cQry += "     AND SE1.E1_NUM     = '" + pDoc + "' "
    cQry += "     AND SE1.E1_CLIENTE = '" + pCliente + "' "
    cQry += "     AND SE1.E1_LOJA    = '" + pLoja + "' "
    cQry := ChangeQuery(cQry)
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),_cAlias,.F.,.T.)

    While !(_cAlias)->(Eof())

        aAdd(aTitulos,{Alltrim((_cAlias)->E1_NUM)+IIF(!Empty((_cAlias)->E1_PARCELA),"-"+Alltrim((_cAlias)->E1_PARCELA),""),;
                       DToC(SToD((_cAlias)->E1_VENCREA)),;
                       Alltrim(AllToChar((_cAlias)->E1_SALDO, cPicture )),;
                       cCondPg})

        (_cAlias)->(DBSkip())
    EndDo
    
    (_cAlias)->(DbCloseArea())

Return (aTitulos)
