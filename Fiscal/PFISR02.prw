#Include "Protheus.ch"
#Include "FwPrintSetup.Ch" 
#Include "RptDef.Ch" 

#DEFINE ENTER Chr(13)+Chr(10)
#define PAD_LEFT		0
#define PAD_RIGHT		1
#define PAD_CENTER   	2

Static cIE := Transform(Alltrim(FWSM0Util():GetSM0Data(cEmpAnt,cFilAnt,{"M0_INSC"})[1,2]), "@R 9999999-99")
Static cCNPJ := TransForm(Alltrim(FWSM0Util():GetSM0Data(cEmpAnt,cFilAnt,{"M0_CGC"})[1,2]), "@R 99.999.999/9999-99")
Static cNomEmp := IIF(!Empty(Alltrim(FWSM0Util():GetSM0Data(cEmpAnt,cFilAnt,{"M0_FULNAME"})[1,2])),;
                      Alltrim(Capital(FWSM0Util():GetSM0Data(cEmpAnt,cFilAnt,{"M0_FULNAME"})[1,2])),;
                      Alltrim(Capital(FWSM0Util():GetSM0Data(cEmpAnt,cFilAnt,{"M0_NOMECOM"})[1,2])))

/*/{Protheus.doc} PFISR03
Impressão Apuração Ressarcimento ICMS
@type function
@version
@author TOTVS Nordeste
@since 28/08/2023
@return 
/*/
User Function PFISR02()

MsAguarde({||GeraExcel()},"Aguarde","Gerando dados para a Planilha",.F.)

Return

/*/{Protheus.doc} GeraExcel
Função para impressão do Resumo de operações por CFOP
@author TOTVS NORDESTE
@since 11/08/2023
@version 1.0
    @return Nil, Função não tem retorno
    @example
    u_GeraExcel()
    @obs 
/*/

Static Function GeraExcel()

Local oExcel
Local oLeft,oRight,oCenter,oQuebraTxt,oAzulLCalN,oAzulLAri,oAzulRAri,oAzulCCalN,oBranLCal,oBranLCalN
Local oAzulCCalNQ,oBranCVerQ,oBranCVerN,oBranLAri,oBranRAri,oBranCCal,oBranCCalN,oFmtNumBr,oFmtNumBrN,oFmtNumAz,oFmtNumAzN
Local nCorAzul,nCorBranc,nBordaAll,nAri10,nAri10N,nCali11,nCali11N,nVerdana9,nVerdana9N,nFmtNum,nLin
Local cNomeFor  := SubSTR(Alltrim(SZH->ZH_NOME),1,At(" ",Alltrim(SZH->ZH_NOME)))
Local cPlan     := cNomeFor+" "+Alltrim(SubSTR(SZH->ZH_PERIODO,1,2))+"-"+Alltrim(SubSTR(SZH->ZH_PERIODO,3))
Local cMesAno   := Upper(SubSTR(MesExtenso(Val(SubSTR(SZH->ZH_PERIODO,1,2))),1,3))+" "+Alltrim(SubSTR(SZH->ZH_PERIODO,3))
Local nTotQtdNF := 0
Local nTotQtdSC := 0
Local nValTotal := 0
Local nTotCredF := 0
Local nTotValRe := 0

	oExcel	:= YExcel():new("PFISR02"+FWTimeStamp(1))
    nCorAzul    := oExcel:CorPreenc("B8CCE4")	//Cor de Fundo Azul
    nCorBranc   := oExcel:CorPreenc("FFFFFF")	//Cor de Fundo Branca
	oLeft	    := oExcel:Alinhamento("left","center")
    oRight	    := oExcel:Alinhamento("right","center")
    oCenter	    := oExcel:Alinhamento("center","center")
	oQuebraTxt	:= oExcel:Alinhamento("center","center",,.T.)
	nBordaAll	:= oExcel:Borda("ALL")
    nAri10		:= oExcel:AddFont(10,"000000","Arial","2",)
    nAri10N		:= oExcel:AddFont(10,"000000","Arial","2",,.T.)
    nCali11		:= oExcel:AddFont(11,"000000","Calibri","2",)
    nCali11N	:= oExcel:AddFont(11,"000000","Calibri","2",,.T.)
    nVerdana9   := oExcel:AddFont(9,"000000","Verdana","2",)
    nVerdana9N  := oExcel:AddFont(9,"000000","Verdana","2",)
	nFmtNum     := oExcel:AddFmtNum(2/*nDecimal*/,.T./*lMilhar*/,/*cPrefixo*/,/*cSufixo*/,/*cNegINI*/,/*cNegFim*/,/*cValorZero*/,/*cCor*/,/*cCorNeg*/,/*nNumFmtId*/)

	oAzulLCalN	:= oExcel:NewStyle():Setfont(nCali11N):Setfill(nCorAzul):SetaValores({oLeft}):Setborder(nBordaAll)
    oAzulLAri	:= oExcel:NewStyle():Setfont(nAri10):Setfill(nCorAzul):SetaValores({oLeft}):Setborder(nBordaAll)
    oAzulRAri	:= oExcel:NewStyle():Setfont(nAri10):Setfill(nCorAzul):SetaValores({oRight}):Setborder(nBordaAll)
    oAzulCCalN	:= oExcel:NewStyle():Setfont(nCali11N):Setfill(nCorAzul):SetaValores({oCenter}):Setborder(nBordaAll)
    oAzulCCalNQ	:= oExcel:NewStyle():Setfont(nCali11N):Setfill(nCorAzul):SetaValores({oQuebraTxt}):Setborder(nBordaAll)
    oBranCVerQ	:= oExcel:NewStyle():Setfont(nVerdana9):Setfill(nCorBranc):SetaValores({oQuebraTxt}):Setborder(nBordaAll)
	oBranCVerN	:= oExcel:NewStyle():Setfont(nVerdana9N):Setfill(nCorBranc):SetaValores({oQuebraTxt}):Setborder(nBordaAll)
    oBranLAri	:= oExcel:NewStyle():Setfont(nAri10):Setfill(nCorBranc):SetaValores({oLeft}):Setborder(nBordaAll)
    oBranLCal	:= oExcel:NewStyle():Setfont(nCali11):Setfill(nCorBranc):SetaValores({oLeft}):Setborder(nBordaAll)
    oBranLCalN	:= oExcel:NewStyle():Setfont(nCali11N):Setfill(nCorBranc):SetaValores({oLeft}):Setborder(nBordaAll)
    oBranRAri	:= oExcel:NewStyle():Setfill(nCorBranc):SetaValores({oRight}):Setborder(nBordaAll)
    oBranCCal	:= oExcel:NewStyle():Setfont(nCali11):Setfill(nCorBranc):SetaValores({oLeft}):Setborder(nBordaAll)
    oBranCCalN	:= oExcel:NewStyle():Setfont(nCali11N):Setfill(nCorBranc):SetaValores({oCenter}):Setborder(nBordaAll)
    oFmtNumBr   := oExcel:NewStyle():SetFont(nAri10):SetnumFmt(nFmtNum):Setfill(nCorBranc):Setborder(nBordaAll)
    oFmtNumBrN  := oExcel:NewStyle():SetFont(nAri10N):SetnumFmt(nFmtNum):Setfill(nCorBranc):Setborder(nBordaAll)
    oFmtNumAz   := oExcel:NewStyle():SetFont(nAri10):SetnumFmt(nFmtNum):Setfill(nCorAzul):Setborder(nBordaAll)
    oFmtNumAzN  := oExcel:NewStyle():SetFont(nAri10N):SetnumFmt(nFmtNum):Setfill(nCorAzul):Setborder(nBordaAll)

	oExcel:ADDPlan(cPlan)
    
    //Colunas
    oExcel:AddTamCol(1,1,13.50)
    oExcel:AddTamCol(2,2,18.00)
	oExcel:AddTamCol(3,3,3.00)
    oExcel:AddTamCol(4,4,6.00)
    oExcel:AddTamCol(5,5,13.00)
    oExcel:AddTamCol(6,6,11.00)
    oExcel:AddTamCol(7,7,23.00)
    oExcel:AddTamCol(8,8,8.00)
	oExcel:AddTamCol(9,9,19.00)
    oExcel:AddTamCol(10,10,27.00)
    oExcel:AddTamCol(11,11,46.00)

    oExcel:mergeCells(1,1,1,11) //Mescla as células A1:K1
    oExcel:Pos(1,1):SetValue("SOLICITANTE: "+Upper(Alltrim(cNomEmp))):SetStyle(oAzulCCalN)
    oExcel:Pos(1,2):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,3):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,4):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,5):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,6):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,7):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,8):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,9):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,10):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(1,11):SetValue(""):SetStyle(oAzulCCalN)
    
    oExcel:mergeCells(2,1,2,11) //Mescla as células A2:K2
    oExcel:Pos(2,1):SetValue('APURAÇÃO RESSARCIMENTO ICMS AQUISIÇÃO FARINHA DE TRIGO ( ART. 8º, II, "a" DO DECRETO Nº 27.987/05)'):SetStyle(oAzulCCalN)
    oExcel:Pos(2,2):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,3):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,4):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,5):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,6):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,7):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,8):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,9):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,10):SetValue(""):SetStyle(oAzulCCalN)
    oExcel:Pos(2,11):SetValue(""):SetStyle(oAzulCCalN)
    
    oExcel:mergeCells(3,1,3,2) //Mescla as células A3:B3
    oExcel:mergeCells(3,6,3,7) //Mescla as células F3:G3
    oExcel:mergeCells(3,8,3,9) //Mescla as células H3:I3
    oExcel:Pos(3,1):SetValue('CÓD. FORC.'):SetStyle(oAzulCCalN)
    oExcel:Pos(3,2):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(3,3):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(3,4):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(3,5):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(3,6):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(3,7):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(3,8):SetValue('COMPETÊNCIA'):SetStyle(oAzulLCalN)
    oExcel:Pos(3,9):SetValue(''):SetStyle(oAzulRAri)
    oExcel:Pos(3,10):SetValue(Alltrim(cMesAno)):SetStyle(oAzulRAri)
    oExcel:Pos(3,11):SetValue(''):SetStyle(oBranRAri)
    oExcel:mergeCells(3,11,8,11) //Mescla as células K3:L8

    oExcel:mergeCells(4,1,4,2) //Mescla as células A4:B4
    oExcel:mergeCells(4,6,4,7) //Mescla as células F4:G4
    oExcel:mergeCells(4,8,4,9) //Mescla as células H4:I4
    oExcel:Pos(4,1):SetValue('FORNECEDOR'):SetStyle(oAzulCCalN)
    oExcel:Pos(4,2):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(4,3):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(4,4):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(4,5):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(4,6):SetValue(Alltrim(SZH->ZH_NOME)):SetStyle(oAzulLAri)
    oExcel:Pos(4,7):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(4,8):SetValue('Vlr Ressarc.'):SetStyle(oAzulLCalN)
    oExcel:Pos(4,9):SetValue(""):SetStyle(oAzulRAri)
    oExcel:Pos(4,10):SetValue(SZH->ZH_VALRESS):SetStyle(oAzulRAri)
    oExcel:Pos(4,11):SetValue(""):SetStyle(oBranRAri)

    oExcel:mergeCells(5,1,5,2) //Mescla as células A5:B5
    oExcel:mergeCells(5,6,5,7) //Mescla as células F5:G5
    oExcel:mergeCells(5,8,5,9) //Mescla as células H5:I5
    oExcel:Pos(5,1):SetValue('CNPJ'):SetStyle(oAzulCCalN)
    oExcel:Pos(5,2):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(5,3):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(5,4):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(5,5):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(5,6):SetValue(Transform(Alltrim(SZH->ZH_CGC), "@R 99.999.999/9999-99" )):SetStyle(oAzulLAri)
    oExcel:Pos(5,7):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(5,8):SetValue('Pauta'):SetStyle(oAzulLCalN)
    oExcel:Pos(5,9):SetValue(''):SetStyle(oAzulLCalN)
    oExcel:Pos(5,10):SetValue(SZH->ZH_PAUTA):SetStyle(oAzulRAri)
    oExcel:Pos(5,11):SetValue(""):SetStyle(oBranRAri)

    oExcel:mergeCells(6,1,6,2) //Mescla as células A6:B6
    oExcel:mergeCells(6,6,6,7) //Mescla as células F6:G6
    oExcel:mergeCells(6,8,6,9) //Mescla as células H6:I6
    oExcel:Pos(6,1):SetValue('UF'):SetStyle(oAzulCCalN)
    oExcel:Pos(6,2):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(6,3):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(6,4):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(6,5):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(6,6):SetValue(SZH->ZH_EST):SetStyle(oAzulLAri)
    oExcel:Pos(6,7):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(6,8):SetValue('Crédito Ind. Alim.'):SetStyle(oAzulLCalN)
    oExcel:Pos(6,9):SetValue(''):SetStyle(oAzulLCalN)
    oExcel:Pos(6,10):SetValue(SZH->ZH_CREDIND):SetStyle(oAzulRAri)
    oExcel:Pos(6,11):SetValue(""):SetStyle(oBranRAri)

    oExcel:mergeCells(7,1,8,7) //Mescla as células A7:G8
    oExcel:mergeCells(7,8,7,9) //Mescla as células H7:I7
    oExcel:Pos(7,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(7,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(7,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(7,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(7,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(7,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(7,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(7,8):SetValue('% Prod. Benef. Ressarcimento'):SetStyle(oAzulCCalN)
    oExcel:Pos(7,9):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(7,10):SetValue(SZH->ZH_PRODBEN):SetStyle(oAzulRAri)
    oExcel:Pos(7,11):SetValue(""):SetStyle(oBranRAri)

    oExcel:mergeCells(8,8,8,9) //Mescla as células H8:I8
    oExcel:Pos(8,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(8,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(8,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(8,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(8,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(8,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(8,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(8,8):SetValue('Crédito Ressarcimento'):SetStyle(oAzulCCalN)
    oExcel:Pos(8,9):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(8,10):SetValue(SZH->ZH_CREDRES):SetStyle(oAzulRAri)
    oExcel:Pos(8,11):SetValue(""):SetStyle(oBranRAri)

    oExcel:mergeCells(9,1,9,11) //Mescla as células A1:K1
    oExcel:Pos(9,1):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,2):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,3):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,4):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,5):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,6):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,7):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,8):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,9):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,10):SetValue(""):SetStyle(oBranRAri)
    oExcel:Pos(9,11):SetValue(""):SetStyle(oBranRAri)

    oExcel:mergeCells(10,1,10,7) //Mescla as células A10:G10
    oExcel:Pos(10,1):SetValue('FARINHA DE TRIGO ADQUIRIDA'):SetStyle(oAzulCCalN)
    oExcel:Pos(10,2):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(10,3):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(10,4):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(10,5):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(10,6):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(10,7):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(10,8):SetValue('ICMS'):SetStyle(oAzulCCalN)
    oExcel:Pos(10,9):SetValue('CRÉDITO'):SetStyle(oAzulCCalN)
    oExcel:Pos(10,10):SetValue('VALOR DO'):SetStyle(oAzulCCalN)
    oExcel:Pos(10,11):SetValue(""):SetStyle(oAzulCCalN)

    oExcel:Pos(11,1):SetValue('Nº NF'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,2):SetValue('DATA DA ENTRADA'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,3):SetValue('KG'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,4):SetValue('KG NF'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,5):SetValue('QTD NF'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,6):SetValue('QTD. SACOS 50KG(A)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,7):SetValue('VALOR TOAL'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,8):SetValue('P/ SACO 50KG (B)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,9):SetValue('FISCAL C = (A)X(B)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,10):SetValue('RESSARCIMENTO (D) = (C) X (%Prod. Benef. Ressarcimento)X(60%)X(90%)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(11,11):SetValue('CHAVE NOTA FISCAL'):SetStyle(oAzulCCalNQ)

    oExcel:mergeCells(12,2,12,11) //Mescla as células A10:G10
    oExcel:Pos(12,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(12,11):SetValue(''):SetStyle(oBranRAri)

    nLin := 13

    DBSelectArea("SZI")
    SZI->(dbSetOrder(1))
    If SZI->(MsSeek(FWxFilial("SZI")+SZH->ZH_COD))
        
        While !SZI->(Eof()) .And. SZI->ZI_COD == SZH->ZH_COD .And. SZI->ZI_PERIODO == SZH->ZH_PERIODO
            
            oExcel:Pos(nLin,1):SetValue(SZI->ZI_DOC+"-"+Alltrim(SZI->ZI_SERIE)):SetStyle(oBranRAri)
            oExcel:Pos(nLin,2):SetValue(DToC(SZI->ZI_DTDIGIT)):SetStyle(oBranRAri)
            oExcel:Pos(nLin,3):SetValue(SZI->ZI_QTDKG):SetStyle(oBranRAri)
            oExcel:Pos(nLin,4):SetValue(SZI->ZI_QTDKGNF):SetStyle(oBranRAri)
            oExcel:Pos(nLin,5):SetValue(SZI->ZI_QTDNF):SetStyle(oBranRAri)
            oExcel:Pos(nLin,6):SetValue(SZI->ZI_QTDSACO):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,7):SetValue(SZI->ZI_VALOR):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,8):SetValue(SZI->ZI_ICMSACO):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,9):SetValue(SZI->ZI_CREDFIS):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,10):SetValue(SZI->ZI_VALRESS):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,11):SetValue(SZI->ZI_CHAVE):SetStyle(oBranRAri)
            
            nTotQtdNF += SZI->ZI_QTDNF
            nTotQtdSC += SZI->ZI_QTDSACO
            nValTotal += SZI->ZI_VALOR
            nTotCredF += SZI->ZI_CREDFIS
            nTotValRe += SZI->ZI_VALRESS

            nLin++

            SZI->(DBSkip())
        EndDo

    EndIf 

    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranRAri)

    nLin++

    oExcel:mergeCells(nLin,1,nLin,2) //Mescla as células
    oExcel:Pos(nLin,1):SetValue('TOTAL'):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,5):SetValue(nTotQtdNF):SetStyle(oFmtNumAzN)
    oExcel:Pos(nLin,6):SetValue(nTotQtdSC):SetStyle(oFmtNumAzN)
    oExcel:Pos(nLin,7):SetValue(nValTotal):SetStyle(oFmtNumAzN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,9):SetValue(nTotCredF):SetStyle(oFmtNumAzN)
    oExcel:Pos(nLin,10):SetValue(nTotValRe):SetStyle(oFmtNumAzN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oAzulCCalN)

    nLin++

    oExcel:mergeCells(nLin,1,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranRAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,2):SetValue('PERDA, VENDA, DEVOLUÇÃO'):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCCalN)

    nLin++

    oExcel:Pos(nLin,1):SetValue('Nº NF'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,2):SetValue('DATA DA ENTRADA'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,3):SetValue('KG'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,4):SetValue('KG NF'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,5):SetValue('QTD NF'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,6):SetValue('QTD. SACOS 50KG(A)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,7):SetValue('VALOR TOAL'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,8):SetValue('P/ SACO 50KG (B)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,9):SetValue('FISCAL C = (A)X(B)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,10):SetValue('RESSARCIMENTO (D) = (C) X (%Prod. Benef. Ressarcimento)X(60%)X(90%)'):SetStyle(oAzulCCalNQ)
    oExcel:Pos(nLin,11):SetValue('CHAVE NOTA FISCAL'):SetStyle(oAzulCCalNQ)

    nLin++

    nTotCredF := 0
    nTotValRe := 0

    DBSelectArea("SZJ")
    SZJ->(dbSetOrder(1))
    If SZJ->(MsSeek(FWxFilial("SZJ")+SZH->ZH_COD))
        
        While !SZJ->(Eof()) .And. SZJ->ZJ_COD == SZH->ZH_COD .And. SZJ->ZJ_PERIODO == SZH->ZH_PERIODO
            
            oExcel:Pos(nLin,1):SetValue(SZJ->ZJ_DOC+"-"+Alltrim(SZJ->ZJ_SERIE)):SetStyle(oBranRAri)
            oExcel:Pos(nLin,2):SetValue(DToC(SZJ->ZJ_DTDIGIT)):SetStyle(oBranRAri)
            oExcel:Pos(nLin,3):SetValue(SZJ->ZJ_QTDKG):SetStyle(oBranRAri)
            oExcel:Pos(nLin,4):SetValue(SZJ->ZJ_QTDKGNF):SetStyle(oBranRAri)
            oExcel:Pos(nLin,5):SetValue(SZJ->ZJ_QTDNF):SetStyle(oBranRAri)
            oExcel:Pos(nLin,6):SetValue(SZJ->ZJ_QTDSACO):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,7):SetValue(SZJ->ZJ_VALOR):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,8):SetValue(SZJ->ZJ_ICMSACO):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,9):SetValue(SZJ->ZJ_CREDFIS):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,10):SetValue(SZJ->ZJ_VALRESS):SetStyle(oFmtNumBr)
            oExcel:Pos(nLin,11):SetValue(SZJ->ZJ_CHAVE):SetStyle(oBranRAri)

            nTotCredF += SZJ->ZJ_CREDFIS
            nTotValRe += SZJ->ZJ_VALRESS

            nLin++

            SZI->(DBSkip())
        EndDo

    EndIf

    nLin++

    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,2):SetValue('Consumo'):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLCalN)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,6) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,2):SetValue('Venda de Farinha reprovada'):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCCalN)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,5) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,2):SetValue('TOTAL RESSARCIMENTO'):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oAzulCCalN)
    oExcel:Pos(nLin,9):SetValue(nTotCredF):SetStyle(oFmtNumAzN)
    oExcel:Pos(nLin,10):SetValue(nTotValRe):SetStyle(oFmtNumAzN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oAzulCCalN)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranRAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,2):SetValue('Informações do Contribuinte: '):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranRAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranRAri)

    nLin++

    oExcel:mergeCells(nLin,3,nLin,5) //Mescla as células
    oExcel:mergeCells(nLin,6,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,2):SetValue('RAZÃO SOCIAL: '):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,6):SetValue(Upper(cNomEmp)):SetStyle(oBranCCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCCalN)

    nLin++

    oExcel:mergeCells(nLin,3,nLin,5) //Mescla as células
    oExcel:mergeCells(nLin,6,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,2):SetValue('INSC. ESTADUAL: '):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,6):SetValue(cIE):SetStyle(oBranCCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCCalN)

    nLin++

    oExcel:mergeCells(nLin,3,nLin,5) //Mescla as células
    oExcel:mergeCells(nLin,6,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,2):SetValue('CNPJ: '):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,6):SetValue(cCNPJ):SetStyle(oBranCCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCCalN)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranRAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranRAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranRAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,2):SetValue('Legislação: '):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,2):SetValue('Decreto PE nº 27.987/2005 e Alterações'):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCCalN)

    nLin++

    oExcel:SetRowH(33.00,nLin)	//Defini o tamanho da linha
    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,2):SetValue('Art. 8º Relativamente às operações promovidas por estabelecimento beneficiário do PRODEPE, '+;
                                'industrial de produtos alimentícios derivados de farinha de trigo ou de suas misturas, '+;
                                'conforme indicados no art. 1º, II, deverá ser observado o seguinte:'):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCCalN)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCCalN)

    nLin++

    oExcel:SetRowH(15.00,nLin)	//Defini o tamanho da linha
    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,2):SetValue('(...)'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,2):SetValue('II – o contribuinte poderá efetuar o ressarcimento do valor'+;
                                ' relativo ao referido benefício, sem prejuízo das demais normas'+;
                                ' previstas neste Decreto, adotando os seguintes procedimentos:'):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranCVerQ)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranCVerQ)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,2):SetValue('(...)'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLAri)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLAri)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,11) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,10,nLin+3,10) //Mescla as células
    oExcel:mergeCells(nLin+4,10,nLin+8,10) //Mescla as células

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('Decreto PE Nº:'):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue("Telefone Contato:"):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,11):SetValue('(81) 3634-1777'):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('27.793/2005'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue('(81) 99792-6222'):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('33.663/2009'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue('(81) 99736-0958'):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('41.042/2014'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue('(81) 99736-0958'):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('43.200/2016'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue("Email Contato: "):SetStyle(oBranLCalN)
    oExcel:Pos(nLin,11):SetValue('coordenador.fiscal@pancristal.com.br'):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('46.081/2018'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue('controladoria@pancristal.com.br'):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('47.223/2019'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('48.601/2020'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLCal)

    nLin++

    oExcel:mergeCells(nLin,2,nLin,9) //Mescla as células
    oExcel:Pos(nLin,1):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,2):SetValue('48.602/2020'):SetStyle(oBranLAri)
    oExcel:Pos(nLin,3):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,4):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,5):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,6):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,7):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,8):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,9):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,10):SetValue(''):SetStyle(oBranLCal)
    oExcel:Pos(nLin,11):SetValue(''):SetStyle(oBranLCal)

    oExcel:Save(GetTempPath())
	oExcel:OpenApp()
	oExcel:Close()

return
