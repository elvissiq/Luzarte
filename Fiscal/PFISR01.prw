#Include "Protheus.ch"
#Include "FwPrintSetup.Ch" 
#Include "RptDef.Ch"
#include 'parmtype.ch'

#DEFINE ENTER Chr(13)+Chr(10)
#define PAD_LEFT		0
#define PAD_RIGHT		1
#define PAD_CENTER   	2

/*/{Protheus.doc} PFISR03
Função para chamada de impressão do relatório GeraExcel
@type function
@version
@author TOTVS Nordeste
@since 24/11/2023
@return 
/*/
User Function PFISR01()

MsAguarde({||GeraExcel()},"Aguarde","Gerando dados para a Planilha",.F.)

Return

/*/{Protheus.doc} GeraExcel
Função para impressão do Resumo de operações por CFOP
@author TOTVS NORDESTE
@since 24/11/2023
@version 1.0
    @return Nil, Função não tem retorno
    @example
    u_GeraExcel()
    @obs 
/*/

Static Function GeraExcel()

Local oExcel
Local oLeft,oCenter,oQuebraTxt,oEstiloB,oEstiloBN,oEstiloLef
Local oFundoCinza,oFundoBranc,oEstiloCent,oEstiloCenN,oFmtNum,oFmtNumN
Local nCorCinza,nCorBranc,nAri10,nAri10N,nCali11N,nFmtNum
Local nBordaAll,nIdImg,nLin
Local cPerFisc := "Periodo Fiscal: "+Upper(MesExtenso(Val(SubSTR(SZG->ZG_PERIODO,1,2))))+Alltrim(SubSTR(SZG->ZG_PERIODO,3))
Local cNomeEmp := UPPER(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_NOMECOM" } )[1][2]))
Local cCNPJ    := Transform(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_CGC" } )[1][2]),"@R 99.999.999/9999-99")
Local cEstado  := UPPER(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_ESTENT" } )[1][2]))
Local cQry     := ""
Local nTotVC1  := 0
Local nTotBC1  := 0
Local nTotICM1 := 0
Local nTotVC2  := 0
Local nTotBC2  := 0
Local nTotICM2 := 0
Local nTotVC3  := 0
Local nTotBC3  := 0
Local nTotICM3 := 0
Local nTotVC4  := 0
Local nTotBC4  := 0
Local nTotICM4 := 0
Local nTotCalc := 0
Local __cAlias := "TEMP"+FWTimeStamp(1)

	oExcel	:= YExcel():new("PFISR01"+FWTimeStamp(1))
    nCorCinza   := oExcel:CorPreenc("CCCCCC")	//Cor de Fundo Cinza
    nCorBranc   := oExcel:CorPreenc("FFFFFF")	//Cor de Fundo Branca
	oLeft	    := oExcel:Alinhamento("left","center")
    oCenter	    := oExcel:Alinhamento("center","center")
	oQuebraTxt	:= oExcel:Alinhamento("center","center",,.T.)
	nBordaAll	:= oExcel:Borda("ALL")
    nAri10		:= oExcel:AddFont(10,"000000","Arial","2",)
    nAri10N		:= oExcel:AddFont(10,"000000","Arial","2",,.T.)
    nCali11N	:= oExcel:AddFont(11,"000000","Calibri","2",,.T.)
	nFmtNum     := oExcel:AddFmtNum(2/*nDecimal*/,.T./*lMilhar*/,/*cPrefixo*/,/*cSufixo*/,/*cNegINI*/,/*cNegFim*/,/*cValorZero*/,/*cCor*/,/*cCorNeg*/,/*nNumFmtId*/)

	oEstiloB	:= oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10):SetaValores({oCenter}):Setborder(nBordaAll)
    oEstiloBN	:= oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10N):SetaValores({oCenter}):Setborder(nBordaAll)
    oEstiloLef  := oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10):SetaValores({oLeft}):Setborder(nBordaAll)
    oEstiloLefN := oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10N):SetaValores({oLeft}):Setborder(nBordaAll)
    oEstiloCent := oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10):SetaValores({oCenter}):Setborder(nBordaAll)
    oEstiloCenN := oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10N):SetaValores({oCenter}):Setborder(nBordaAll)
    oFmtNum     := oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10):SetnumFmt(nFmtNum):Setborder(nBordaAll)
    oFmtNumN    := oExcel:NewStyle():Setfill(nCorBranc):SetFont(nAri10N):SetnumFmt(nFmtNum):Setborder(nBordaAll)

	oFundoCinza	:= oExcel:NewStyle():Setfont(nCali11N):Setfill(nCorCinza):SetaValores({oCenter})
	oFundoBranc	:= oExcel:NewStyle():Setfill(nCorBranc)

	oExcel:ADDPlan("APURACAO INCENTIVADA - "+Upper(SubSTR(MesExtenso(Val(SubSTR(SZG->ZG_PERIODO,1,2))),1,3))+"."+Alltrim(SubSTR(SZG->ZG_PERIODO,3)))
	
	oExcel:Pos(1,17):SetStyle(oFundoBranc)
    oExcel:mergeCells(1,1,5,5) //Mescla
    
    //Colunas
    oExcel:AddTamCol(1,1,6.00)
    oExcel:AddTamCol(2,2,60.00)
	oExcel:AddTamCol(3,5,12.50)

	//imagem
	If File(CurDir()+"logoprodepe_"+Alltrim(cFilAnt)+".png")
		nIDImg		:= oExcel:ADDImg(CurDir()+"logoprodepe_"+Alltrim(cFilAnt)+".png")	//Imagem no Protheus_data
		oExcel:Img(nIDImg,1,8,250,095,"px",)
	EndIf

    nLin := 6
	//Para alterações deve primeiro posicionar na Celula pelo Pos(linha,coluna) ou PosR(referencia)
	oExcel:Pos(nLin,1):SetValue(cPerFisc + " - Apuração Incentivada"):SetStyle(oFundoCinza)
	oExcel:mergeCells(nLin,1,nLin,5) 
	
    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,2)
    oExcel:mergeCells(nLin,3,7,5)
    oExcel:Pos(nLin,1):SetValue("Empresa"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetValue(cNomeEmp):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,2)
    oExcel:mergeCells(nLin,3,nLin,5)
    oExcel:Pos(nLin,1):SetValue("CACEPE nº"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)
    
    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,2)
    oExcel:mergeCells(nLin,3,nLin,5)
    oExcel:Pos(nLin,1):SetValue("CNPJ/MF nº"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetValue(cCNPJ):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Entradas"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:Pos(nLin,1):SetValue("CFOP"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,2):SetValue("NATUREZA"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,3):SetValue("VC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,4):SetValue("BC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,5):SetValue("ICMS"):SetStyle(oEstiloCent)
    
    cQry := " SELECT 'SZM' as TABELA, SZM.ZM_CFOP as CFOP, SZM.ZM_DESCFOP as DESCFOP, SZM.ZM_VALCONT as VALCONT, SZM.ZM_BASEICM as BASEICM, SZM.ZM_VALICM as VALICM" 
    cQry += " FROM "+ RetSqlName("SZM") +" SZM "
    cQry += " WHERE SZM.D_E_L_E_T_ <> '*' "
    cQry += " AND	SZM.ZM_FILIAL  = '"+FwxFilial('SZM')+"' " 
    cQry += " AND	SZM.ZM_COD     = '"+SZG->ZG_COD+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    nLin += 1
    While!(__cAlias)->(EOF())
        
        oExcel:Pos(nLin,1):SetValue((__cAlias)->CFOP):SetStyle(oEstiloB)
        oExcel:Pos(nLin,2):SetValue(Capital((__cAlias)->DESCFOP)):SetStyle(oEstiloLef)
        oExcel:Pos(nLin,3):SetValue((__cAlias)->VALCONT):SetStyle(oFmtNum)
        oExcel:Pos(nLin,4):SetValue((__cAlias)->BASEICM):SetStyle(oFmtNum)
        oExcel:Pos(nLin,5):SetValue((__cAlias)->VALICM):SetStyle(oFmtNum)

        nTotVC1 += (__cAlias)->VALCONT
        nTotBC1 += (__cAlias)->BASEICM
        nTotICM1 += (__cAlias)->VALICM
      
     nLin += 1
     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf   

    oExcel:Pos(nLin,1):SetValue("TOTAL DAS ENTRADAS"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloLefN)
    oExcel:Pos(nLin,3):SetValue(nTotVC1):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,4):SetValue(nTotBC1):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,5):SetValue(nTotICM1):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Saídas"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:Pos(nLin,1):SetValue("CFOP"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,2):SetValue("NATUREZA"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,3):SetValue("VC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,4):SetValue("BC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,5):SetValue("ICMS"):SetStyle(oEstiloCent)
    
    cQry := " SELECT 'SZO' as TABELA, SZO.ZO_CFOP as CFOP, SZO.ZO_DESCFOP as DESCFOP, SZO.ZO_VALCONT as VALCONT, SZO.ZO_BASEICM as BASEICM, SZO.ZO_VALICM as VALICM" 
    cQry += " FROM "+ RetSqlName("SZO") +" SZO "
    cQry += " WHERE SZO.D_E_L_E_T_ <> '*' "
    cQry += " AND	SZO.ZO_FILIAL  = '"+FwxFilial('SZO')+"' " 
    cQry += " AND	SZO.ZO_COD     = '"+SZG->ZG_COD+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    nLin += 1
    While!(__cAlias)->(EOF())
        
        oExcel:Pos(nLin,1):SetValue((__cAlias)->CFOP):SetStyle(oEstiloB)
        oExcel:Pos(nLin,2):SetValue(Capital((__cAlias)->DESCFOP)):SetStyle(oEstiloLef)
        oExcel:Pos(nLin,3):SetValue((__cAlias)->VALCONT):SetStyle(oFmtNum)
        oExcel:Pos(nLin,4):SetValue((__cAlias)->BASEICM):SetStyle(oFmtNum)
        oExcel:Pos(nLin,5):SetValue((__cAlias)->VALICM):SetStyle(oFmtNum)

        nTotVC2 += (__cAlias)->VALCONT
        nTotBC2 += (__cAlias)->BASEICM
        nTotICM2 += (__cAlias)->VALICM
     
     nLin += 1
     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf   

    oExcel:Pos(nLin,1):SetValue("TOTAL DAS SAÍDAS"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloLefN)
    oExcel:Pos(nLin,3):SetValue(nTotVC2):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,4):SetValue(nTotBC2):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,5):SetValue(nTotICM2):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Apuração"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("SALDO DE CRÉDITO DO PERÍODO ANTERIOR"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Créditos das Entradas"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotICM1):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SF1") +" SF1 "
    cQry += " WHERE SF1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SF1.F1_FILIAL  = '"+FwxFilial('SF1')+"' " 
    cQry += " AND	SF1.F1_ESPECIE ='NF3E' "
    cQry += " AND	SF1.F1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALMERC

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros Créditos - TRANSF DA APUR. 1 - ENERGIA"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SF1") +" SF1 "
    cQry += " WHERE SF1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SF1.F1_FILIAL  = '"+FwxFilial('SF1')+"' " 
    cQry += " AND	SF1.F1_ESPECIE ='CTE' "
    cQry += " AND	SF1.F1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALMERC

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros Créditos - TRANSF DA APUR. 1 - FRETE INTER."):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros créditos (CIAP)"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SF1") +" SF1 "
    cQry += " WHERE SF1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SF1.F1_FILIAL  = '"+FwxFilial('SF1')+"' " 
    cQry += " AND	SF1.F1_ESPECIE ='NF3E' "
    cQry += " AND	SF1.F1_EST     = '"+cEstado+"' " 
    cQry += " AND	SF1.F1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALICM

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros créditos (ICMS DA SUBVENÇÃO ENERGIA ELÉTRICA)"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SF1") +" SF1 "
    cQry += " WHERE SF1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SF1.F1_FILIAL  = '"+FwxFilial('SF1')+"' " 
    cQry += " AND	SF1.F1_ESPECIE ='NF3E' "
    cQry += " AND	SF1.F1_EST     <> '"+cEstado+"' " 
    cQry += " AND	SF1.F1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALICM

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros créditos ENERGIA ELÉTRICA EM OUTRA UF"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Estorno de Débito"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Total do crédito"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Débito das Saídas"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotICM2):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros Débitos"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Estorno de Crédito - ref. Devolução de Matéria Prima"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Total dos Débitos"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Saldo do ICMS - Apuração 03"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotICM1+nTotICM2):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Dedução FECEP"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Saldo do ICMS depois da dedução FECEP - Apuração 03"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("ESTORNO S. CREDOR A TRANSF. P/PERÍODO SEGUINTE"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Saldo do ICMS antes da apuração PRODEPE"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("CRÉDITO PRESUMIDO PRODEPE INDUSTRIA PRIORITÁRIO - 90%"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("SALDO DO ICMS DEPOIS DA APURAÇÃO PRODEPE"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Saldo do ICMS - FECEP"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("TAXA AD-DIPPER"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("FEEF"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("ICMS A RECOLHER (INC + N.INC)"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    /*-----------------------------------------------------------------------*/
    /*                                                                       */
    /*                      Segunda Planilha                                 */
    /*                                                                       */
    /*-----------------------------------------------------------------------*/

    oExcel:ADDPlan("APURACAO NAO INCENTIVADA"+Upper(SubSTR(MesExtenso(Val(SubSTR(SZG->ZG_PERIODO,1,2))),1,3))+"."+Alltrim(SubSTR(SZG->ZG_PERIODO,3)))
	
	oExcel:Pos(1,17):SetStyle(oFundoBranc)
    oExcel:mergeCells(1,1,5,5) //Mescla
    
    //Colunas
    oExcel:AddTamCol(1,1,6.00)
    oExcel:AddTamCol(2,2,60.00)
	oExcel:AddTamCol(3,5,12.50)

	//imagem
	If File(CurDir()+"logoprodepe_"+Alltrim(cFilAnt)+".png")
		nIDImg		:= oExcel:ADDImg(CurDir()+"logoprodepe_"+Alltrim(cFilAnt)+".png")	//Imagem no Protheus_data
		oExcel:Img(nIDImg,1,8,250,095,"px",)
	EndIf

    nLin := 6
	//Para alterações deve primeiro posicionar na Celula pelo Pos(linha,coluna) ou PosR(referencia)
	oExcel:Pos(nLin,1):SetValue(cPerFisc + " - Apuração Incentivada"):SetStyle(oFundoCinza)
	oExcel:mergeCells(nLin,1,nLin,5) 
	
    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,2)
    oExcel:mergeCells(nLin,3,7,5)
    oExcel:Pos(nLin,1):SetValue("Empresa"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetValue(cNomeEmp):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,2)
    oExcel:mergeCells(nLin,3,nLin,5)
    oExcel:Pos(nLin,1):SetValue("CACEPE nº"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)
    
    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,2)
    oExcel:mergeCells(nLin,3,nLin,5)
    oExcel:Pos(nLin,1):SetValue("CNPJ/MF nº"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetValue(cCNPJ):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Entradas"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:Pos(nLin,1):SetValue("CFOP"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,2):SetValue("NATUREZA"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,3):SetValue("VC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,4):SetValue("BC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,5):SetValue("ICMS"):SetStyle(oEstiloCent)
    
    cQry := " SELECT 'SZP' as TABELA, SZP.ZP_CFOP as CFOP, SZP.ZP_DESCFOP as DESCFOP, SZP.ZP_VALCONT as VALCONT, SZP.ZP_BASEICM as BASEICM, SZP.ZP_VALICM as VALICM" 
    cQry += " FROM "+ RetSqlName("SZP") +" SZP "
    cQry += " WHERE SZP.D_E_L_E_T_ <> '*' "
    cQry += " AND	SZP.ZP_FILIAL  = '"+FwxFilial('SZP')+"' " 
    cQry += " AND	SZP.ZP_COD     = '"+SZG->ZG_COD+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    nLin += 1
    While!(__cAlias)->(EOF())
        
        oExcel:Pos(nLin,1):SetValue((__cAlias)->CFOP):SetStyle(oEstiloB)
        oExcel:Pos(nLin,2):SetValue(Capital((__cAlias)->DESCFOP)):SetStyle(oEstiloLef)
        oExcel:Pos(nLin,3):SetValue((__cAlias)->VALCONT):SetStyle(oFmtNum)
        oExcel:Pos(nLin,4):SetValue((__cAlias)->BASEICM):SetStyle(oFmtNum)
        oExcel:Pos(nLin,5):SetValue((__cAlias)->VALICM):SetStyle(oFmtNum)

        nTotVC3 += (__cAlias)->VALCONT
        nTotBC3 += (__cAlias)->BASEICM
        nTotICM3 += (__cAlias)->VALICM
     
     nLin += 1
     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf   

    oExcel:Pos(nLin,1):SetValue("TOTAL DAS ENTRADAS"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloLefN)
    oExcel:Pos(nLin,3):SetValue(nTotVC3):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,4):SetValue(nTotBC3):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,5):SetValue(nTotICM3):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Saídas"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:Pos(nLin,1):SetValue("CFOP"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,2):SetValue("NATUREZA"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,3):SetValue("VC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,4):SetValue("BC"):SetStyle(oEstiloCent)
    oExcel:Pos(nLin,5):SetValue("ICMS"):SetStyle(oEstiloCent)
    
    cQry := " SELECT 'SZQ' as TABELA, SZQ.ZQ_CFOP as CFOP, SZQ.ZQ_DESCFOP as DESCFOP, SZQ.ZQ_VALCONT as VALCONT, SZQ.ZQ_BASEICM as BASEICM, SZQ.ZQ_VALICM as VALICM" 
    cQry += " FROM "+ RetSqlName("SZQ") +" SZQ "
    cQry += " WHERE SZQ.D_E_L_E_T_ <> '*' "
    cQry += " AND	SZQ.ZQ_FILIAL  = '"+FwxFilial('SZQ')+"' " 
    cQry += " AND	SZQ.ZQ_COD     = '"+SZG->ZG_COD+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    nLin += 1
    While!(__cAlias)->(EOF())
        
        oExcel:Pos(nLin,1):SetValue((__cAlias)->CFOP):SetStyle(oEstiloB)
        oExcel:Pos(nLin,2):SetValue(Capital((__cAlias)->DESCFOP)):SetStyle(oEstiloLef)
        oExcel:Pos(nLin,3):SetValue((__cAlias)->VALCONT):SetStyle(oFmtNum)
        oExcel:Pos(nLin,4):SetValue((__cAlias)->BASEICM):SetStyle(oFmtNum)
        oExcel:Pos(nLin,5):SetValue((__cAlias)->VALICM):SetStyle(oFmtNum)

        nTotVC4 += (__cAlias)->VALCONT
        nTotBC4 += (__cAlias)->BASEICM
        nTotICM4 += (__cAlias)->VALICM
     
     nLin += 1
     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf   

    oExcel:Pos(nLin,1):SetValue("TOTAL DAS SAÍDAS"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloLefN)
    oExcel:Pos(nLin,3):SetValue(nTotVC4):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,4):SetValue(nTotBC4):SetStyle(oFmtNumN)
    oExcel:Pos(nLin,5):SetValue(nTotICM4):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Apuração"):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("SALDO DE CRÉDITO DO PERÍODO ANTERIOR"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetValue("Créditos das Entradas"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotICM1):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SF1") +" SF1 "
    cQry += " WHERE SF1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SF1.F1_FILIAL  = '"+FwxFilial('SF1')+"' " 
    cQry += " AND	SF1.F1_ESPECIE ='CTE' "
    cQry += " AND	SF1.F1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALMERC

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Créditos das Entradas - Fretes Interestaduais"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SF1") +" SF1 "
    cQry += " WHERE SF1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SF1.F1_FILIAL  = '"+FwxFilial('SF1')+"' " 
    cQry += " AND	SF1.F1_ESPECIE = 'NF3E' "
    cQry += " AND	SF1.F1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALMERC

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Créditos das Entradas - Energia"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SD1") +" SD1 "
    cQry += " INNER JOIN "+ RetSqlName("SB1") +" SB1 "
    cQry += " ON SB1.B1_COD = SD1.D1_COD  "
    cQry += " WHERE SD1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SB1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SD1.D1_FILIAL  = '"+FwxFilial('SD1')+"' "
    cQry += " AND	SB1.B1_FILIAL  = '"+FwxFilial('SB1')+"' "
    cQry += " AND	SB1.B1_TIPO    = 'MP' "
    cQry += " AND	SD1.D1_TIPO    = 'D' "
    cQry += " AND	SD1.D1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALMERC

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Créditos das Entradas - Devolução de MP"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nTotCalc := 0

    cQry := " SELECT * " 
    cQry += " FROM "+ RetSqlName("SD1") +" SD1 "
    cQry += " INNER JOIN "+ RetSqlName("SB1") +" SB1 "
    cQry += " ON SB1.B1_COD = SD1.D1_COD  "
    cQry += " WHERE SD1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SB1.D_E_L_E_T_ <> '*' "
    cQry += " AND	SD1.D1_FILIAL  = '"+FwxFilial('SD1')+"' "
    cQry += " AND	SB1.B1_FILIAL  = '"+FwxFilial('SB1')+"' "
    cQry += " AND	SB1.B1_TIPO    = 'MC' "
    cQry += " AND	SD1.D1_TIPO    = 'D' "
    cQry += " AND	SD1.D1_EMISSAO BETWEEN '"+ZG_DATAINI+"' AND "+ "'"+ZG_DATAFIM+"' "
    cQry := ChangeQuery(cQry)
    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

    While!(__cAlias)->(EOF())

        nTotCalc += (__cAlias)->F1_VALMERC

     (__cAlias)->(DBSkip())
    EndDo

    IF Select(__cAlias) <> 0
        (__cAlias)->(DbCloseArea())
    EndIf

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros créditos -ICMS REF DEVOLUÇÃO DE PROD USO E CONSUMO"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros créditos- diferencial de alíquotas mercadoria adquirida para revenda"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros Créditos - CRÉDITO PRESUMIDO Decreto n° 44.650/2017, art. 17, § 1º, II, art. 21, I, art. 22, art. 58, I, § 2º, art. 75, II"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotCalc):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Estorno de Débito"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Total do crédito"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Débito das Saídas"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotICM2):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Outros Débitos (AJUSTE DE DÉBITO REF. ICMS NÃO DESTACADO ERRONEAMENTE NAS NFES N° 116466,116523,116553,116729,116804,116806,116884 e 116885"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Estorno de crédito - Frete Interestadual"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Estorno de crédito ( ENERGIA ELÉTRICA)"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Total dos Débitos"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Saldo do ICMS - Apuração 01"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(nTotICM1+nTotICM2):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("ESTORNO S. CREDOR A TRANSF. P/PERÍODO SEGUINTE"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("Saldo do ICMS - Apuração 01"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("CRÉDITO FRETES A TRANSF. P/APURAÇÃO 3"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("CRÉDITO ENERGIA  TRANSF. P/APURAÇÃO 3"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetValue(0):SetStyle(oFmtNumN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,5)
    oExcel:Pos(nLin,1):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

    nLin += 1
    oExcel:mergeCells(nLin,1,nLin,4)
    oExcel:Pos(nLin,1):SetValue("ICMS A RECOLHER (INC + N.INC)"):SetStyle(oEstiloLef)
    oExcel:Pos(nLin,2):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,3):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,4):SetStyle(oEstiloBN)
    oExcel:Pos(nLin,5):SetStyle(oEstiloBN)

	oExcel:Save(GetTempPath())
	oExcel:OpenApp()
	oExcel:Close()

Return
