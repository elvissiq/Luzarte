//Bibliotecas
#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
#Include "TBICONN.CH"
#Include "TopConn.ch"

#Define ENTER Chr(13)

Static cCampoFornec := "ZH_FORNECE;ZH_LOJA;ZH_NOME;ZH_CGC;ZH_EST;ZH_COD;"
Static cCampoPercen := "ZH_PERIODO;ZH_PAUTA;ZH_CREDIND;ZH_PRODBEN;ZH_CREDRES;ZH_VALRESS;"

//----------------------------------------------------------------------
/*/{PROTHEUS.DOC} PFISF02
FUNÇÃO PFISF02 - Apuração Ressarcimento ICMS
@OWNER Pan Cristal
@VERSION PROTHEUS 12
@SINCE 02/08/2023
@Tratamento para calculo do PRODEPE
/*/
//----------------------------------------------------------------------

User Function PFISF02()
Local aArea := GetArea()
Local oBrowse    

oBrowse := FWMBrowse():New()
oBrowse:SetAlias("SZH")
oBrowse:SetDescription("Apuração Ressarcimento ICMS")
oBrowse:Activate()
     
RestArea(aArea)
Return

/*---------------------------------------------------------------------*
 | Func:  MenuDef                                                      |
 | Desc:  Criação do menu MVC                                          |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function MenuDef()
Local aRot := {}
     
ADD OPTION aRot TITLE 'Incluir'     ACTION 'VIEWDEF.PFISF02'     OPERATION 3 ACCESS 0 //OPERATION 3
ADD OPTION aRot TITLE 'Alterar'     ACTION 'U_PFISF08("SZH",4)'  OPERATION 4 ACCESS 0 //OPERATION 4
ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.PFISF02'     OPERATION 2 ACCESS 0 //OPERATION 2
ADD OPTION aRot TITLE 'Imprimir'    ACTION 'U_PFISR02'           OPERATION 7 ACCESS 0 //OPERATION 7
ADD OPTION aRot TITLE 'Termo'       ACTION 'U_PFISRTR'           OPERATION 8 ACCESS 0 //OPERATION 8
ADD OPTION aRot TITLE 'Excluir'     ACTION 'VIEWDEF.PFISF02'     OPERATION 5 ACCESS 0 //OPERATION 5
    
Return aRot

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ModelDef()
Local oModel as object
Local oStrField as object
Local oStrEntrada as object
Local oStrSaida as object

oStrField := FWFormStruct(1, 'SZH' )
oStrEntrada := FWFormStruct(1, 'SZI')
oStrSaida := FWFormStruct(1, 'SZJ')

oModel := MPFormModel():New('PFISF02M',/*bPre*/,/*bPost*/,/*bCommit*/,/*bCancel*/)
oModel:AddFields('CABEC',/*cOwner*/,oStrField/*bPre*/,/*bPos*/,/*bLoad*/)
oModel:AddGrid('ENTRADA','CABEC',oStrEntrada,/*bLinePre*/,/*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
oModel:AddGrid('SAIDA','CABEC',oStrSaida,/*bLinePre*/,/*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
oModel:SetPrimaryKey({})

oModel:SetRelation("ENTRADA",{{"ZI_FILIAL", "xFilial('SZI')"},; 
                              {"ZI_COD", "ZH_COD"}}, SZI->(IndexKey(1)))

oModel:SetRelation("SAIDA" ,{{"ZJ_FILIAL", "xFilial('SZJ')"},; 
                             {"ZJ_COD", "ZH_COD"}}, SZJ->(IndexKey(1)))

oModel:SetDescription("Apuração Ressarcimento ICMS")

oStrField:SetProperty("ZH_COD", MODEL_FIELD_INIT, FWBuildFeature(STRUCT_FEATURE_INIPAD, "GetSXENum('SZH','ZH_COD')"))
oStrField:SetProperty("ZH_PERIODO", MODEL_FIELD_VALID, FwBuildFeature(STRUCT_FEATURE_VALID, 'ExistChav("SZH", FWFldGet("ZH_PERIODO"))'))

oModel:AddCalc( 'CALCENTRADA',  'CABEC', 'ENTRADA', 'ZI_QTDNF'  , 'QTDNF'  , 'SUM' , { || .T. },,'Total Qtd. NF'   )
oModel:AddCalc( 'CALCENTRADA',  'CABEC', 'ENTRADA', 'ZI_QTDSACO', 'QTDSACO', 'SUM' , { || .T. },,'Total Qtd. Saco' )
oModel:AddCalc( 'CALCENTRADA',  'CABEC', 'ENTRADA', 'ZI_VALOR'  , 'VALOR'  , 'SUM' , { || .T. },,'Total Valor' )
oModel:AddCalc( 'CALCENTRADA',  'CABEC', 'ENTRADA', 'ZI_CREDFIS', 'CREDFIS', 'SUM' , { || .T. },,'Total Cred. Fiscal' )
oModel:AddCalc( 'CALCENTRADA',  'CABEC', 'ENTRADA', 'ZI_VALRESS', 'VALRESS', 'SUM' , { || .T. },,'Total Vlr. Ressarcimento' )

oModel:AddCalc( 'CALCSAIDA',  'CABEC', 'SAIDA', 'ZJ_QTDNF'  , 'QTDNF'  , 'SUM' , { || .T. },,'Total Qtd. NF'   )
oModel:AddCalc( 'CALCSAIDA',  'CABEC', 'SAIDA', 'ZJ_QTDSACO', 'QTDSACO', 'SUM' , { || .T. },,'Total Qtd. Saco' )
oModel:AddCalc( 'CALCSAIDA',  'CABEC', 'SAIDA', 'ZJ_VALOR'  , 'VALOR'  , 'SUM' , { || .T. },,'Total Valor' )
oModel:AddCalc( 'CALCSAIDA',  'CABEC', 'SAIDA', 'ZJ_CREDFIS', 'CREDFIS', 'SUM' , { || .T. },,'Total Cred. Fiscal' )
oModel:AddCalc( 'CALCSAIDA',  'CABEC', 'SAIDA', 'ZJ_VALRESS', 'VALRESS', 'SUM' , { || .T. },,'Total Vlr. Ressarcimento' )

oModel:GetModel('ENTRADA'):SetOptional(.T.)
oModel:GetModel('SAIDA'):SetOptional(.T.)

Return oModel
 
/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
Local oView as object
Local oModel as object
Local oStrFornec as object
Local oStrPercen as object
Local oStrEntrada as object
Local oStrSaida as object
Local oStrCalcEnt as object
Local oStrCalcSai as object 

oModel := FWLoadModel("PFISF02")
oStrFornec  := FWFormStruct(2, 'SZH', {|cCampo| Alltrim(cCampo) $ cCampoFornec})
oStrPercen  := FWFormStruct(2, 'SZH', {|cCampo| Alltrim(cCampo) $ cCampoPercen})
oStrEntrada := FWFormStruct(2, 'SZI')
oStrSaida   := FWFormStruct(2, 'SZJ')
oStrCalcEnt := FWCalcStruct(oModel:GetModel('CALCENTRADA'))
oStrCalcSai := FWCalcStruct(oModel:GetModel('CALCSAIDA'))

oView := FwFormView():New()
oView := FWFormView():New()
oView:SetModel(oModel)
oView:SetProgressBar(.T.)
oView:AddField("VIEW_FORNEC"  , oStrFornec , "CABEC")
oView:AddField("VIEW_PERCEN"  , oStrPercen , "CABEC")
oView:AddGrid("VIEW_ENTRADA"  , oStrEntrada, "ENTRADA")
oView:AddGrid("VIEW_SAIDA"    , oStrSaida  , "SAIDA")
oView:AddField('VIEW_CALCENT' , oStrCalcEnt,'CALCENTRADA')
oView:AddField('VIEW_CALCSAI' , oStrCalcEnt,'CALCSAIDA')

oView:CreateHorizontalBox("BOX_SUPERIOR", 20 )
oView:CreateHorizontalBox("BOX_INFERIOR", 80 )
oView:CreateFolder('ABAS','BOX_INFERIOR')
oView:AddSheet('ABAS', 'ABA_1', 'Entradas')
oView:AddSheet('ABAS', 'ABA_2', 'Perda, Venda, Devolução')
oView:CreateHorizontalBox("BOX_ENTRADA", 90, /*cIdOwner*/, /*lUsePixel*/, 'ABAS', 'ABA_1')
oView:CreateHorizontalBox("BOX_SAIDA", 90, /*cIdOwner*/, /*lUsePixel*/, 'ABAS', 'ABA_2')
oView:CreateHorizontalBox('CALCENT',10, /*cIdOwner*/, /*lUsePixel*/, 'ABAS', 'ABA_1')
oView:CreateHorizontalBox('CALCSAI',10, /*cIdOwner*/, /*lUsePixel*/, 'ABAS', 'ABA_2')

oView:CreateVerticalBox("BOX_SUPERIOR_ESQUERDO", 50, "BOX_SUPERIOR")
oView:CreateVerticalBox("BOX_SUPERIOR_DIREITO" , 50, "BOX_SUPERIOR")

oView:setOwnerView("VIEW_FORNEC" , "BOX_SUPERIOR_ESQUERDO")
oView:setOwnerView("VIEW_PERCEN" , "BOX_SUPERIOR_DIREITO")
oView:setOwnerView("VIEW_ENTRADA", "BOX_ENTRADA")
oView:SetOwnerView("VIEW_CALCENT", "CALCENT")
oView:setOwnerView("VIEW_SAIDA"  , "BOX_SAIDA")
oView:SetOwnerView("VIEW_CALCSAI", "CALCSAI")

oView:AddUserButton( 'Consultar', 'MAGIC_BMP',;
                        {|| IIF(oView:GetOperation() == 3,DadosGrid(),)},;
                         /*cToolTip  | Comentário do botão*/,;
                         /*nShortCut | Codigo da Tecla para criação de Tecla de Atalho*/,;
                         /*aOptions  | */,;
                         /*lShowBar */ .T.)

oView:SetAfterViewActivate({|oView| ViewActv(oView)})

Return oView

/*---------------------------------------------------------------------*
 | Func:  DadosGrid                                                    |
 | Desc:  Realiza a carga de dados nos grids da tela.                  |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function DadosGrid()
Local oModel    := FWModelActive()
Local oView     := FWViewActive()
Local oStrField := oModel:GetModel("CABEC")
Local oStrEntra := oModel:GetModel("ENTRADA")
Local oStrSaida := oModel:GetModel("SAIDA")
Local oStruSZH  := FWFormStruct( 1, 'SZH' , /* bAvalCampo */, /* lViewUsado */ )
Local aFields   := oStruSZH:GetFields()
Local aFieldObg := {}
Local aProdutos := {}
Local cMsg      := ""
Local __cAlias  := "TEMP"+FWTimeStamp(1)
Local cQry      := ""
Local cDtIni    := DToS(FirstDate(CToD("01/"+SubSTr(FWFldGet("ZH_PERIODO"),1,2)+"/"+SubSTr(FWFldGet("ZH_PERIODO"),3))))
Local cDtFim    := DToS(LastDate(CToD("01/"+SubSTr(FWFldGet("ZH_PERIODO"),1,2)+"/"+SubSTr(FWFldGet("ZH_PERIODO"),3))))
Local nValRess  := 0
Local nY, nLinWhile, nPosProd

For nY := 1 To Len(aFields)
    If aFields[nY,10] .AND. Empty(FWFldGet(aFields[nY,3]))
         aAdd(aFieldObg,{aFields[nY,3],aFields[nY,1]})
    EndIf 
Next nY 

If !Empty(aFieldObg)
    
    cMsg += 'O(s) campo(s): '+ENTER
    For nY := 1 To Len(aFieldObg)
        cMsg += Alltrim(aFieldObg[nY,1])+" - "+Alltrim(aFieldObg[nY,2])+ENTER
    Next nY
    cMsg += ENTER+"Deve ser preenchido(s)."

    FWAlertHelp('Existem campos obrigatórios não preechidos.',cMsg)
    
    Return

EndIf 

DBSelectArea("SZH")
SZH->(DBSetOrder(3))
If SZH->(MSSeek(FWxFilial("SZH")+Pad(FWFldGet("ZH_FORNECE"),TamSX3("ZH_FORNECE")[1])+;
                                 Pad(FWFldGet("ZH_LOJA"),TamSX3("ZH_LOJA")[1])+;
                                 Pad(FWFldGet("ZH_PERIODO"),TamSX3("ZH_PERIODO")[1])))
    
    FWAlertWarning("Já existe o mesmo periodo cadastrado para este fornecedor.")

    oStrField:LoadValue("ZH_VALRESS",   0)

    oStrEntra:ClearData(.T.)
    oStrSaida:ClearData(.T.)

    oModel:GetModel('CALCENTRADA'):LoadValue("QTDNF",   0)
    oModel:GetModel('CALCENTRADA'):LoadValue("QTDSACO", 0)
    oModel:GetModel('CALCENTRADA'):LoadValue("VALOR",   0)
    oModel:GetModel('CALCENTRADA'):LoadValue("CREDFIS", 0)
    oModel:GetModel('CALCENTRADA'):LoadValue("VALRESS", 0)

    oModel:GetModel('CALCSAIDA'):LoadValue("QTDNF",   0)
    oModel:GetModel('CALCSAIDA'):LoadValue("QTDSACO", 0)
    oModel:GetModel('CALCSAIDA'):LoadValue("VALOR",   0)
    oModel:GetModel('CALCSAIDA'):LoadValue("CREDFIS", 0)
    oModel:GetModel('CALCSAIDA'):LoadValue("VALRESS", 0)
    
    oView:Refresh()

    Return   
EndIf

oStrEntra:ClearData(.T.)
oStrSaida:ClearData(.T.)

oStrField:LoadValue("ZH_VALRESS",   0)

oModel:GetModel('CALCENTRADA'):LoadValue("QTDNF",   0)
oModel:GetModel('CALCENTRADA'):LoadValue("QTDSACO", 0)
oModel:GetModel('CALCENTRADA'):LoadValue("VALOR",   0)
oModel:GetModel('CALCENTRADA'):LoadValue("CREDFIS", 0)
oModel:GetModel('CALCENTRADA'):LoadValue("VALRESS", 0)

oModel:GetModel('CALCSAIDA'):LoadValue("QTDNF",   0)
oModel:GetModel('CALCSAIDA'):LoadValue("QTDSACO", 0)
oModel:GetModel('CALCSAIDA'):LoadValue("VALOR",   0)
oModel:GetModel('CALCSAIDA'):LoadValue("CREDFIS", 0)
oModel:GetModel('CALCSAIDA'):LoadValue("VALRESS", 0)
    
oView:Refresh()

cQry := " SELECT    SD1.D1_FILIAL, "
cQry += " 		    SD1.D1_DOC, "
cQry += " 		    SD1.D1_SERIE, " 
cQry += " 		    SD1.D1_DTDIGIT, " 
cQry += " 		    SUM(SD1.D1_QUANT) AS D1_QUANT, "
cQry += " 		    SUM(SD1.D1_TOTAL) AS D1_TOTAL, "
cQry += " 		    SUM(SD1.D1_VALICM) AS D1_VALICM, " 
cQry += " 		    SF1.F1_CHVNFE "
cQry += " FROM "+ RetSqlName("SD1") +" SD1 "
cQry += " INNER JOIN "+ RetSqlName("SF1") +" SF1 ON SF1.F1_FILIAL = SD1.D1_FILIAL "
cQry += " 					  AND SF1.F1_DOC = SD1.D1_DOC "
cQry += " 					  AND SF1.F1_SERIE = SD1.D1_SERIE "
cQry += " 					  AND SF1.F1_FORNECE = SD1.D1_FORNECE "
cQry += " 					  AND SF1.F1_LOJA = SD1.D1_LOJA "
cQry += " WHERE SD1.D_E_L_E_T_ <> '*' "
cQry += " AND	SF1.D_E_L_E_T_ <> '*' "
cQry += " AND	SD1.D1_FILIAL  = '"+FwxFilial('SD1')+"' " 
cQry += " AND 	SD1.D1_FORNECE = '"+FWFldGet("ZH_FORNECE")+"' "
cQry += " AND	SD1.D1_LOJA    = '"+FWFldGet("ZH_LOJA")+"' "
cQry += " AND 	SD1.D1_EMISSAO BETWEEN '"+ cDtIni +"' AND '"+ cDtFim +"' "
cQry += " GROUP BY D1_FILIAL, D1_DOC, D1_SERIE, D1_DTDIGIT, F1_CHVNFE "
cQry := ChangeQuery(cQry)
IF Select(__cAlias) <> 0
    (__cAlias)->(DbCloseArea())
EndIf
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

nLinWhile := 0

DBSelectArea("SD1")
SD1->(DBSetOrder(1))

While!(__cAlias)->(EOF())

    nLinWhile++
    If nLinWhile > 1
        oStrEntra:AddLine()
    EndIF

    oStrEntra:LoadValue("ZI_COD"     , FWFldGet("ZH_COD"))
    oStrEntra:LoadValue("ZI_FORNECE" , FWFldGet("ZH_FORNECE"))
    oStrEntra:LoadValue("ZI_LOJA"    , FWFldGet("ZH_LOJA"))
    oStrEntra:LoadValue("ZI_PERIODO" , FWFldGet("ZH_PERIODO"))
    oStrEntra:LoadValue("ZI_DOC"     , (__cAlias)->D1_DOC)
    oStrEntra:LoadValue("ZI_SERIE"   , (__cAlias)->D1_SERIE)
    oStrEntra:LoadValue("ZI_DTDIGIT" , SToD((__cAlias)->D1_DTDIGIT))
    oStrEntra:LoadValue("ZI_QTDKG"   , 50.00)
    oStrEntra:LoadValue("ZI_QTDKGNF" , 1.00)
    oStrEntra:LoadValue("ZI_QTDNF"   , (__cAlias)->D1_QUANT)
    oStrEntra:LoadValue("ZI_QTDSACO" , (__cAlias)->D1_QUANT/50)
    oStrEntra:LoadValue("ZI_VALOR"   , (__cAlias)->D1_TOTAL)
    oStrEntra:LoadValue("ZI_ICMSACO" , FWFldGet("ZH_PAUTA"))
    oStrEntra:LoadValue("ZI_CREDFIS" , FWFldGet("ZI_QTDSACO") * FWFldGet("ZI_ICMSACO"))
    oStrEntra:LoadValue("ZI_VALRESS" ,    FWFldGet("ZI_CREDFIS");
                                        * (FWFldGet("ZH_CREDIND")/100);
                                        * (FWFldGet("ZH_PRODBEN")/100);
                                        * (FWFldGet("ZH_CREDRES")/100))
    oStrEntra:LoadValue("ZI_CHAVE"   , (__cAlias)->F1_CHVNFE)

    oView:Refresh("VIEW_ENTRADA")
    oView:Refresh("VIEW_CALCENT")

    nValRess += oStrEntra:GetValue("ZI_VALRESS")

    //Código dos produtos da Nota Fiscal para serem utilizados na próxima Query (Perda/Venda/Devoluções)
    If SD1->(MSSeek(FwxFilial('SD1')+(__cAlias)->D1_DOC+(__cAlias)->D1_SERIE+FWFldGet("ZH_FORNECE")+FWFldGet("ZH_LOJA")))
        
        While ! SD1->(EOF()) .AND. SD1->D1_FILIAL == FwxFilial('SD1') ;
                .AND. SD1->D1_DOC == (__cAlias)->D1_DOC .AND. SD1->D1_SERIE == (__cAlias)->D1_SERIE ;
                .AND. SD1->D1_FORNECE == FWFldGet("ZH_FORNECE") .AND. SD1->D1_LOJA == FWFldGet("ZH_LOJA")
            
            nPosProd := aScan(aProdutos, {|x| AllTrim(x) == Alltrim(SD1->D1_COD)})

            If Empty(nPosProd)
                aAdd(aProdutos,Alltrim(SD1->D1_COD))
            EndIf 

        SD1->(DBSkip())
        EndDo
    EndIF 

(__cAlias)->(DBSkip())
EndDo

If Select(__cAlias) > 0                                 
    (__cAlias)->(dbCloseArea())
EndIf

cQry := " SELECT    SD2.D2_FILIAL, "
cQry += " 		    SD2.D2_DOC, "
cQry += " 		    SD2.D2_SERIE, " 
cQry += " 		    SD2.D2_DTDIGIT, " 
cQry += " 		    SUM(SD2.D2_QUANT) AS D2_QUANT, "
cQry += " 		    SUM(SD2.D2_TOTAL) AS D2_TOTAL, "
cQry += " 		    SUM(SD2.D2_VALICM) AS D2_VALICM, " 
cQry += " 		    SF2.F2_CHVNFE "
cQry += " FROM "+ RetSqlName("SD2") +" SD2 "
cQry += " INNER JOIN "+ RetSqlName("SF2") +" SF2 ON SF2.F2_FILIAL = SD2.D2_FILIAL "
cQry += " 					  AND SF2.F2_DOC = SD2.D2_DOC "
cQry += " 					  AND SF2.F2_SERIE = SD2.D2_SERIE "
cQry += " 					  AND SF2.F2_CLIENTE = SD2.D2_CLIENTE "
cQry += " 					  AND SF2.F2_LOJA = SD2.D2_LOJA "
cQry += " WHERE SD2.D_E_L_E_T_ <> '*' "
cQry += " AND	SF2.D_E_L_E_T_ <> '*' "
cQry += " AND	SF2.F2_TIPO = 'D' "
cQry += " AND	SD2.D2_FILIAL  = '"+FwxFilial('SD2')+"' " 
cQry += " AND 	SD2.D2_EMISSAO BETWEEN '"+ cDtIni +"' AND '"+ cDtFim +"' "
cQry += " AND	SD2.D2_COD  IN "+FormatIn(ArrTokStr(aProdutos, "/"), "/")+" "
cQry += " GROUP BY D2_FILIAL, D2_DOC, D2_SERIE, D2_DTDIGIT, F2_CHVNFE "
cQry := ChangeQuery(cQry)
IF Select(__cAlias) <> 0
    (__cAlias)->(DbCloseArea())
EndIf
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

nLinWhile := 0

While !(__cAlias)->(EOF())

    nLinWhile++
    If nLinWhile > 1
        oStrSaida:AddLine()
    EndIF

    oStrSaida:LoadValue("ZJ_COD"     , FWFldGet("ZH_COD"))
    oStrSaida:LoadValue("ZJ_PERIODO" , FWFldGet("ZH_PERIODO"))
    oStrSaida:LoadValue("ZJ_DOC"     , (__cAlias)->D2_DOC)
    oStrSaida:LoadValue("ZJ_SERIE"   , (__cAlias)->D2_SERIE)
    oStrSaida:LoadValue("ZJ_DTDIGIT" , (__cAlias)->D2_DTDIGIT)
    oStrSaida:LoadValue("ZJ_QTDKG"   , 50.00)
    oStrSaida:LoadValue("ZJ_QTDKGNF" , 1.00)
    oStrSaida:LoadValue("ZJ_QTDNF"   , (__cAlias)->D2_QUANT)
    oStrSaida:LoadValue("ZJ_QTDSACO" , (__cAlias)->D2_QUANT/50)
    oStrSaida:LoadValue("ZJ_VALOR"   , (__cAlias)->D2_TOTAL)
    oStrSaida:LoadValue("ZJ_ICMSACO" , FWFldGet("ZH_PAUTA"))
    oStrSaida:LoadValue("ZJ_CREDFIS" , FWFldGet("ZJ_QTDSACO") * FWFldGet("ZJ_ICMSACO"))
    oStrSaida:LoadValue("ZJ_VALRESS" ,    FWFldGet("ZJ_CREDFIS");
                                        * (FWFldGet("ZH_CREDIND")/100);
                                        * (FWFldGet("ZH_PRODBEN")/100);
                                        * (FWFldGet("ZH_CREDRES")/100))
    oStrSaida:LoadValue("ZJ_CHAVE"   , (__cAlias)->F2_CHVNFE)

    oView:Refresh("VIEW_SAIDA")
    oView:Refresh("VIEW_CALCSAI")
    
    nValRess -= oStrSaida:GetValue("ZJ_VALRESS")

(__cAlias)->(DBSkip())
EndDo

oStrField:LoadValue("ZH_VALRESS" , nValRess)

oStrEntra:GoLine(1)
oStrSaida:GoLine(1)
oView:Refresh()

If Select(__cAlias) > 0                                 
    (__cAlias)->(dbCloseArea())
EndIf

Return

/*---------------------------------------------------------------------*
 | Func:  ViewActv                                                     |
 | Desc:  Realiza o PUT nos campos para gravação na tabela SE1         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ViewActv(oView)
Local oModel := FWModelActive() 

    If oModel:GetOperation() == 3
        oModel:GetModel("CABEC"):LoadValue("ZH_CREDIND", SuperGetMV("PC_PERCRED",.F.,60))
        oModel:GetModel("CABEC"):LoadValue("ZH_CREDRES", SuperGetMV("PC_CREDRES",.F.,90))
        oView:Refresh('VIEW_FORNEC')
        oView:Refresh('VIEW_PERCEN')
    EndIf 

Return
