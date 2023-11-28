//Bibliotecas
#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
#Include "TBICONN.CH"
#Include "TopConn.ch"

#Define ENTER Chr(13)
Static cTabela := 'SZG'

//----------------------------------------------------------------------
/*/{PROTHEUS.DOC} PFISF01
FUNÇÃO PFISF01 - Tela resumo de operações por CFOP
@OWNER Pan Cristal
@VERSION PROTHEUS 12
@SINCE 24/11/2023
@Tratamento para calculo do PRODEPE
/*/
//----------------------------------------------------------------------

User Function PFISF01()
Local aArea := GetArea()
Local oBrowse    

oBrowse := FWMBrowse():New()
oBrowse:SetAlias(cTabela)
oBrowse:SetDescription("Operações por CFOP")
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
     
ADD OPTION aRot TITLE 'Incluir'     ACTION 'VIEWDEF.PFISF01'     OPERATION 3 ACCESS 0 //OPERATION 3
ADD OPTION aRot TITLE 'Alterar'     ACTION 'VIEWDEF.PFISF01'     OPERATION 4 ACCESS 0 //OPERATION 4
ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.PFISF01'     OPERATION 2 ACCESS 0 //OPERATION 2
ADD OPTION aRot TITLE 'Excluir'     ACTION 'VIEWDEF.PFISF01'     OPERATION 5 ACCESS 0 //OPERATION 5
ADD OPTION aRot TITLE 'Imprimir'    ACTION 'U_PFISR01'           OPERATION 7 ACCESS 0 //OPERATION 7
    
Return aRot

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ModelDef()
Local oModel as object
Local oStrField as object
Local oStrSEnt as object
Local oStrSSai as object
Local oStrNENT as object
Local oStrNSAI as object

oStrField := FWFormStruct(1, 'SZG')
oStrSEnt := FWFormStruct(1, 'SZM')
oStrSSai := FWFormStruct(1, 'SZO')
oStrNENT := FWFormStruct(1, 'SZP')
oStrNSAI := FWFormStruct(1, 'SZQ')

oStrField:SetProperty("ZG_COD", MODEL_FIELD_INIT, FWBuildFeature(STRUCT_FEATURE_INIPAD, "GetSXENum('SZG','ZG_COD')"))
oStrField:SetProperty("ZG_DATAFIM", MODEL_FIELD_VALID, FwBuildFeature(STRUCT_FEATURE_VALID, 'IIF(SubStr(DToC(FWFldGet("ZG_DATAINI")),4) == SubStr(DToC(FWFldGet("ZG_DATAFIM")),4),.T.,.F.)'))
oStrField:AddTrigger("ZG_DATAFIM"  ,"ZG_PERIODO"  ,{ || IIF(SubStr(DToC(FWFldGet("ZG_DATAINI")),4) == SubStr(DToC(FWFldGet("ZG_DATAFIM")),4),.T.,.F.)},;
                                                   { || SubStr(DToC(FWFldGet("ZG_DATAFIM")),4)})

oModel := MPFormModel():New('PFISF01M',/*bPre*/,/*bPost*/,/*bCommit*/,/*bCancel*/)
oModel:AddFields('CABEC',/*cOwner*/,oStrField/*bPre*/,/*bPos*/,/*bLoad*/)
oModel:AddGrid('SENT','CABEC',oStrSEnt,/*bLinePre*/,/*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
oModel:AddGrid('SSAI','CABEC',oStrSSai,/*bLinePre*/,/*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
oModel:AddGrid('NENT','CABEC',oStrNENT,/*bLinePre*/,/*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
oModel:AddGrid('NSAI','CABEC',oStrNSAI,/*bLinePre*/,/*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)

oModel:SetRelation("SENT", {{"ZM_FILIAL", "xFilial('SZM')"},; 
                             {"ZM_COD", "ZG_COD"}}, SZM->(IndexKey(1)))
oModel:SetRelation("SSAI", {{"ZO_FILIAL", "xFilial('SZO')"},; 
                             {"ZO_COD", "ZG_COD"}}, SZO->(IndexKey(1)))
oModel:SetRelation("NENT", {{"ZP_FILIAL", "xFilial('SZP')"},; 
                             {"ZP_COD", "ZG_COD"}}, SZP->(IndexKey(1)))
oModel:SetRelation("NSAI", {{"ZQ_FILIAL", "xFilial('SZQ')"},; 
                             {"ZQ_COD", "ZG_COD"}}, SZQ->(IndexKey(1)))

oModel:SetPrimaryKey({})

oModel:AddCalc( 'CALCULOSENT',  'CABEC', 'SENT', 'ZM_VALCONT' , 'VALCONT' , 'SUM' , { || .T. },,'Total VC'   )
oModel:AddCalc( 'CALCULOSENT',  'CABEC', 'SENT', 'ZM_BASEICM' , 'BASEICM' , 'SUM' , { || .T. },,'Total BC'   )
oModel:AddCalc( 'CALCULOSENT',  'CABEC', 'SENT', 'ZM_VALICM'  , 'VALICM'  , 'SUM' , { || .T. },,'Total ICMS' )

oModel:AddCalc( 'CALCULOSSAI',  'CABEC', 'SSAI', 'ZO_VALCONT' , 'VALCONT' , 'SUM' , { || .T. },,'Total VC'   )
oModel:AddCalc( 'CALCULOSSAI',  'CABEC', 'SSAI', 'ZO_BASEICM' , 'BASEICM' , 'SUM' , { || .T. },,'Total BC'   )
oModel:AddCalc( 'CALCULOSSAI',  'CABEC', 'SSAI', 'ZO_VALICM'  , 'VALICM'  , 'SUM' , { || .T. },,'Total ICMS' )

oModel:AddCalc( 'CALCULONENT',  'CABEC', 'NENT', 'ZP_VALCONT' , 'VALCONT' , 'SUM' , { || .T. },,'Total VC'   )
oModel:AddCalc( 'CALCULONENT',  'CABEC', 'NENT', 'ZP_BASEICM' , 'BASEICM' , 'SUM' , { || .T. },,'Total BC'   )
oModel:AddCalc( 'CALCULONENT',  'CABEC', 'NENT', 'ZP_VALICM'  , 'VALICM'  , 'SUM' , { || .T. },,'Total ICMS' )

oModel:AddCalc( 'CALCULONSAI',  'CABEC', 'NSAI', 'ZQ_VALCONT' , 'VALCONT' , 'SUM' , { || .T. },,'Total VC'   )
oModel:AddCalc( 'CALCULONSAI',  'CABEC', 'NSAI', 'ZQ_BASEICM' , 'BASEICM' , 'SUM' , { || .T. },,'Total BC'   )
oModel:AddCalc( 'CALCULONSAI',  'CABEC', 'NSAI', 'ZQ_VALICM'  , 'VALICM'  , 'SUM' , { || .T. },,'Total ICMS' )

oModel:SetDescription("Movimentos por CFOP")

oModel:GetModel('SENT'):SetOptional(.T.)
oModel:GetModel('SSAI'):SetOptional(.T.)
oModel:GetModel('NENT'):SetOptional(.T.)
oModel:GetModel('NSAI'):SetOptional(.T.)

Return oModel

/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
Local oView as object
Local oModel as object
Local oStrCab as object
Local oStrSEnt as object
Local oStrNSAI as object
Local oStrNENT as object
Local oStrCalcSENT as object
Local oStrCalcSSAI as object
Local oStrCalcNENT as object

oModel := FWLoadModel("PFISF01")
oStrCab := FWFormStruct(2, 'SZG')
oStrSEnt := FWFormStruct(2, 'SZM')
oStrSSai := FWFormStruct(2, 'SZO')
oStrNENT := FWFormStruct(2, 'SZP')
oStrNSAI := FWFormStruct(2, 'SZQ')
oStrCalcSENT := FWCalcStruct(oModel:GetModel('CALCULOSENT'))
oStrCalcSSAI := FWCalcStruct(oModel:GetModel('CALCULOSSAI'))
oStrCalcNENT := FWCalcStruct(oModel:GetModel('CALCULONENT'))
oStrCalcNSAI := FWCalcStruct(oModel:GetModel('CALCULONSAI'))

oView := FwFormView():New()
oView:SetModel(oModel)
oView:SetProgressBar(.T.)
oView:AddField("VIEW_MASTER", oStrCab,"CABEC")
oView:AddGrid("VIEW_SENT"  , oStrSEnt,"SENT")
oView:AddGrid("VIEW_SSAI"  , oStrSSai,"SSAI")
oView:AddGrid("VIEW_NENT"  , oStrNENT,"NENT")
oView:AddGrid("VIEW_NSAI"  , oStrNSAI,"NSAI")
oView:AddField('VIEW_CALCSENT' , oStrCalcSENT,'CALCULOSENT')
oView:AddField('VIEW_CALCSSAI' , oStrCalcSSAI,'CALCULOSSAI')
oView:AddField('VIEW_CALCNENT' , oStrCalcNENT,'CALCULONENT')
oView:AddField('VIEW_CALCNSAI' , oStrCalcNENT,'CALCULONSAI')

oView:CreateHorizontalBox("BOX_CABEC", 10 )
oView:CreateHorizontalBox("BOX_MEIO" , 45 )
oView:CreateHorizontalBox("BOX_BAIXO", 45 )

oView:CreateFolder('ABAMEIO','BOX_MEIO')
oView:AddSheet('ABAMEIO', 'ABAMEIO_1', 'Entradas Incentivadas')
oView:AddSheet('ABAMEIO', 'ABAMEIO_2', 'Saidas Incentivadas')
oView:CreateHorizontalBox("BOX_SENT", 80, /*cIdOwner*/, /*lUsePixel*/, 'ABAMEIO', 'ABAMEIO_1')
oView:CreateHorizontalBox('CALCSENT', 20, /*cIdOwner*/, /*lUsePixel*/, 'ABAMEIO', 'ABAMEIO_1')
oView:CreateHorizontalBox("BOX_SSAI", 80, /*cIdOwner*/, /*lUsePixel*/, 'ABAMEIO', 'ABAMEIO_2')
oView:CreateHorizontalBox('CALCSSAI', 20, /*cIdOwner*/, /*lUsePixel*/, 'ABAMEIO', 'ABAMEIO_2')

oView:CreateFolder('ABABAIXO','BOX_BAIXO')
oView:AddSheet('ABABAIXO', 'ABABAIXO_1', 'Entradas Não Incentivadas')
oView:AddSheet('ABABAIXO', 'ABABAIXO_2', 'Saidas Não Incentivadas')
oView:CreateHorizontalBox("BOX_NENT", 80, /*cIdOwner*/, /*lUsePixel*/, 'ABABAIXO', 'ABABAIXO_1')
oView:CreateHorizontalBox('CALCNENT', 20, /*cIdOwner*/, /*lUsePixel*/, 'ABABAIXO', 'ABABAIXO_1')
oView:CreateHorizontalBox("BOX_NSAI", 80, /*cIdOwner*/, /*lUsePixel*/, 'ABABAIXO', 'ABABAIXO_2')
oView:CreateHorizontalBox('CALCNSAI', 20, /*cIdOwner*/, /*lUsePixel*/, 'ABABAIXO', 'ABABAIXO_2')


oView:setOwnerView("VIEW_MASTER", "BOX_CABEC")
oView:setOwnerView("VIEW_SENT", "BOX_SENT")
oView:SetOwnerView('VIEW_CALCSENT' ,'CALCSENT')
oView:setOwnerView("VIEW_SSAI", "BOX_SSAI")
oView:SetOwnerView('VIEW_CALCSSAI' ,'CALCSSAI')
oView:setOwnerView("VIEW_NENT", "BOX_NENT")
oView:SetOwnerView('VIEW_CALCNENT' ,'CALCNENT')
oView:setOwnerView("VIEW_NSAI", "BOX_NSAI")
oView:SetOwnerView('VIEW_CALCNSAI' ,'CALCNSAI')

oView:setDescription( "" )

oView:AddUserButton( 'Consultar', 'MAGIC_BMP',;
                        {|| IIF(oView:GetOperation() == 3,DadosGrid(),)},;
                         /*cToolTip  | Comentário do botão*/,;
                         /*nShortCut | Codigo da Tecla para criação de Tecla de Atalho*/,;
                         /*aOptions  | */,;
                         /*lShowBar */ .T.)

Return oView

/*---------------------------------------------------------------------*
 | Func:  DadosGrid                                                    |
 | Desc:  Realiza a carga de dados nos grids da tela.                  |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function DadosGrid()
Local oModel      := FWModelActive()
Local oView       := FWViewActive()
Local oModelCabec := oModel:GetModel("CABEC")
Local oModelSENT  := oModel:GetModel("SENT")
Local oModelSSAI  := oModel:GetModel("SSAI")
Local oModelNENT  := oModel:GetModel("NENT")
Local oModelNSAI  := oModel:GetModel("NSAI")
Local oStruSZG    := FWFormStruct( 1, 'SZG' , /* bAvalCampo */, /* lViewUsado */ )
Local aFields     := oStruSZG:GetFields()
Local aFieldObg   := {}
Local cMsg        := ""
Local __cAlias    := "TEMP"+FWTimeStamp(1)
Local cQry        := ""
Local nY

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

DBSelectArea(cTabela)
SZG->(DBSetOrder(2))
If SZG->(MSSeek(FWxFilial(cTabela)+Pad(SubStr(DToC(FWFldGet("ZG_DATAFIM")),4),TamSX3("ZG_PERIODO")[1])))
    
    FWAlertWarning("Já existe o mesmo periodo cadastrado.")

    oModelSENT:ClearData(.T.)
    oModelSSAI:ClearData(.T.)
    oModelNENT:ClearData(.T.)

    oModel:GetModel('CALCULOSENT'):LoadValue("VALCONT", 0)
    oModel:GetModel('CALCULOSENT'):LoadValue("BASEICM", 0)
    oModel:GetModel('CALCULOSENT'):LoadValue("VALICM" , 0)
    oModel:GetModel('CALCULOSSAI'):LoadValue("VALCONT", 0)
    oModel:GetModel('CALCULOSSAI'):LoadValue("BASEICM", 0)
    oModel:GetModel('CALCULOSSAI'):LoadValue("VALICM" , 0)
    oModel:GetModel('CALCULONENT'):LoadValue("VALCONT", 0)
    oModel:GetModel('CALCULONENT'):LoadValue("BASEICM", 0)
    oModel:GetModel('CALCULONENT'):LoadValue("VALICM" , 0)
    oModel:GetModel('CALCULONSAI'):LoadValue("VALCONT", 0)
    oModel:GetModel('CALCULONSAI'):LoadValue("BASEICM", 0)
    oModel:GetModel('CALCULONSAI'):LoadValue("VALICM" , 0)

    oView:Refresh()

    Return
Else
    oModelCabec:LoadValue("ZG_PERIODO", SubStr(DToC(FWFldGet("ZG_DATAFIM")),4))    
EndIf

oModelSENT:ClearData(.T.)
oModelSSAI:ClearData(.T.)
oModelNENT:ClearData(.T.)

oModel:GetModel('CALCULOSENT'):LoadValue("VALCONT", 0)
oModel:GetModel('CALCULOSENT'):LoadValue("BASEICM", 0)
oModel:GetModel('CALCULOSENT'):LoadValue("VALICM" , 0)
oModel:GetModel('CALCULOSSAI'):LoadValue("VALCONT", 0)
oModel:GetModel('CALCULOSSAI'):LoadValue("BASEICM", 0)
oModel:GetModel('CALCULOSSAI'):LoadValue("VALICM" , 0)
oModel:GetModel('CALCULONENT'):LoadValue("VALCONT", 0)
oModel:GetModel('CALCULONENT'):LoadValue("BASEICM", 0)
oModel:GetModel('CALCULONENT'):LoadValue("VALICM" , 0)
oModel:GetModel('CALCULONSAI'):LoadValue("VALCONT", 0)
oModel:GetModel('CALCULONSAI'):LoadValue("BASEICM", 0)
oModel:GetModel('CALCULONSAI'):LoadValue("VALICM" , 0)

cQry := " SELECT FT_FILIAL, FT_CFOP, SUM(FT_VALCONT) AS FT_VALCONT, SUM(FT_BASEICM) AS FT_BASEICM, SUM(FT_VALICM) AS FT_VALICM "
cQry += " FROM "+ RetSqlName("SFT") + " SFT "
cQry += " INNER JOIN "+ RetSqlName("SZX") +" SZX ON SZX.ZX_PRODUTO = SFT.FT_PRODUTO "
cQry += " INNER JOIN "+ RetSqlName("SZY") +" SZY ON SZY.ZY_DECRETO = SZX.ZX_DECRET "
cQry += " WHERE SFT.D_E_L_E_T_ <> '*' " 
cQry += " AND   SZX.D_E_L_E_T_ <> '*' "
cQry += " AND   SZY.D_E_L_E_T_ <> '*' "
cQry += " AND   SZX.ZX_FILIAL  = '"+FwxFilial('SZX')+"' "
cQry += " AND   SZX.ZX_INDESP = 'S'  "
cQry += " AND   SZY.ZY_FILIAL  = '"+FwxFilial('SZY')+"' "
cQry += " AND   SZY.ZY_DATAINI <= '"+DToS(FWFldGet("ZG_DATAINI"))+"' "
cQry += " AND   SZY.ZY_DATAFIN >= '"+DToS(FWFldGet("ZG_DATAFIM"))+"' "
cQry += " AND   SZY.ZY_ATIVO  = 'T' "
cQry += " AND   SFT.FT_FILIAL  = '"+FwxFilial('SFT')+"' "
cQry += " AND   SFT.FT_EMISSAO BETWEEN "+ DToS(FWFldGet("ZG_DATAINI")) +" AND "+ DToS(FWFldGet("ZG_DATAFIM"))
cQry += " AND   SFT.FT_TIPOMOV = 'E' "
cQry += " AND   SFT.FT_DTCANC = '' "
cQry += " GROUP BY FT_FILIAL, FT_CFOP "
cQry := ChangeQuery(cQry)
IF Select(__cAlias) <> 0
    (__cAlias)->(DbCloseArea())
EndIf
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

nY := 0

While!(__cAlias)->(EOF())
    
    nY++
    If nY > oModelSENT:Length()
        oModelSENT:AddLine()
    EndIF

    oModelSENT:LoadValue("ZM_FILIAL"  , FwxFilial("SZM"))
    oModelSENT:LoadValue("ZM_COD"     , FWFldGet("ZG_COD"))
    oModelSENT:LoadValue("ZM_CFOP"    , (__cAlias)->FT_CFOP)
    oModelSENT:LoadValue("ZM_DESCFOP" , FWGetSX5('13',(__cAlias)->FT_CFOP,'pt-br')[1][4])
    oModelSENT:LoadValue("ZM_VALCONT" , (__cAlias)->FT_VALCONT)
    oModelSENT:LoadValue("ZM_BASEICM" , (__cAlias)->FT_BASEICM)
    oModelSENT:LoadValue("ZM_VALICM"  , (__cAlias)->FT_VALICM)    

    oView:Refresh("VIEW_SENT")
    oView:Refresh("VIEW_CALCSENT")

(__cAlias)->(DBSkip())
EndDo

If Select(__cAlias) > 0                                 
    (__cAlias)->(dbCloseArea())
EndIf

cQry := " SELECT FT_FILIAL, FT_CFOP, SUM(FT_VALCONT) AS FT_VALCONT, SUM(FT_BASEICM) AS FT_BASEICM, SUM(FT_VALICM) AS FT_VALICM "
cQry += " FROM "+ RetSqlName("SFT") + " SFT "
cQry += " INNER JOIN "+ RetSqlName("SZX") +" SZX ON SZX.ZX_PRODUTO = SFT.FT_PRODUTO "
cQry += " INNER JOIN "+ RetSqlName("SZY") +" SZY ON SZY.ZY_DECRETO = SZX.ZX_DECRET "
cQry += " WHERE SFT.D_E_L_E_T_ <> '*' " 
cQry += " AND   SZX.D_E_L_E_T_ <> '*' "
cQry += " AND   SZY.D_E_L_E_T_ <> '*' "
cQry += " AND   SZX.ZX_FILIAL  = '"+FwxFilial('SZX')+"' "
cQry += " AND   SZX.ZX_INDESP = 'S'  "
cQry += " AND   SZY.ZY_FILIAL  = '"+FwxFilial('SZY')+"' "
cQry += " AND   SZY.ZY_DATAINI <= '"+DToS(FWFldGet("ZG_DATAINI"))+"' "
cQry += " AND   SZY.ZY_DATAFIN >= '"+DToS(FWFldGet("ZG_DATAFIM"))+"' "
cQry += " AND   SZY.ZY_ATIVO  = 'T' "
cQry += " AND   SFT.FT_FILIAL  = '"+FwxFilial('SFT')+"' "
cQry += " AND   SFT.FT_EMISSAO BETWEEN "+ DToS(FWFldGet("ZG_DATAINI")) +" AND "+ DToS(FWFldGet("ZG_DATAFIM"))
cQry += " AND   SFT.FT_TIPOMOV = 'S' "
cQry += " AND   SFT.FT_DTCANC = '' "
cQry += " GROUP BY FT_FILIAL, FT_CFOP "
cQry := ChangeQuery(cQry)
IF Select(__cAlias) <> 0
    (__cAlias)->(DbCloseArea())
EndIf
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

nY := 0

While!(__cAlias)->(EOF())
    
    nY++
    If nY > oModelSSAI:Length()
        oModelSSAI:AddLine()
    EndIF

    oModelSSAI:LoadValue("ZO_FILIAL"  , FwxFilial("SZO"))
    oModelSSAI:LoadValue("ZO_COD"     , FWFldGet("ZG_COD"))
    oModelSSAI:LoadValue("ZO_CFOP"    , (__cAlias)->FT_CFOP)
    oModelSSAI:LoadValue("ZO_DESCFOP" , FWGetSX5('13',(__cAlias)->FT_CFOP,'pt-br')[1][4])
    oModelSSAI:LoadValue("ZO_VALCONT" , (__cAlias)->FT_VALCONT)
    oModelSSAI:LoadValue("ZO_BASEICM" , (__cAlias)->FT_BASEICM)
    oModelSSAI:LoadValue("ZO_VALICM"  , (__cAlias)->FT_VALICM)

    oView:Refresh("VIEW_SSAI")
    oView:Refresh("VIEW_CALCSSAI")

(__cAlias)->(DBSkip())
EndDo

If Select(__cAlias) > 0                                 
    (__cAlias)->(dbCloseArea())
EndIf

cQry := " SELECT FT_FILIAL, FT_CFOP, SUM(FT_VALCONT) AS FT_VALCONT, SUM(FT_BASEICM) AS FT_BASEICM, SUM(FT_VALICM) AS FT_VALICM "
cQry += " FROM "+ RetSqlName("SFT") + " SFT "
cQry += " INNER JOIN "+ RetSqlName("SZX") +" SZX ON SZX.ZX_PRODUTO = SFT.FT_PRODUTO "
cQry += " INNER JOIN "+ RetSqlName("SZY") +" SZY ON SZY.ZY_DECRETO = SZX.ZX_DECRET "
cQry += " WHERE SFT.D_E_L_E_T_ <> '*' " 
cQry += " AND   SZX.D_E_L_E_T_ <> '*' "
cQry += " AND   SZY.D_E_L_E_T_ <> '*' "
cQry += " AND   SZX.ZX_FILIAL  = '"+FwxFilial('SZX')+"' "
cQry += " AND   SZX.ZX_INDESP = 'N'  "
cQry += " AND   SZY.ZY_FILIAL  = '"+FwxFilial('SZY')+"' "
cQry += " AND   SZY.ZY_DATAINI <= '"+DToS(FWFldGet("ZG_DATAINI"))+"' "
cQry += " AND   SZY.ZY_DATAFIN >= '"+DToS(FWFldGet("ZG_DATAFIM"))+"' "
cQry += " AND   SZY.ZY_ATIVO  = 'T' "
cQry += " AND   SFT.FT_FILIAL  = '"+FwxFilial('SFT')+"' "
cQry += " AND   SFT.FT_EMISSAO BETWEEN "+ DToS(FWFldGet("ZG_DATAINI")) +" AND "+ DToS(FWFldGet("ZG_DATAFIM"))
cQry += " AND   SFT.FT_TIPOMOV = 'E' "
cQry += " AND   SFT.FT_DTCANC = '' "
cQry += " GROUP BY FT_FILIAL, FT_CFOP "
cQry := ChangeQuery(cQry)
IF Select(__cAlias) <> 0
    (__cAlias)->(DbCloseArea())
EndIf
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

nY := 0

While!(__cAlias)->(EOF())

    nY++
    If nY > oModelNENT:Length()
        oModelNENT:AddLine()
    EndIF

    oModelNENT:LoadValue("ZP_FILIAL"  , FwxFilial("SZP"))
    oModelNENT:LoadValue("ZP_COD"     , FWFldGet("ZG_COD"))
    oModelNENT:LoadValue("ZP_CFOP"    , (__cAlias)->FT_CFOP)
    oModelNENT:LoadValue("ZP_DESCFOP" , FWGetSX5('13',(__cAlias)->FT_CFOP,'pt-br')[1][4])
    oModelNENT:LoadValue("ZP_VALCONT" , (__cAlias)->FT_VALCONT)
    oModelNENT:LoadValue("ZP_BASEICM" , (__cAlias)->FT_BASEICM)
    oModelNENT:LoadValue("ZP_VALICM"  , (__cAlias)->FT_VALICM)

    oView:Refresh("VIEW_NENT")
    oView:Refresh("VIEW_CALCNENT")

(__cAlias)->(DBSkip())
EndDo

If Select(__cAlias) > 0                                 
    (__cAlias)->(dbCloseArea())
EndIf

cQry := " SELECT FT_FILIAL, FT_CFOP, SUM(FT_VALCONT) AS FT_VALCONT, SUM(FT_BASEICM) AS FT_BASEICM, SUM(FT_VALICM) AS FT_VALICM "
cQry += " FROM "+ RetSqlName("SFT") + " SFT "
cQry += " INNER JOIN "+ RetSqlName("SZX") +" SZX ON SZX.ZX_PRODUTO = SFT.FT_PRODUTO "
cQry += " INNER JOIN "+ RetSqlName("SZY") +" SZY ON SZY.ZY_DECRETO = SZX.ZX_DECRET "
cQry += " WHERE SFT.D_E_L_E_T_ <> '*' " 
cQry += " AND   SZX.D_E_L_E_T_ <> '*' "
cQry += " AND   SZY.D_E_L_E_T_ <> '*' "
cQry += " AND   SZX.ZX_FILIAL  = '"+FwxFilial('SZX')+"' "
cQry += " AND   SZX.ZX_INDESP = 'N'  "
cQry += " AND   SZY.ZY_FILIAL  = '"+FwxFilial('SZY')+"' "
cQry += " AND   SZY.ZY_DATAINI <= '"+DToS(FWFldGet("ZG_DATAINI"))+"' "
cQry += " AND   SZY.ZY_DATAFIN >= '"+DToS(FWFldGet("ZG_DATAFIM"))+"' "
cQry += " AND   SZY.ZY_ATIVO  = 'T' "
cQry += " AND   SFT.FT_FILIAL  = '"+FwxFilial('SFT')+"' "
cQry += " AND   SFT.FT_EMISSAO BETWEEN "+ DToS(FWFldGet("ZG_DATAINI")) +" AND "+ DToS(FWFldGet("ZG_DATAFIM"))
cQry += " AND   SFT.FT_TIPOMOV = 'S' "
cQry += " AND   SFT.FT_DTCANC = '' "
cQry += " GROUP BY FT_FILIAL, FT_CFOP "
cQry := ChangeQuery(cQry)
IF Select(__cAlias) <> 0
    (__cAlias)->(DbCloseArea())
EndIf
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

nY := 0

While!(__cAlias)->(EOF())

    nY++
    If nY > oModelNSAI:Length()
        oModelNSAI:AddLine()
    EndIF

    oModelNSAI:LoadValue("ZQ_FILIAL"  , FwxFilial("SZQ"))
    oModelNSAI:LoadValue("ZQ_COD"     , FWFldGet("ZG_COD"))
    oModelNSAI:LoadValue("ZQ_CFOP"    , (__cAlias)->FT_CFOP)
    oModelNSAI:LoadValue("ZQ_DESCFOP" , FWGetSX5('13',(__cAlias)->FT_CFOP,'pt-br')[1][4])
    oModelNSAI:LoadValue("ZQ_VALCONT" , (__cAlias)->FT_VALCONT)
    oModelNSAI:LoadValue("ZQ_BASEICM" , (__cAlias)->FT_BASEICM)
    oModelNSAI:LoadValue("ZQ_VALICM"  , (__cAlias)->FT_VALICM)

    oView:Refresh("VIEW_NSAI")
    oView:Refresh("VIEW_CALCNSAI")

(__cAlias)->(DBSkip())
EndDo

If Select(__cAlias) > 0                                 
    (__cAlias)->(dbCloseArea())
EndIf

oModelSENT:GoLine(1)
oModelSSAI:GoLine(1)
oModelNENT:GoLine(1)
oModelNSAI:GoLine(1)
oView:Refresh()

Return
