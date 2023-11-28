//Bibliotecas
#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
#Include "TBICONN.CH"
#Include "TopConn.ch"

Static cAlias  := "SZY"
Static cTitulo := "Cadastro de Decreto Prodepe"

//----------------------------------------------------------------------
/*/{PROTHEUS.DOC} PFISF07
FUNÇÃO PFISF07 - Tela para de Decreto Prodepe
@OWNER VAZAO  
@VERSION PROTHEUS 12
@SINCE 06/09/2023
@Tratamento para calculo do PRODEPE
/*/
//----------------------------------------------------------------------

User Function PFISF07()
Local aArea   := GetArea()
Local oBrowse

oBrowse := FWMBrowse():New()
oBrowse:SetAlias(cAlias)
oBrowse:SetDescription(cTitulo)

oBrowse:SetMenuDef("PFISF07")

oBrowse:Activate()

RestArea(aArea)
Return

/*---------------------------------------------------------------------*
 | Func:  MenuDef                                                      |
 | Desc:  Criação do Menu MVC                                          |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function MenuDef()
Local aRotFISF7 := FWMVCMenu("PFISF07")

Return (aRotFISF7)

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ModelDef()
Local oModel
Local oStruct := FWFormStruct(1, cAlias)

    oModel := MPFormModel():New("PFISF07M", /*bPre*/,/*bPost*/,/*bCommit*/,/*bCancel*/)
    oModel:AddFields(cAlias+"MASTER", /*cOwner*/, oStruct)
    oModel:SetPrimaryKey({})

Return oModel
 
/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
Local oModel := FWLoadModel("PFISF07")
Local oStruct := FWFormStruct(2, cAlias)
Local oView

    oView := FWFormView():New()    
    oView:SetModel(oModel)
    oView:SetProgressBar(.T.)
    
    oView:AddField("VIEW_"+cAlias, oStruct, cAlias+"MASTER")

    oView:CreateHorizontalBox("TELA" , 100 )
    oView:SetOwnerView("VIEW_"+cAlias, "TELA")

    oView:SetCloseOnOk({||.T.})
     
Return oView
