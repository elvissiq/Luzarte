//Bibliotecas
#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
#Include "TBICONN.CH"
#Include "TopConn.ch"

Static cAlias  := "SZX"
Static cTitulo := "Conf. - Produto x Prodepe"

//----------------------------------------------------------------------
/*/{PROTHEUS.DOC} PFISF06
FUNÇÃO PFISF06 - Tela para preenchimento dos dados Prodepe no Produto
@OWNER VAZAO  
@VERSION PROTHEUS 12
@SINCE 30/08/2023
@Tratamento para calculo do PRODEPE
/*/
//----------------------------------------------------------------------

User Function PFISF06()
Local aArea   := GetArea()
Local oBrowse

oBrowse := FWMBrowse():New()
oBrowse:SetAlias(cAlias)
oBrowse:SetDescription(cTitulo)

oBrowse:AddLegend("U_PStatus('A')", 'BR_VERDE'    , 'Conf. Prodepe realizada',"1",.T.)
oBrowse:AddLegend("U_PStatus('P')", 'BR_AZUL'     , 'Conf. Prodepe realizada parciamente',"1",.T.)
oBrowse:AddLegend("U_PStatus('F')", 'BR_VERMELHO' , 'Conf. Prodepe não realizada',"1",.T.)

oBrowse:SetMenuDef("PFISF06")

oBrowse:Activate()

RestArea(aArea)
Return

/*---------------------------------------------------------------------*
 | Func:  zPFISF06                                                     |
 | Desc:  Executa FWExecView (Chamada dentro da rotina MATA010)        |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
User Function zPFISF06(pCodPro,pDescPro,pOpc)
Local aArea   := GetArea()
Local cFilAux := cFilAnt
Local aRetFil := {}
Local nY 

Private cCodProd := pCodPro
Private cDescPro := Alltrim(pDescPro)
Private nOption  := pOpc

aRetFil := FwListBranches()
  
  For nY := 1 To Len(aRetFil)
    
    cFilAnt := Alltrim(aRetFil[nY,2])

    If nOption == 3
        FWExecView("","PFISF06",nOption,,{|| .T.},,50,/*aButtons*/)
    Else
        DBSelectArea(cAlias)
        SZX->(DBSetOrder(1))
        If SZX->(MSSeek(FWxFilial(cAlias)+cCodProd))
            FWExecView("","PFISF06",nOption,,{|| .T.},,50,/*aButtons*/)
        Else
            FWExecView("","PFISF06",3,,{|| .T.},,50,/*aButtons*/)
        EndIf 
    EndIf
  
  Next nY

cFilAnt := cFilAux
 
RestArea(aArea)
Return

/*---------------------------------------------------------------------*
 | Func:  MenuDef                                                      |
 | Desc:  Criação do Menu MVC                                          |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function MenuDef()
Local aRotFISF6 := FWMVCMenu("PFISF06")
    
    ADD OPTION aRotFISF6 TITLE 'Cad. Decreto PRODEPE' ACTION 'U_PFISF07' OPERATION 7 ACCESS 0 //OPERATION 7

Return (aRotFISF6)

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ModelDef()
Local oModel
Local oStruct := FWFormStruct(1, cAlias)

    oModel := MPFormModel():New("PFISF06M", /*bPre*/,/*bPost*/,/*bCommit*/,/*bCancel*/)
    oModel:AddFields("SZXMASTER", /*cOwner*/, oStruct)
    oModel:SetPrimaryKey({})

    oModel:SetDescription(cTitulo +" ("+cFilAnt +" - "+Alltrim(FwFilialName())+")")

Return oModel
 
/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
Local oModel := FWLoadModel("PFISF06")
Local oStruct := FWFormStruct(2, cAlias)
Local oView

    oView := FWFormView():New()    
    oView:SetModel(oModel)
    oView:SetProgressBar(.T.)
    
    oView:AddField("VIEW_SZX", oStruct, "SZXMASTER")

    If Alltrim(FunName()) == "MATA010"
      oView:SetAfterViewActivate({|oView| ViewActv(oView)})
      oView:EnableTitleView("VIEW_SZX",cTitulo +" ("+cFilAnt +" - "+Alltrim(FwFilialName())+")")
    Else 
      oStruct:SetProperty('ZX_PRODUTO', MVC_VIEW_CANCHANGE, .T.)
    EndIF 

    oView:CreateHorizontalBox("TELA" , 100 )
    oView:SetOwnerView("VIEW_SZX", "TELA")

    //oView:SetCloseOnOk({||.T.})
     
Return oView

/*---------------------------------------------------------------------*
 | Func:  ViewActv                                                     |
 | Desc:  Realiza o LoadValue nos campos Código e Descrição            |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ViewActv(oView)
Local oModel := FWModelActive() 
Local oModelSZX := oModel:GetModel("SZXMASTER")

    oModelSZX:LoadValue("ZX_PRODUTO" , cCodProd)
    oModelSZX:LoadValue("ZX_DESCRIC" , cDescPro)

    oView:Refresh()

Return

/*---------------------------------------------------------------------*
 | Func:  PStatus                                                   |
 | Desc:  Pesquisa o status do Pedido Agenciamento                     |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/

User Function PStatus(pTipo)
Local aArea    := GetArea()
Local aCampoZX := FWSX3Util():GetAllFields( cAlias, .F. )
Local lRet     := .F.
Local nCont    := 0
Local nY       := 0

  For nY := 1 To Len(aCampoZX)
      If !Empty(SZX->&(aCampoZX[nY])) .And. !(aCampoZX[nY] $("ZX_FILIAL/ZX_PRODUTO/ZX_DESCRIC"))
        nCont++
      EndIf 
  Next nY

  Do Case 
    Case pTipo == "A" //Totalmente preenchido

      nCont := (nCont+3) 
      If nCont == Len(aCampoZX)
        lRet := .T.
      EndIf

    Case pTipo == "P" //Preenchido parcialmente
      
      nCont := (nCont+3) 
      If nCont > 3 .And. nCont < Len(aCampoZX)
        lRet := .T.
      EndIf 

    Case pTipo == "F" //Não preenchido
      
      nCont := (nCont+3) 
      If nCont <= 3 .And. nCont < Len(aCampoZX)
        lRet := .T.
      EndIf
  EndCase 

RestArea(aArea)
Return lRet
