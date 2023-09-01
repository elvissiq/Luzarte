#Include "PROTHEUS.CH"
#Include "TopConn.CH"

/*/{protheusDoc.marcadores_ocultos} MATA650
  Função MT650C1
  @parâmetro Nã há
  @author Totvs Nordeste - Elvis Siqueira

  @sample
 // MT650C1 - Este P.E. esta localizado na função A650GravC1 
              (Grava Solicitação de Compras).

  Return Nil
  @historia
  31/08/2023 - Desenvolvimento da Rotina.
/*/
User Function MT650C1()
Local aArea  := GetArea()
Local aCabec := {}
Local aItens := {}
Local aLinha := {}
Local cNumSC := SC1->C1_NUM
Local cCCustoOP := SC2->C2_CC

Private lMsErroAuto := .F.
Private lMsHelpAuto	:= .T. 

  DBSelectArea("SC1")
  SC1->(DBSetOrder(1))
  IF SC1->(MsSeek(FWxFilial("SC1")+cNumSC))

    aCabec :={	{ "C1_FILIAL"	  , SC1->C1_FILIAL	, NIL},;
					   		{ "C1_SOLICIT"	, SC1->C1_SOLICIT	, NIL},;
					   		{ "C1_EMISSAO"	, SC1->C1_EMISSAO	, NIL}}

    While !SC1->(EOF()) .AND. SC1->C1_NUM == cNumSC
      
      aLinha := {}
      aLinha := { { "C1_ITEM"	   , SC1->C1_ITEM	   , NIL},;
					   		  { "C1_PRODUTO" , SC1->C1_PRODUTO , NIL},;
                  { "C1_QUANT"   , SC1->C1_QUANT   , NIL},;
                  { "C1_DATPRF"  , SC1->C1_DATPRF  , NIL},;
                  { "C1_LOCAL"   , SC1->C1_LOCAL   , NIL},;
					   		  { "C1_CC"	     , cCCustoOP	     , NIL}}
      
      aAdd(aItens, aLinha )
    SC1->(DBSkip())
    EndDo

    lMsErroAuto := .F.
    MSExecAuto({|X,Y,Z| Mata110(X,Y,Z)}, aCabec, aItens, 4)

    If lMsErroAuto
			MostraErro()
		EndIf 

  EndIF 
  
RestArea(aArea)
Return
