#Include "PROTHEUS.CH"
#Include "TopConn.CH"

/*/{protheusDoc.marcadores_ocultos} COMXFUN
  Fun��o SCGRVINI
  @par�metro N� h�
  @author Totvs Nordeste - Elvis Siqueira

  @sample
 // SCGRVINI - Inicializa os campos da Solicita��o de Compras
              (RecLock antes de gravar Solicita��o de Compras).

  Return Nil
  @historia
  05/09/2023 - Desenvolvimento da Rotina.
/*/
User Function SCGRVINI()
Local aArea  := GetArea()
Local lMATA650 := FwIsInCallStack("MATA650")

  If lMATA650
    SC1->C1_CC := SC2->C2_CC
  EndIF 

RestArea(aArea)
Return
