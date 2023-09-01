#Include "PROTHEUS.CH"
#Include "TopConn.CH"

/*/{protheusDoc.marcadores_ocultos} MATA650
  Fun��o A650PROC
  @par�metro N� h�
  @author Totvs Nordeste - Elvis Siqueira

  @sample
 // MT110ROT -  O ponto de entrada 'A650PROC' � executado ap�s o 
                processamento da inclus�o da(s) Op(s) e/ou solicita��o de compra(s). 
                
                Dependendo do n�mero de Op�s ou solicita��es de compras que foram processadas 
                n�o � poss�vel estar posicionado em tais registros, ou seja, se o cliente 
                necessitar posicionar em um Op ou solicita��o de compras espec�fica, 
                o mesmo dever� se encarregar disso.


  Return aRotina - rotina com a chamado do programa
  @historia
  31/08/2023 - Desenvolvimento da Rotina.
/*/
User Function A650PROC()
Local aArea  := GetArea()
Local cAlias := "TMP"+FWTimeStamp(1)
Local cCCustoOP, cQry

IF Select(cAlias) <> 0
  (cAlias)->(DbCloseArea())
EndIf

  cQry := " SELECT * "
  cQry += " FROM " + RetSqlName("SC1") + " SC1 "
  cQry += " WHERE SC1.D_E_L_E_T_ <> '*'"
  cQry += " And SC1.C1_FILIAL  = '" + FWxFilial("SC1") + "'"
  cQry += " And SC1.C1_OP <> '' "
  cQry += " And SC1.C1_CC = '' "
  cQry += " And SC1.C1_ORIGEM = 'MATA650'"
  cQry := ChangeQuery(cQry)
  dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),cAlias,.F.,.T.)

  While ! (cAlias)->(Eof())
    
    DBSelectArea("SC2")
    SC2->(DBSetOrder(1))
    If SC2->(MSSeek((cAlias)->C1_FILIAL+SubSTR((cAlias)->C1_OP,1,6)+SubSTR((cAlias)->C1_OP,7,2)+SubSTR((cAlias)->C1_OP,9)))

      cCCustoOP := SC2->C2_CC

      If !Empty(cCCustoOP)
        DBSelectArea("SC1")
        SC1->(DBSetOrder(1))
        If SC1->(MSSeek((cAlias)->C1_FILIAL+(cAlias)->C1_NUM))
              RecLock("SC1",.F.)
                SC1->C1_CC := cCCustoOP
              SC1->(MSUnLock())
        EndIf 
      EndIF 
      
    EndIF

  (cAlias)->(DBSkip())
  EndDo 

IF Select(cAlias) <> 0
  (cAlias)->(DbCloseArea())
EndIf

RestArea(aArea)  
Return 
