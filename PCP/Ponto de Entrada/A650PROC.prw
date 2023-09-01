#Include "PROTHEUS.CH"
#Include "TopConn.CH"

/*/{protheusDoc.marcadores_ocultos} MATA650
  Função A650PROC
  @parâmetro Nã há
  @author Totvs Nordeste - Elvis Siqueira

  @sample
 // MT110ROT -  O ponto de entrada 'A650PROC' é executado após o 
                processamento da inclusão da(s) Op(s) e/ou solicitação de compra(s). 
                
                Dependendo do número de Op´s ou solicitações de compras que foram processadas 
                não é possível estar posicionado em tais registros, ou seja, se o cliente 
                necessitar posicionar em um Op ou solicitação de compras específica, 
                o mesmo deverá se encarregar disso.


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
