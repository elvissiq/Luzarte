#Include "PROTHEUS.CH"
#Include "TopConn.CH"

/*/{protheusDoc.marcadores_ocultos} JOB
  Este P.E. esta localizado na funcao A650GravC1 (Grava Solicitacao de Compras).
  @parâmetro Nã há
  @author Totvs Nordeste - Elvis Siqueira

  @sample
  É chamado apos gravar os dados no arquivo SC1 (Solic. de Compras).
  Link - https://tdn.totvs.com/pages/releaseview.action?pageId=6089297
  
  Return Nil
  @historia
  02/10/2023 - Desenvolvimento da Rotina.
/*/
User Function MT650C1()

Local aArea   := GetArea()
Local aCabec  := {}
Local aItens  := {}
Local aLinha  := {}
Local cAlias  := "TEMP"+FWTimeStamp(1)
Local cBarras := If(isSRVunix(),"/","\")
Local cErro   := ""
Local cQry    := ""
Local nY

Private lMsHelpAuto := .T.
Private lMsErroAuto := .F.
Private lAutoErrNoFile := .T.

  cQry := " SELECT * "
  cQry += " FROM "+ RetSqlName("SC1") + " SC1 "
  cQry += " WHERE SC1.D_E_L_E_T_ <> '*' " 
  cQry += " AND   SC1.C1_OP = '"+SC2->C2_NUM+SC2->C2_ITEM+SC2->C2_SEQUEN+"' " 
  cQry += " AND   SC1.C1_CC = '' " 
  cQry += " ORDER BY C1_NUM, C1_OP "
  cQry := ChangeQuery(cQry)
  IF Select(cAlias) <> 0
      (cAlias)->(DbCloseArea())
  EndIf
  dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),cAlias,.T.,.T.)

  DBSelectArea("SC1")
  SC1->(DBSetOrder(1))
  SC1->(DBGoTop())

  While !(cAlias)->(EOF()) 
    
    If SC1->(MSSeek((cAlias)->C1_FILIAL+(cAlias)->C1_NUM))

      aCabec := {}

      aAdd(aCabec,{"C1_NUM"     ,SC1->(C1_NUM),Nil})
      aAdd(aCabec,{"C1_SOLICIT" ,SC1->(C1_SOLICIT),Nil})
      aAdd(aCabec,{"C1_EMISSAO" ,SC1->(C1_EMISSAO),Nil})

      aLinha := {}               
      aAdd(aLinha,{"C1_ITEM",SC1->(SC1->C1_ITEM),Nil})            
      aAdd(aLinha,{"C1_PRODUTO",SC1->(SC1->C1_PRODUTO),Nil})         
      aAdd(aLinha,{"C1_QUANT",SC1->(SC1->C1_QUANT),Nil})
      aAdd(aLinha,{"C1_OP",SC1->(SC1->C1_OP),Nil})
      aAdd(aLinha,{"C1_CODORCA", SC1->C1_CODORCA,Nil})
      aAdd(aLinha,{"C1_TIPOEMP", SC1->C1_TIPOEMP,Nil})
      aAdd(aLinha,{"C1_ESPEMP", SC1->C1_ESPEMP,Nil})
      aAdd(aLinha,{"C1_TIPO", SC1->C1_TIPO,Nil})
      aAdd(aLinha,{"C1_MOEDA", SC1->C1_MOEDA,Nil})
      aAdd(aLinha,{"C1_GERACTR", SC1->C1_GERACTR,Nil})
      aAdd(aLinha,{"C1_ACCPROC", SC1->C1_ACCPROC,Nil})
      aAdd(aLinha,{"C1_COMPRAC", SC1->C1_COMPRAC,Nil}) 
      aAdd(aLinha,{"C1_XSOL", SC1->C1_XSOL,Nil})
      aAdd(aLinha,{"C1_CC",SC2->C2_CC,Nil})
      aAdd(aItens,aLinha) 
        
      lMsErroAuto := .F.

      MSExecAuto({|x,y,z| MATA110(x,y,z)},aCabec,aItens,4) //Alteração

      If lMsErroAuto
        aLog := GetAutoGRLog()
          For nY := 1 To Len(aLog)
              If !Empty(cErro)
                  cErro += CRLF
              EndIf
              cErro += aLog[nY]
          Next nY
        
        MakeDir(cBarras + "logatualizacc")
        MakeDir(cBarras + "logatualizacc" + cBarras+DToS(Date()))
        MemoWrite(cBarras + "logatualizacc" + cBarras+DToS(Date()) + cBarras +"log_"+ SubStr(FWTimeStamp(3),1,10) + "_" + SC2->C2_NUM + ".txt", cErro)
          
      EndIF 
    
    EndIf  
    (cAlias)->(DBSkip())
  EndDo 

  IF Select(cAlias) <> 0
      (cAlias)->(DbCloseArea())
  EndIf

RestArea(aArea)
Return
