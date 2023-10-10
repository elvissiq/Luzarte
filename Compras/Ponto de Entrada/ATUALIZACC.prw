#Include "PROTHEUS.CH"
#Include "TopConn.CH"

/*/{protheusDoc.marcadores_ocultos} JOB
  Função ATUALIZACC
  @parâmetro Nã há
  @author Totvs Nordeste - Elvis Siqueira

  @sample
 //ATUALIZACC - Realiza Query nas SCs geradas pela rotina de Ordem de Produção
                (MsExecAuto para alterar a Solicitação de Compras com o Centro de Custo da OP).

  Return Nil
  @historia
  02/10/2023 - Desenvolvimento da Rotina.
/*/
User Function ATUALIZACC(aParam)

Local aArea     := GetArea()
Local aCabec    := {}
Local aItens    := {}
Local aLinha    := {}
Local aLog      := {}
Local cBarras   := If(isSRVunix(),"/","\")
Local cRootPath := AllTrim(GetSrvProfString("RootPath",cBarras))
Local __cAlias  := "TEMP"+FWTimeStamp(1)
Local cFile  := ""
Local cQry      := ""
Local cCCusto   := ""
Local cErro     := ""
Local nY

Private lMsHelpAuto := .T.
Private lMsErroAuto := .F.
Private lAutoErrNoFile := .T.
  
  If IsBlind()
    RpcClearEnv()
    RpcSetType(2) 
    RpcSetEnv(aParam[1],aParam[2],,,"COM")
  EndIf 

  cQry := " SELECT * "
  cQry += " FROM "+ RetSqlName("SC1") + " SC1 "
  cQry += " WHERE SC1.D_E_L_E_T_ <> '*' " 
  cQry += " AND   SC1.C1_OP <> '' " 
  cQry += " AND   SC1.C1_CC = '' " 
  cQry += " ORDER BY C1_NUM, C1_OP "
  cQry := ChangeQuery(cQry)
  IF Select(__cAlias) <> 0
      (__cAlias)->(DbCloseArea())
  EndIf
  dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry),__cAlias,.T.,.T.)

  DBSelectArea("SC2")
  SC2->(DBSetOrder(1))

  While!(__cAlias)->(EOF())

    SC2->(DBGoTop())

    If SC2->(MSSeek((__cAlias)->(C1_FILIAL)+(__cAlias)->(C1_OP)))

      cCCusto := SC2->C2_CC

      DBSelectArea("SC1")
      SC1->(DBSetOrder(1))
      SC1->(DBGoTop())
      If SC1->(MSSeek((__cAlias)->(C1_FILIAL)+(__cAlias)->(C1_NUM)+(__cAlias)->(C1_ITEM)))

        aAdd(aCabec,{"C1_NUM"     ,SC1->(C1_NUM),Nil})
        aAdd(aCabec,{"C1_SOLICIT" ,SC1->(C1_SOLICIT),Nil})
        aAdd(aCabec,{"C1_EMISSAO" ,SC1->(C1_EMISSAO),Nil})

        While !SC1->(EOF()) .And. SC1->C1_FILIAL == (__cAlias)->(C1_FILIAL) .And. SC1->C1_NUM == (__cAlias)->(C1_NUM) 
            aLinha := {}               
            aadd(aLinha,{"C1_ITEM",SC1->(SC1->C1_ITEM),Nil})            
            aadd(aLinha,{"C1_PRODUTO",SC1->(SC1->C1_PRODUTO),Nil})         
            aadd(aLinha,{"C1_QUANT",SC1->(SC1->C1_QUANT),Nil})
            aadd(aLinha,{"C1_OP",SC1->(SC1->C1_OP),Nil})
            aadd(aLinha,{"C1_CC",cCCusto,Nil})
            aadd(aLinha,{"C1_CODORCA", SC1->C1_CODORCA,Nil})
            aadd(aLinha,{"C1_TIPOEMP", SC1->C1_TIPOEMP,Nil})
            aadd(aLinha,{"C1_ESPEMP", SC1->C1_ESPEMP,Nil})
            aadd(aLinha,{"C1_TIPO", SC1->C1_TIPO,Nil})
            aadd(aLinha,{"C1_MOEDA", SC1->C1_MOEDA,Nil})
            aadd(aLinha,{"C1_GERACTR", SC1->C1_GERACTR,Nil})
            aadd(aLinha,{"C1_ACCPROC", SC1->C1_ACCPROC,Nil})
            aadd(aLinha,{"C1_COMPRAC", SC1->C1_COMPRAC,Nil}) 
            aadd(aLinha,{"C1_XSOL", SC1->C1_XSOL,Nil})

            aadd(aItens,aLinha) 
          SC1->(DBSkip())
        EndDo

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
            
            //Gera arquivo de log
            //cFile := cRootPath + cBarras + "logatualizacc" + cBarras + DToS(Date()) + cBarras +"log_" +FWTimeStamp(1)+ ".txt"
            cFile := cRootPath + cBarras + "logatualizacc" +  cBarras +"log_" +FWTimeStamp(1)+ ".txt"
            
            CONOUT("============================")
            CONOUT(File(cFile))
            CONOUT("============================")

            If File(cFile)
              nHandCtr := fCreate(cFile)
              fWrite(nHand,cErro)
              FClose(nHand)
            EndIf 
            //-------------------

            CONOUT(cErro)
        EndIF 

      EndIf 

    EndIF 

    (__cAlias)->(DBSkip())
  EndDo

IF Select(__cAlias) <> 0
  (__cAlias)->(DbCloseArea())
EndIf

If IsBlind()
  RPCClearEnv()
EndIf

RestArea(aArea)
Return
