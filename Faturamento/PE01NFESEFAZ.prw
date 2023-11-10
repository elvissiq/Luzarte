//Bibliotecas
#Include 'totvs.ch'

#Define ENTER Chr(13)+Chr(10)

/*/{Protheus.doc} PE01NFESEFAZ
Ponto de entrada localizado na fun��o XmlNfeSef do rdmake NFESEFAZ. 
Atrav�s deste ponto � poss�vel realizar manipula��es nos dados do produto, 
mensagens adicionais, destinat�rio, dados da nota, pedido de venda ou compra, antes da 
montagem do XML, no momento da transmiss�o da NFe.
@author TOTVS NORDESTE (Elvis Siqueira)
@since 05/10/2023
@version 1.0
    @return Nil
        PE01NFESEFAZ - Manipula��o em dados do produto ( [ aParam ] ) --> aRetorno
    @example
        Nome	 	 	Tipo	 	 	    Descri��o	 	 	                        	 
 	    aParam   	 	Array of Record	 	aProd     := PARAMIXB[1]
                                            cMensCli  := PARAMIXB[2]
                                            cMensFis  := PARAMIXB[3]
                                            aDest     := PARAMIXB[4]
                                            aNota     := PARAMIXB[5]
                                            aInfoItem := PARAMIXB[6]
                                            aDupl     := PARAMIXB[7]
                                            aTransp   := PARAMIXB[8]
                                            aEntrega  := PARAMIXB[9]
                                            aRetirada := PARAMIXB[10]
                                            aVeiculo  := PARAMIXB[11]
                                            aReboque  := PARAMIXB[12]
                                            aNfVincRur:= PARAMIXB[13]
                                            aEspVol   := PARAMIXB[14]
                                            aNfVinc   := PARAMIXB[15]
                                            aDetPag   := PARAMIXB[16]
                                            aObsCont  := PARAMIXB[17]
                                            aProcRef  := PARAMIXB[18]
    @obs https://tdn.totvs.com/pages/viewpage.action?pageId=274327446
/*/

User Function PE01NFESEFAZ()
Local aProd     := PARAMIXB[1]
Local cMensCli  := PARAMIXB[2]
Local cMensFis  := PARAMIXB[3]
Local aDest     := PARAMIXB[4] 
Local aNota     := PARAMIXB[5]
Local aInfoItem := PARAMIXB[6]
Local aDupl     := PARAMIXB[7]
Local aTransp   := PARAMIXB[8]
Local aEntrega  := PARAMIXB[9]
Local aRetirada := PARAMIXB[10]
Local aVeiculo  := PARAMIXB[11]
Local aReboque  := PARAMIXB[12]
Local aNfVincRur:= PARAMIXB[13]
Local aEspVol   := PARAMIXB[14]
Local aNfVinc   := PARAMIXB[15]
Local adetPag   := PARAMIXB[16]
Local aObsCont  := PARAMIXB[17]
Local aProcRef  := PARAMIXB[18]
Local aRetorno  := {}

Local aArea		:= GetArea()
Local aAreaSB1  := SB1->(GetArea())
Local aMsgImp   := {}
Local nPosChave := 0
Local _nI	    := 0

//@ Bloco respons�vel por alterar a descri��o do produto do campo B5_DESC para B5_ESPECIF. INICIO
SB5->(dbSelectArea("SB5"))
SB5->(dbSetOrder(1))
For _nI :=1  to Len(aProd)
	SB5->(dbSeek(FWxFilial("SB5")+aProd[_nI,2]))
	If !Empty(SB5->B5_CEME)
        aProd[_nI][4] := Alltrim(SB5->B5_CEME)
    EndIf

    //////////// Mensagem na Nota Fiscal ///////////////
    nPosChave := 0
    DBSelectArea("ZZ0")
    ZZ0->(DBSetOrder(1)) //ZZ0_FILIAL+ZZ0_TES+ZZ0_NCM+ZZ0_EST
    IF ZZ0->(MSSeeK(FWxFilial("ZZ0")+Pad(aProd[_nI,27],TamSX3('ZZ0_TES')[1])+Pad(aProd[_nI,5],TamSX3('ZZ0_NCM')[1])+Pad(aDest[9],TamSX3('ZZ0_EST')[1]))) 
        nPosChave := aScan(aMsgImp, {|x| x == FWxFilial("ZZ0")+Pad(aProd[_nI,27],TamSX3('ZZ0_TES')[1])+Pad(aProd[_nI,5],TamSX3('ZZ0_NCM')[1])+Pad(aDest[9],TamSX3('ZZ0_EST')[1]) })
        If Empty(nPosChave)
            cMensCli += IIF(Empty(cMensCli),Alltrim(ZZ0->ZZ0_MENSAG),ENTER+Alltrim(ZZ0->ZZ0_MENSAG))
            aAdd(aMsgImp,FWxFilial("ZZ0")+Pad(aProd[_nI,27],TamSX3('ZZ0_TES')[1])+Pad(aProd[_nI,5],TamSX3('ZZ0_NCM')[1])+Pad(aDest[9],TamSX3('ZZ0_EST')[1]))
        EndIF 
    ElseIf ZZ0->(MSSeeK(FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+Pad(aProd[_nI,5],TamSX3('ZZ0_NCM')[1])+Pad(aDest[9],TamSX3('ZZ0_EST')[1]))) 
        nPosChave := aScan(aMsgImp, {|x| x == FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+Pad(aProd[_nI,5],TamSX3('ZZ0_NCM')[1])+Pad(aDest[9],TamSX3('ZZ0_EST')[1]) })
        If Empty(nPosChave)
            cMensCli += IIF(Empty(cMensCli),Alltrim(ZZ0->ZZ0_MENSAG),ENTER+Alltrim(ZZ0->ZZ0_MENSAG))
            aAdd(aMsgImp,FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+Pad(aProd[_nI,5],TamSX3('ZZ0_NCM')[1])+Pad(aDest[9],TamSX3('ZZ0_EST')[1]))
        EndIF 
    ElseIF ZZ0->(MSSeeK(FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+aProd[_nI,5]+Pad("*",TamSX3('ZZ0_EST')[1])))
        nPosChave := aScan(aMsgImp, {|x| x == FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+aProd[_nI,5]+Pad("*",TamSX3('ZZ0_EST')[1]) })
        If Empty(nPosChave)
            cMensCli += IIF(Empty(cMensCli),Alltrim(ZZ0->ZZ0_MENSAG),ENTER+Alltrim(ZZ0->ZZ0_MENSAG))
            aAdd(aMsgImp,FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+aProd[_nI,5]+Pad("*",TamSX3('ZZ0_EST')[1]))
        EndIF
    ElseIF ZZ0->(MSSeeK(FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+Pad("*",TamSX3('ZZ0_NCM')[1])+Pad("*",TamSX3('ZZ0_EST')[1])))
        nPosChave := aScan(aMsgImp, {|x| x == FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+Pad("*",TamSX3('ZZ0_NCM')[1])+Pad("*",TamSX3('ZZ0_EST')[1]) })
        If Empty(nPosChave)
            cMensCli += IIF(Empty(cMensCli),Alltrim(ZZ0->ZZ0_MENSAG),ENTER+Alltrim(ZZ0->ZZ0_MENSAG))
            aAdd(aMsgImp,FWxFilial("ZZ0")+Pad("*",TamSX3('ZZ0_TES')[1])+Pad("*",TamSX3('ZZ0_NCM')[1])+Pad("*",TamSX3('ZZ0_EST')[1]))
        EndIF
    EndIf 
    //////////// FIM da Mensagem na Nota Fiscal ///////////////

Next _nI
//@ Bloco respons�vel por alterar a descri��o do produto do campo B5_DESC para B5_ESPECIF. FIM

RestArea(aAreaSB1)
RestArea(aArea)

aadd(aRetorno,aProd)
aadd(aRetorno,cMensCli)
aadd(aRetorno,cMensFis)
aadd(aRetorno,aDest)
aadd(aRetorno,aNota)
aadd(aRetorno,aInfoItem)
aadd(aRetorno,aDupl)
aadd(aRetorno,aTransp)
aadd(aRetorno,aEntrega)
aadd(aRetorno,aRetirada)
aadd(aRetorno,aVeiculo)
aadd(aRetorno,aReboque)
aadd(aRetorno,aNfVincRur)
aadd(aRetorno,aEspVol)
aadd(aRetorno,aNfVinc)
aadd(aRetorno,AdetPag)
aadd(aRetorno,aObsCont)
aadd(aRetorno,aProcRef) 

Return aRetorno
