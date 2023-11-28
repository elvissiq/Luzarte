#Include "Protheus.ch"
#Include "FwPrintSetup.Ch" 
#Include "RptDef.Ch" 

#DEFINE ENTER Chr(13)+Chr(10)
#define PAD_LEFT		0
#define PAD_RIGHT		1
#define PAD_CENTER   	2

/*/{Protheus.doc} PFISRTR
Termo de concentimento
@type function
@version
@author TOTVS Nordeste
@since 02/08/2023
@return 
/*/
User Function PFISRTR()

Local nLin     := 50
Local nCol     := 13
Local cDirLogo := CurDir()+"logotermo_"+Alltrim(cEmpAnt)+".png"
Local cCGC     := Transform(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_CGC" } )[1][2]),"@R 99.999.999/9999-99")
Local cEndEnt  := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_ENDENT" } )[1][2]))
Local cBaiEnt  := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_BAIRENT" } )[1][2]))
Local cCepEnt  := Transform(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_CEPENT" } )[1][2]),"@R 99999-999")
Local cCidEnt  := Capital(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_CIDENT" } )[1][2]))
Local cEstEnt  := Upper(Alltrim(FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , { "M0_ESTENT" } )[1][2]))
Local cNomeFor := Upper(Alltrim(SZH->ZH_NOME))
Local cMesZH   := Upper(MesExtenso(Val(SubStr(SZH->ZH_PERIODO,1,2))))
Local cAnoZH   := SubStr(SZH->ZH_PERIODO,3)
Local aSemana  := {"Domingo","Segunda-feira","Terça-feira","Quarta-feira","Quinta-feira","Sexta-feira","Sábado"}
Local cDataHj  := aSemana[Dow(dDataBase)]+", "+cValToChar(Day(dDataBase))+" de "+MesExtenso(Month(dDataBase))+" de "+cValToChar(Year(dDataBase))

oFont12  := TFont():New( "Calibri",,12,,.F.,,,,,.F. )
oFont12B := TFont():New( "Calibri",,12,,.T.,,,,,.F. )

oPrint:=FWMSPrinter():New("termo"+Time()+".rel",IMP_PDF, .F., , .T.)
oPrint:SetResolution(72)
oPrint:SetPortrait()
oPrint:SetPaperSize(DMPAPER_A4)
oPrint:SetMargin(60,60,60,60) // nEsquerda, nSuperior, nDireita, nInferior
oPrint:StartPage()

oPrint:SayBitmap(nLin, nCol,cDirLogo,320,110)
nLin += 150
oPrint:Say (nLin, nCol, cCidEnt +"/"+ cEstEnt +", "+cDataHj, oFont12,,,,PAD_LEFT)
nLin += 30
oPrint:Say (nLin, nCol, "À Secretaria da Fazenda do Estado de Pernambuco", oFont12,,,,PAD_LEFT)
nLin += 15
oPrint:Say (nLin, nCol, "Att: DPC", oFont12,,,,PAD_LEFT)
nLin += 40
oPrint:Say (nLin, nCol, "A Empresa Pancristal Ltda, devidamente inscrita no CACEPE sob o nº. 013546040 e no CNPJ sob", oFont12,,,,PAD_LEFT)
nLin += 10
oPrint:Say (nLin, nCol, "o nº. "+cCGC+", estabelecida na "+cEndEnt+", "+cBaiEnt+", ", oFont12,,,,PAD_LEFT)
nLin += 10
oPrint:Say (nLin, nCol, cCidEnt +"/"+ cEstEnt+", CEP: "+cCepEnt+", vem pelo presente solicitar autorização para emissão de Nota Fiscal,", oFont12,,,,PAD_LEFT)
nLin += 10
oPrint:Say (nLin, nCol, "a ser emitida contra a empresa: "+cNomeFor+",", oFont12,,,,PAD_LEFT)
nLin += 10
oPrint:Say (nLin, nCol, "relativa ao Ressarcimento (PRODEPE) competente ao mês de: "+cMesZH+" "+cAnoZH, oFont12,,,,PAD_LEFT)
nLin += 40
oPrint:Say (nLin, nCol, "Decretos Concessivos Nºs: 27.793/2005; 33.663/2009; 41.042/2014; 43.200/2016; 46.081/2018;" , oFont12,,,,PAD_LEFT)
nLin += 10
oPrint:Say (nLin, nCol, "47.223/2019; 48.601/2020; 48.602/2020." , oFont12,,,,PAD_LEFT)
nLin += 40
oPrint:Say (nLin, nCol, "Nesses termos, aguarda deferimento." , oFont12,,,,PAD_LEFT)

oPrint:EndPage()
oPrint:cPathPDF := "C:\temp\"
oPrint:Preview()
FreeObj(oPrint)
oPrint := Nil

Return
