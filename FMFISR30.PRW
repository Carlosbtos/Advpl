//Bibliotecas
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"

/*
* Rotina		: FMFISR30 - Minverva S.A.
* Autor			: Carlos Henrique de Oliveira
* Data			: 29/04/2021
* Descricao		: Exporta��o Excel - Carga Funrural - Reinf R-2055
* Modulo		: FIS
* Observa��es		:
*/

User Function FMFISR30()

	Local cTitle   := "Exporta��o Excel - Carga FunRural - Reinf R-2055"
	Local cDescrip := "Exporta��o Excel para carga das informa��es referentes ao FunRural - Reinf R-2055"
	Local bProcess := {|oSelf|Process()}
	Local aButtons := {}
	Local cFunction:= "FMFISR30" //AllTrim(FUNNAME())
	Local cPerg    := "FMFISR30" //AllTrim(FUNNAME())

	Pergunte(cPerg,.t.)

	//���������������������������������������������������������Ŀ
	//�Utilizacao da funcao ProcLogIni para permitir a gravacao �
	//�do log no CV8 quando do uso da classe tNewProcess que    �
	//�grava o LOG no SXU (FNC 00000028259/2009)                �
	//�����������������������������������������������������������
	ProcLogIni(aButtons)

	tNewProcess():New( cFunction, cTitle, bProcess, cDescrip, cPerg )

Return

/*/{Protheus.doc} ExecSql
Fun��o de processamento do relat�rio
@author Carlos Henrique de Oliveira
@since 16/04/2019
@version 1.0
@return Nil, Fun��o n�o tem retorno
/*/
Static Function Process()

	FWMsgRun(, {|oSay| ExecSql() }, "Processando", "Processando rel�torio.")

Return

/*/{Protheus.doc} ExecSql
Fun��o de execucao de query ou procedure
@author Carlos Henrique de Oliveira
@since 16/04/2019
@version 1.0
@return Nil, Fun��o n�o tem retorno
/*/
Static Function ExecSql()

	Local cAlTrb    := GetNextAlias()
	Local cEmp      := MV_PAR01
	Local cFilIni   := MV_PAR02
	Local cFilFim   := MV_PAR03
	Local cDataIni  := DtoS(MV_PAR04)
	Local cDataFim  := DtoS(MV_PAR05)
	Local cTipo     := "2" //Carga Funrural - Reinf Evt-2055
	Local cTable   	:= "Exporta��o Excel - Carga FunRural - Reinf R-2055 Per�odo De "+DtoC(MV_PAR04)+" a "+DtoC(MV_PAR05)
	Local cWorkSheet:= "R-2055"
	Local cSql      := ""
	Local cQuebra   := "/"
	Local cSepara   := "\"
	Local lWeb      := .F.
	Local lSoma     := .F.
	Local lSheet    := .F.
	Local cProcs    := SuperGetMv("ZZ_SPRFUNR",.F.,"SPR_PROTHEUS_RELAT_FUNRURAL_SENAR/SPR_PROTHEUS_RELAT_FUNRURAL")
	Local cProc     := IIF(cQuebra$cProcs,Substr(cProcs,AT(cQuebra,cProcs)+1,AT(cSepara,cProcs)-AT(cQuebra,cProcs)-1),cProcs) // Retorna: Segunda Procedure  - SPR_PROTHEUS_RELAT_FUNRURAL

	If Select(cAlTrb) > 0
		
		(cAlTrb)->(DbCloseArea())
		
	EndIf

	cSql:= "EXEC "+cProc+" '"+cEmp+"','"+cFilIni+"','"+cFilFim+"','"+cDataIni+"','"+cDataFim+"','"+cTipo+"'"
	TcQuery cSql New Alias &cAlTrb

	IncProc("Exportando dados para o excel...")

	If (cAlTrb)->(!Eof())
		
		//StaticCall( GenToExcel, ExpToExcel, cTable, cWorkSheet, (cAlTrb) , lWeb, lSoma, lSheet)
		U_GenToExcel(cTable, cWorkSheet, (cAlTrb) , lWeb, lSoma, lSheet)
		
	Else
		
		Aviso("Sem Informa��es","N�o h� informa��es para os par�metros informados!",{"Ok"}, 2)
		
	EndIf

	If Select(cAlTrb) > 0
		
		(cAlTrb)->(DbCloseArea())
		
	EndIf

Return