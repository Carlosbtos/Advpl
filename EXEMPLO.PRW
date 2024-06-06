//Bibliotecas
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"

/*
* Rotina		: EXEMPLO
* Autor			: Carlos Henrique de Oliveira
* Data			: 05/06/2024
* Descricao		: Exemplo de utilização da rotina GenToExcel
* Modulo		: FIS
* Observações		:
*/

User Function EXEMPLO()

	Local cTitle   := "Exemplo de Exportação para o Excel"
	Local cDescrip := "Rotina de exemplo de Exportação para o Excel"
	Local bProcess := {|oSelf|Process()}
	Local aButtons := {}
	Local cFunction:= "EXEMPLO" 
	Local cPerg    := "EXEMPLO"
		
	Pergunte(cPerg,.F.)
	
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Utilizacao da funcao ProcLogIni para permitir a gravacao ³
	//³do log no CV8 quando do uso da classe tNewProcess que    ³
	//³grava o LOG no SXU (FNC 00000028259/2009)                ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	ProcLogIni(aButtons)
	
	tNewProcess():New( cFunction, cTitle, bProcess, cDescrip, cPerg )
	
Return   

/*/{Protheus.doc} ExecSql
Função de processamento da rotina de exportação
@author Carlos Henrique de Oliveira
@since 05/06/2024
@version 1.0
	@return Nil, Função não tem retorno
/*/
Static Function Process()

	FWMsgRun(, {|oSay| ExecSql() }, "Processando", "Processando relátorio.")

Return
  
/*/{Protheus.doc} ExecSql
Função de execucao de query ou procedure 
@author Carlos Henrique de Oliveira
@since 05/06/2024
@version 1.0
	@return Nil, Função não tem retorno
/*/
Static Function ExecSql()     

	Local cAlTrb    := GetNextAlias()
	Local cSheet    := cValToChar(MV_PAR01) 	    
	Local cTable   	:= "Exemplo Exportação Excel - Opção "+cSheet   
	Local cWorkSheet:= "Planilha"       
	Local cSql      := ""             
	Local lSoma     := .T.   
	Local lSheet    := IIF(cSheet=="1",.T.,.F.)                        
	Local cProc     := SuperGetMv("ZZ_SPRFUNR",.F.,"SPR_EXEMPLO_LISTA_EMPRESAS")                                        
	
	If Select(cAlTrb) > 0 
	
		(cAlTrb)->(DbCloseArea())
	
	EndIf 

	cSql:= "EXEC "+cProc+"" //EXEC SPR_EXEMPLO_LISTA_EMPRESAS
	TcQuery cSql New Alias &cAlTrb        
   
	IncProc("Exportando dados para o excel...")     	
	
	If (cAlTrb)->(!Eof())
	
		U_GenToExcel(cTable, cWorkSheet, (cAlTrb) , lSoma, lSheet) 

	Else
	
		Aviso("Sem Informações","Não há informações para os parâmetros informados!",{"Ok"}, 2)  
		
	EndIf	
	
	If Select(cAlTrb) > 0   
	
		(cAlTrb)->(DbCloseArea())
	
	EndIf
	                         
Return	 
