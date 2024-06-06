//Bibliotecas
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"                                                                                                                                               

/*
* Rotina		: GenToExcel - Minverva S.A.
* Autor			: Carlos Henrique de Oliveira 
* Data			: 05/06/2024
* Descricao		: Função genérica para exportação de informações para o excel
* Modulo		: GEN
* Observações	: 
*/      

User Function GenToExcel(cTitle,cSheet,cAlias,lSoma,lSheet)   
 
	Default cTitulo     := "Relatorio Generico"
	Default cWorkSht    := "Planilha"   
	Default lSoma       := .T.         
	Default lSheet      := .F. 
	
	&&Função a ser executada via StaticCall
	ExpToExcel(cTitle,cSheet,cAlias,lSoma,lSheet)

Return  

/*/{Protheus.doc} ExpToExcel
Função para exportar as informações para o excel no forma xml
@author Carlos Henrique de Oliveira
@since 15/09/2017
@version 1.0
	@return Nil, Função não tem retorno
/*/
Static Function ExpToExcel(cTitulo,cWorkSht,cAlQry,lSoma,lSheet)  
          
Local oExcel		:= Nil
Local cWorkSheet	:= ""                             
Local cTable		:= ""
Local cArquivo		:= ""          
Local cArqTrab		:= ""  
Local cRootPath		:= ""       
Local cMensRet		:= ""
Local cExtensao		:= ".xls"
Local aRows		:= {}   
Local aInf		:= {}    
Local nInf		:= 1      
Local nColunas		:= 0 
Local cTime		:= StrTran(Time(),":","")       
Local lRet		:= .T.    
Local nIt      
	
If (cAlQry)->(Eof())       

	lRet:= .F.        
	
	cMensRet := "Nao ha dados a serem exportados. Reveja os parametros da pesquisa!"

	Aviso("Sem Dados",cMensRet,{"Ok"}, 2)  
	
	Return {lRet,cArquivo,cMensRet}   
			
Else

	cWorkSheet	:= cWorkSht                             
	cTable		:= cTitulo		
	oExcel		:= FWMsExcelEx():New()            

	cLib			:= ""
	
	GetRemoteType(@cLib)     
				
	If Left(cLib,4) == "HTML"           
	
		cArquivo += cLocalSv+AllTrim(FUNNAME())+cTime+cExtensao   
		cArqTrab += AllTrim(FUNNAME())+cTime+cExtensao  
			
	Else			                                                           
		
		cRootPath := GetSRVProfString("ROOTPATH","")  
		cArquivo  := cGetFile( ,"Selecione o local para salvar a planilha",, cRootPath, .F., ;
		nOR( GETF_LOCALHARD, GETF_LOCALFLOPPY, GETF_RETDIRECTORY ),.T., .T. )  
		
	
		//Verificando se foi informado o nome para o arquivo
		If Substr(cArquivo,Len(cArquivo),1) != "\"   
		
			cLocalSv := cArquivo
			cArquivo += cExtensao
			cArqTrab += cExtensao
		
		Else                          
			
			cLocalSv := cArquivo
			cArquivo += AllTrim(FUNNAME())+cTime+cExtensao  
			cArqTrab += AllTrim(FUNNAME())+cTime+cExtensao  
						
		EndIf	 
		
	EndIf                                    
			
	&&Função de montagem da estrutura das informações
	Struct((cAlQry), @aInf)       
	
	ProcRegua(Len(aInf))    
	
	If !lSheet
	
		&&Define uma estrutura sequencial de controle dentro do fluxo de execução para a query, permitindo a sua interrupção na ocorrência de um determinado evento
		Begin Sequence     
	
			oExcel:AddworkSheet(cWorkSheet)
			oExcel:AddTable(cWorkSheet,cTable)    
				
			For nInf := 1 To Len(aInf)      
												
				&&Adiciona uma nova linha ao array aRows
				aAdd(aRows,{})
					
				For nIt := 1 To Len(aInf[nInf])       
																					
					If ValType(aInf[nInf,nIt,2]) # "N"         
					
						&&Adiciona as informações a linha do array no formato texto 
						aAdd(aRows[nInf],aInf[nInf,nIt,2]) 
						
						&&Adiciona as colunas uma única vez
						If nInf == 1    
						
							&&Adiciona as colunas com formato texto
							oExcel:AddColumn(cWorkSheet,cTable,aInf[nInf,nIt,1],1,1)        
							
						EndIf
								
					Else
						
						&&Adiciona as informações a linha do array no formato numérico
						aAdd(aRows[nInf],Round(aInf[nInf,nIt,2],2))  
							
						&&Adiciona as colunas uma única vez
						If nInf == 1   
										
							&&Adiciona as colunas com formato numérico
							oExcel:AddColumn(cWorkSheet,cTable,aInf[nInf,nIt,1],3,2,lSoma) 
							
						EndIf     
						
					EndIf       
					
				Next nIt        
				
				&&Adiciona as Linhas 
				oExcel:AddRow(cWorkSheet,cTable,aRows[nInf])	                     
											
			Next nInf     
						
			aRows := {}      
			
			oExcel:Activate()              
			oExcel:GetXMLFile(cArquivo)  
			
			oExcelApp:= MsExcel():New()
			oExcelApp:SetVisible(.F.)    
		
		&&Caso ocorra erro na execução da query	
		Recover   
		
			lRet:= .F.        
			
			cArquivo:= "Nao_Gerado"        
			
			cMensRet:="Erro na exportacao dos dados, as sheets nao podem ser geradas por uma coluna numerica!"
			
			Return {lRet,cArquivo,cMensRet}		
		
		
		End Sequence    
		
	Else     
	
		&&Define uma estrutura sequencial de controle dentro do fluxo de execução para a query, permitindo a sua interrupção na ocorrência de um determinado evento
		Begin Sequence 
	
			While nInf <= Len(aInf)       
				
				cWorkSheet := aInf[nInf,1,2]        
				nColunas   := 0          
			
				If ValType(cWorkSheet) # "C"   
				
					Break
					
				EndIf				      
				
				oExcel:AddworkSheet(cWorkSheet)
				oExcel:AddTable(cWorkSheet,cTable)                      
												
				While cWorkSheet == aInf[nInf,1,2] 
				
					&&Adiciona uma nova linha ao array aRows
					aAdd(aRows,{})        
					
					nColunas:= nColunas+1      
					
					For nIt := 1 To Len(aInf[nInf])       
						
						If nIt > 1 
																						
							If ValType(aInf[nInf,nIt,2]) # "N"         
							
								&&Adiciona as informações a linha do array no formato texto 
								aAdd(aRows[nColunas],aInf[nInf,nIt,2]) 
								
								&&Adiciona as colunas uma única vez
								If nColunas == 1    
								
									&&Adiciona as colunas com formato texto
									oExcel:AddColumn(cWorkSheet,cTable,aInf[nInf,nIt,1],1,1)        
									
								EndIf
										
							Else
								
								&&Adiciona as informações a linha do array no formato numérico
								aAdd(aRows[nColunas],Round(aInf[nInf,nIt,2],2))  
									
								&&Adiciona as colunas uma única vez
								If nColunas == 1   
												
									&&Adiciona as colunas com formato numérico
									oExcel:AddColumn(cWorkSheet,cTable,aInf[nInf,nIt,1],2,3,lSoma) 
									
								EndIf     
								
							EndIf           
							
						EndIf   
						
					Next nIt           
					
					&&Adiciona as Linhas 
					oExcel:AddRow(cWorkSheet,cTable,aRows[nColunas])	                   
											
					nInf    := nInf+1                      
					
					If nInf > Len(aInf)
					
						Exit
						
					EndIf        
							
				EndDo        
				
				aRows := {}   
			
			EndDo      
			
			oExcel:Activate()              
			oExcel:GetXMLFile(cArquivo)  
			
			oExcelApp:= MsExcel():New()
			oExcelApp:SetVisible(.F.)   
			
		&&Caso ocorra erro na execução da query	
		Recover    
						
			lRet:= .F.                     
			
			cArquivo:= "Nao_Gerado"
			
			cMensRet :="Erro na exportacao dos dados, verifique as informacoes da query/procedure!"
			
			Return {lRet,cArquivo,cMensRet}					
		
		End Sequence    		
		
	EndIf
	
	If Left(cLib,4) == "HTML"      
	
		cMensRet:= "Por estar utilizando a versão HTML o Download do arquivo "+cArqTrab+CRLF+" será iniciado automaticamente."     
			
		Aviso("Download",cMensRet,{"Ok"}, 2)
	
		nRet:= CpyS2TW(cArquivo, .T.)    
		
		FErase(cArquivo)                   
		
		If nRet == 0           
		
			cMensRet:= "Arquivo "+cArqTrab+" enviado para Download com sucesso!"  
			
		Else              
		
			cMensRet:= "Falha no Download do arquivo "+cArqTrab+"."  
		
		EndIf 
		
		Aviso("Download",cMensRet,{"Ok"}, 2) 
		
	Else

		cMensRet:= "O arquivo "+cArqTrab+" foi gerado com sucesso no diretório "+cLocalSv 
	
		Aviso("Geração",cMensRet,{"Ok"}, 2)   
		
	EndIf  
	
EndIf

Return {lRet,cArquivo,cMensRet}

/*/{Protheus.doc} Struct
Função de montagem da estrutura das informações 
@author Carlos Henrique de Oliveira
@since 15/09/2017
@version 1.0
@return Nil, Função não tem retorno
/*/
Static Function Struct(cAlQry, aInf)  
													
Local aStruct := DbStruct() 
Local aProv   := {}
Local nStru

While !(cAlQry)->(EOF())     

	For nStru:= 1 To Len(aStruct)    
	
		If aStruct[nStru][2] == "C" .And. Left(aStruct[nStru][1],2) # "DT"
		
			&&Adiciona as informações no formato caracter
			aAdd(aProv,{aStruct[nStru][1],&(FieldName(nStru))})      
			
		EndIf
		
		If aStruct[nStru][2] == "N"
		
			&&Adiciona as informações no formato de numero
			aAdd(aProv,{aStruct[nStru][1],Round(&(FieldName(nStru)),2)})
		
		EndIf      
		
		If aStruct[nStru][2] == "D"
		
			&&Adiciona as informações no formato de data
			aAdd(aProv,{aStruct[nStru][1],DtoC(&(FieldName(nStru)))})
		
		EndIf
					
		If Left(aStruct[nStru][1],2) == "DT"
		
			&&Adiciona as informações no formato de data baseado na nomenclatura do campo
			aAdd(aProv,{aStruct[nStru][1],DtoC(StoD(&(FieldName(nStru))))})
		
		EndIf    
		
	Next nStru  
	
	aAdd(aInf, aProv)  

	aProv:= {}      
			
	(cAlQry)->(dbSkip())
	
EndDo	

Return	
