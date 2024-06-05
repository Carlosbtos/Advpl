//Bibliotecas
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"                                                                                                                                               

/*
* Rotina		: GenToExcel - Minverva S.A.
* Autor			: Carlos Henrique de Oliveira 
* Data			: 15/09/2017
* Descricao		: Fun��o gen�rica para exporta��o de informa��es para o excel
* Modulo		: GEN
* Observa��es	: Cria estrutura TReport
*/      

User Function GenToExcel(cTitle,cSheet,cAlias,lWeb,lSoma,lSheet)    

	Default lSoma       := .T.        
	Default lWeb        := .F. 
	Default lSheet      := .F. 
	
	&&Fun��o a ser executada via StaticCall
	ExpToExcel(cTitle,cSheet,cAlias,lWeb,lSoma,lSheet)

Return  

/*/{Protheus.doc} ExpToExcel
Fun��o para exportar as informa��es para o excel no forma xml
@author Carlos Henrique de Oliveira
@since 15/09/2017
@version 1.0
	@return Nil, Fun��o n�o tem retorno
/*/
Static Function ExpToExcel(cTitulo,cWorkSht,cAlQry,lWeb,lSoma,lSheet)  
 
	//Local cAlias        := ""            
	Local oExcel 	 	:= Nil
	Local cWorkSheet	:= ""                             
	Local cTable   		:= ""
	Local cArquivo 		:= ""          
	Local cArqTrab      := ""  
	Local cLocArq       := ""
	Local cRootPath		:= ""       
	Local cMensRet      := ""
	Local cExtensao     := ".xls"
	Local aRows  		:= {}   
	//Local aProv         := {}
	Local aInf          := {}    
	Local nInf          := 1   
	//Local nLin    		:= 0    
	Local nColunas      := 0 
	Local cTime         := StrTran(Time(),":","")       
	Local cLocalSv 		:= "\webarq\"
	Local lRet          := .T.    
	Local lComando      := .F.
	Local nIt
	Default cTitulo     := "Relatorio Generico"
	Default cWorkSht    := "Planilha"   
	Default lSoma       := .T.        
	Default lWeb        := .F. 
	Default lSheet      := .F.       
	
	&&Cria o diret�rio onde ser�o salvos os arquivos caso o mesmo ainda n�o exista
	MakeDir(cLocalSv)
		
	If (cAlQry)->(Eof())       
	
		lRet:= .F.        
		
		cMensRet := "Nao ha dados a serem exportados. Reveja os parametros da pesquisa!"
	
		If lWeb
		
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] "+cMensRet) 
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("")          
			
		Else
	
			Aviso("Sem Dados",cMensRet,{"Ok"}, 2)  
			
		EndIf   
		
		Return {lRet,cArquivo,cMensRet}   
               
	Else
	
		cWorkSheet:= cWorkSht                             
		cTable    := cTitulo		
		oExcel 	  := FWMsExcelEx():New()     
		
		If !lWeb         
		
			cLib:= ""
			
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
			
		Else      
			
			cArquivo+= cLocalSv+cTitulo+cTime+cExtensao    
			cArqTrab+= cTitulo+cTime+cExtensao     
			MakeDir(cLocalSv)   
			
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] Criacao do arquivo. "+cArquivo) 
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("")     
			
		EndIf                                     
				
		&&Fun��o de montagem da estrutura das informa��es
		Struct((cAlQry), @aInf)       
		
		ProcRegua(Len(aInf))    
		
		If !lSheet
		
			&&Define uma estrutura sequencial de controle dentro do fluxo de execu��o para a query, permitindo a sua interrup��o na ocorr�ncia de um determinado evento
			Begin Sequence     
		
				oExcel:AddworkSheet(cWorkSheet)
				oExcel:AddTable(cWorkSheet,cTable)    
					
				For nInf := 1 To Len(aInf)      
				                                   
					&&Adiciona uma nova linha ao array aRows
					aAdd(aRows,{})
						
					For nIt := 1 To Len(aInf[nInf])       
								                                                      
						If ValType(aInf[nInf,nIt,2]) # "N"         
						
							&&Adiciona as informa��es a linha do array no formato texto 
							aAdd(aRows[nInf],aInf[nInf,nIt,2]) 
			                
			                &&Adiciona as colunas uma �nica vez
			     			If nInf == 1    
			     			
								&&Adiciona as colunas com formato texto
								oExcel:AddColumn(cWorkSheet,cTable,aInf[nInf,nIt,1],1,1)        
								
							EndIf
									
						Else
							
							&&Adiciona as informa��es a linha do array no formato num�rico
							aAdd(aRows[nInf],Round(aInf[nInf,nIt,2],2))  
			                 
							&&Adiciona as colunas uma �nica vez
			     			If nInf == 1   
			     			  				
								&&Adiciona as colunas com formato num�rico
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
			
			&&Caso ocorra erro na execu��o da query	
			Recover   
			
				lRet:= .F.        
				
				cArquivo:= "Nao_Gerado"        
				
				cMensRet:="Erro na exportacao dos dados, as sheets nao podem ser geradas por uma coluna numerica!"
				
				Return {lRet,cArquivo,cMensRet}		
			
			
			End Sequence    
			
		Else     
		
			&&Define uma estrutura sequencial de controle dentro do fluxo de execu��o para a query, permitindo a sua interrup��o na ocorr�ncia de um determinado evento
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
								
									&&Adiciona as informa��es a linha do array no formato texto 
									aAdd(aRows[nColunas],aInf[nInf,nIt,2]) 
					                
					                &&Adiciona as colunas uma �nica vez
					     			If nColunas == 1    
					     			
										&&Adiciona as colunas com formato texto
										oExcel:AddColumn(cWorkSheet,cTable,aInf[nInf,nIt,1],1,1)        
										
									EndIf
											
								Else
									
									&&Adiciona as informa��es a linha do array no formato num�rico
									aAdd(aRows[nColunas],Round(aInf[nInf,nIt,2],2))  
					                 
									&&Adiciona as colunas uma �nica vez
					     			If nColunas == 1   
					     			  				
										&&Adiciona as colunas com formato num�rico
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
				
			&&Caso ocorra erro na execu��o da query	
			Recover    
				           
				lRet:= .F.                     
				
				cArquivo:= "Nao_Gerado"
				
				cMensRet :="Erro na exportacao dos dados, verifique as informacoes da procedure!"
				
				Return {lRet,cArquivo,cMensRet}					
			
			End Sequence    		
			
		EndIf
		     
		If !lWeb 
		
			If Left(cLib,4) == "HTML"      
			
				cMensRet:= "Por estar utilizando a vers�o HTML o Download do arquivo "+cArqTrab+CRLF+"ser� iniciado automaticamente."     
					
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
		
				cMensRet:= "O arquivo "+cArqTrab+" foi gerado com sucesso no diret�rio "+cLocalSv 
			
				Aviso("Gera��o",cMensRet,{"Ok"}, 2)   
				
			EndIf  
				
		Else  

			MakeDir("C:\TOTVS 12\Microsiga\protheus_data\webarq")                
		
			cLocArq := '"C:\TOTVS 12\Microsiga\protheus_data\webarq\"'   
			
			cArqGer := '"C:\TOTVS 12\Microsiga\protheus_data\webarq\'+cArqTrab+'"'   
			  
			cMensRet:= "O arquivo "+cArqTrab+" foi gerado com sucesso no diretorio "+cLocalSv+" "   
			
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] "+cMensRet) 
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("") 
			
			cArquivo:= cArqTrab

			cMensRet:= "\\190.1.1.19\FTP\"+cArquivo	
			
			/*
			cArquivo:= Substr(cArqTrab,1,Len(cArqTrab)-Len(cExtensao)) 

			cArqZip := cArquivo+".zip"        
			
			cArquivo:= '"C:\TOTVS 12\Microsiga\protheus_data\webarq\'+cArquivo+'.zip"'         
			
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] Comando que sera executado. "+"7z.exe a -tzip "+cArquivo+" "+cArqGer) 
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("")                   
			
			lComando:= WaitRunSrv("7z.exe a -tzip "+cArquivo+" "+cArqGer, .T., "C:\FTP\")      
			
			If !lComando  

				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] Nao foi possivel criar o arquivo zip.")            
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("")           
				
			Else                          
				
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] Arquivo zip criado com sucesso.")            
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("") 			
				
			EndIf  	  

			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] Comando que sera executado. "+"xcopy "+cArquivo+" C:\FTP\") 
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("")   
             
			&&Efetua a copia do arquivo para a pasta C:\FTP
			lComando:= WaitRunSrv("xcopy "+cArquivo+" C:\FTP\", .T., "C:\FTP\")     
			
			If lComando          
			
				cMensRet += "e copiado para a pasta C:\FTP\."       
			
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] "+cMensRet) 
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("")           
				
				cMensRet:= "\\190.1.1.19\FTP\"+cArquivo //cArqZip
				
			Else       
				
				cMensRet += "porem nao foi possivel executar a copia do arquivo gerado para a pasta C:\FTP\. Processe as informacoes novamente."
   
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] "+cMensRet)            
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("")   
				
			EndIf     
			
			lComando:= WaitRunSrv("C:\BAT\del_arq.bat", .T., "C:\BAT\")  
			
			lComando:= WaitRunSrv("C:\BAT\del_arq_ftp.bat", .T., "C:\BAT\") 
			 
			/*
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] Comando que sera executado. "+"xcopy "+cLocArq+" C:\FTP\") 
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("")   
             
			&&Efetua a copia do arquivo para a pasta C:\FTP
			lComando:= WaitRunSrv("xcopy "+cLocArq+" C:\FTP\", .T., "C:\FTP\")     
			
			If lComando          
			
				cMensRet += "e copiado para a pasta C:\FTP\."       
			
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] "+cMensRet) 
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("")    
				
			Else       
				
				cMensRet += "porem, houve falha ao executar a copia do mesmo para a pasta C:\FTP\. Processe as informacoes novamente."
   
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] "+cMensRet)            
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("")   
				
			EndIf      
			
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] "+cArqTrab)            
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("")   	
			  
			cArquivo:= Substr(cArqTrab,1,Len(cArqTrab)-Len(cExtensao))  
			
			cArquivo:= cArquivo+".zip"
			
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("[WEBSERVICE] "+cArquivo)            
			U_zCONOUT(Replicate("-",50))
			U_zCONOUT("")  
			  
			lComando:= WaitRunSrv("7z.exe a -tzip C:\FTP\"+cArquivo+" C:\FTP\"+cArqTrab, .T., "C:\FTP\")  
			
			If !lComando  
			
				cArquivo:= cArqTrab  
			
				cMensRet:= "\\190.1.1.19\FTP\"+cArqTrab

				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] Nao foi possivel criar o arquivo zip.")            
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("")           
				
			Else                          
				
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("[WEBSERVICE] Arquivo zip criado com sucesso.")            
				U_zCONOUT(Replicate("-",50))
				U_zCONOUT("") 			
				
			EndIf  	
			
			cMensRet:= "\\190.1.1.19\FTP\"+cArquivo	
			*/		
			
		EndIf
		
	EndIf
	
Return {lRet,cArquivo,cMensRet}

/*/{Protheus.doc} Struct
Fun��o de montagem da estrutura das informa��es 
@author Carlos Henrique de Oliveira
@since 15/09/2017
@version 1.0
	@return Nil, Fun��o n�o tem retorno
/*/
Static Function Struct(cAlQry, aInf)  
                                                       
	Local aStruct := DbStruct() //DbStruct(cAlQry)   
	Local aProv  := {}
	Local nStru
	
	While !(cAlQry)->(EOF())     
	
		For nStru:= 1 To Len(aStruct)    
		
			If aStruct[nStru][2] == "C" .And. Left(aStruct[nStru][1],2) # "DT"
			
				&&Adiciona as informa��es no formato caracter
				aAdd(aProv,{aStruct[nStru][1],&(FieldName(nStru))})      
				
			EndIf
			
			If aStruct[nStru][2] == "N"
			
				&&Adiciona as informa��es no formato de numero
				aAdd(aProv,{aStruct[nStru][1],Round(&(FieldName(nStru)),2)})
			
			EndIf      
			
			If aStruct[nStru][2] == "D"
			
				&&Adiciona as informa��es no formato de data
				aAdd(aProv,{aStruct[nStru][1],DtoC(&(FieldName(nStru)))})
			
			EndIf
						
			If Left(aStruct[nStru][1],2) == "DT"
			
				&&Adiciona as informa��es no formato de data baseado na nomenclatura do campo
				aAdd(aProv,{aStruct[nStru][1],DtoC(StoD(&(FieldName(nStru))))})
			
			EndIf    
			
		Next nStru  
		
		aAdd(aInf, aProv)  
	
		aProv:= {}      
				
   		(cAlQry)->(dbSkip())
		
	EndDo	
	
Return	