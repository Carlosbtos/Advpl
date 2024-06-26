//Bibliotecas
#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'FWMVCDEF.CH'

/*/{Protheus.doc} GenToCad
Fun��o gen�rica de cadastro para Modelo 1 em MVC
@author Carlos Henrique de Oliveira
@since 08/09/2017
@version 1.0
	@return Nil, Fun��o n�o tem retorno
	@example
	u_GenToCad()
	@obs N�o se pode executar fun��o MVC dentro do f�rmulas
/*/
User Function GenToCad(cTab, aLeg, aOpc, cCmpStru, cFiltro, aSetKey)       

	//Cria o objeto de apresenta��o do browse
	Local oBrowse

	//Guarda a �rea atual
	Local aArea     := GetArea()

	//Guarda fun��o de chamada anterior
	Local cFunBkp   := FunName()

	//Declara a vari�vel para o t�tulo
	Local cTitulo   := ""

	//Declara a vari�vel para o FOR
	Local nLeg

	//Declara a vari�vel cTabela como private para manter seu valor durante toda execu��o da rotina
	Private cTabela := ""

	//Declara o vetor aLegenda como private para manter seu valor durante toda execu��o da rotina
	Private aLegenda:= {}

	//Declara a vari�vel cTabela como private para manter seu valor durante toda execu��o da rotina
	Private aOpcoes := {}

	//Declara a vari�vel cCmpRet como private para manter seu valor durante toda execu��o da rotina
	Private cCmpRet := ""

	//Define a vari�vel cTab com um valor Default
	Default cTab    := ""

	//Define o vetor aLeg com um valor Default
	Default aLeg    := {}

	//Define o vetor aLeg com um valor Default
	Default aOpc    := {}

	//Define o vetor aCmpStru com um valor Default
	Default cCmpStru:= {}

	//Define a vari�vel cFiltro com um valor Default
	Default cFiltro := ""

	//Atribui valor a vari�vel cTabela
	cTabela := cTab

	//Atribui valor ao vetor aLegenda
	aLegenda:= aLeg

	//Atribui valor ao vetor aOpcoes
	aOpcoes := aOpc

	//Atribui valor ao vetor aCmpRet
	cCmpRet := cCmpStru

	//Seta a fun��o atual
	SetFunName("GenToCad")

	//Cria a tabela utilizada no processo
	ChkFile(cTabela)

	//Define o t�tulo da tabela baseado no arquivo SX2
	cTitulo:= Capital(FwSX2Util():GetX2Name(cTabela, .T.))

	//Inst�nciando FWMBrowse - Somente com dicion�rio de dados
	oBrowse:= FWMBrowse():New()

	//Setando a tabela de cadastro de Autor/Interprete
	oBrowse:SetAlias(cTabela)

	//Setando a descri��o da rotina
	oBrowse:SetDescription(cTitulo)

	//Setando os filtros da rotina
	If !Empty(cFiltro)

		oBrowse:SetFilterDefault(cFiltro)

	EndIf

	//Legendas
	If Len(aLeg) > 0
		
		For nLeg:= 1 To Len(aLeg)
			
			//Vetor aLeg: aAdd(aLeg{Regra_Legenda,Cor_Legenda,Texto_Legenda})
			oBrowse:AddLegend( aLeg[nLeg,1], aLeg[nLeg,2], aLeg[nLeg,3] )
			
		Next nLeg
		
	EndIf

	//Ativa a Browse
	oBrowse:Activate()

	//Seta a fun��o de chamada anterior
	SetFunName(cFunBkp)

	RestArea(aArea)

Return Nil

/*---------------------------------------------------------------------*
| Func:  MenuDef                                                      |
| Autor: Carlos Henrique de Oliveira                                  |
| Data:  08/09/2017                                                   |
| Desc:  Cria��o do menu MVC                                          |
| Obs.:  /                                                            |
*---------------------------------------------------------------------*/
Static Function MenuDef()

	Local aRot := {}

	//Adicionando op��es padr�es
	If Len(aOpcoes) == 0
		
		ADD OPTION aRot TITLE 'Visualizar' ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
		ADD OPTION aRot TITLE 'Incluir'    ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION 3
		ADD OPTION aRot TITLE 'Alterar'    ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 4
		ADD OPTION aRot TITLE 'Excluir'    ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
		ADD OPTION aRot TITLE 'Copiar'     ACTION 'VIEWDEF.GenToCad' OPERATION 9                      ACCESS 0 //OPERATION 9
		
	EndIf

	//Adicionando op��es padr�es e/ou customizadas
	If Len(aOpcoes) > 0

		If Len(aOpcoes) == 1

			If aOpcoes[1][1] == 'OPER_ADC'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE 'Incluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION 3
				ADD OPTION aRot TITLE 'Alterar'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 4
				ADD OPTION aRot TITLE 'Excluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
				ADD OPTION aRot TITLE 'Copiar'      ACTION 'VIEWDEF.GenToCad' OPERATION 9                      ACCESS 0 //OPERATION 9
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION 6                      ACCESS 0 //OPERATION X

			EndIf

			If aOpcoes[1][1] == 'OPER_CUS'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION aOpcoes[1][4]          ACCESS 0 //OPERATION X

			EndIf

		EndIf

		If Len(aOpcoes) == 2

			If aOpcoes[1][1] == 'OPER_ADC'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE 'Incluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION 3
				ADD OPTION aRot TITLE 'Alterar'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 4
				ADD OPTION aRot TITLE 'Excluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
				ADD OPTION aRot TITLE 'Copiar'      ACTION 'VIEWDEF.GenToCad' OPERATION 9                      ACCESS 0 //OPERATION 9
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION aOpcoes[1][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[2][2] ACTION aOpcoes[2][3]      OPERATION aOpcoes[2][4]          ACCESS 0 //OPERATION X

			EndIf

			If aOpcoes[1][1] == 'OPER_CUS'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION aOpcoes[1][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[2][2] ACTION aOpcoes[2][3]      OPERATION aOpcoes[2][4]          ACCESS 0 //OPERATION X

			EndIf

		EndIf	

		If Len(aOpcoes) == 3

			If aOpcoes[1][1] == 'OPER_ADC'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE 'Incluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION 3
				ADD OPTION aRot TITLE 'Alterar'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 4
				ADD OPTION aRot TITLE 'Excluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
				ADD OPTION aRot TITLE 'Copiar'      ACTION 'VIEWDEF.GenToCad' OPERATION 9                      ACCESS 0 //OPERATION 9
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION aOpcoes[1][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[2][2] ACTION aOpcoes[2][3]      OPERATION aOpcoes[2][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[3][2] ACTION aOpcoes[3][3]      OPERATION aOpcoes[3][4]          ACCESS 0 //OPERATION X

			EndIf

			If aOpcoes[1][1] == 'OPER_CUS'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION aOpcoes[1][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[2][2] ACTION aOpcoes[2][3]      OPERATION aOpcoes[2][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[3][2] ACTION aOpcoes[3][3]      OPERATION aOpcoes[3][4]          ACCESS 0 //OPERATION X

			EndIf

		EndIf	

		If Len(aOpcoes) == 4

			If aOpcoes[1][1] == 'OPER_ADC'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE 'Incluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION 3
				ADD OPTION aRot TITLE 'Alterar'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 4
				ADD OPTION aRot TITLE 'Excluir'     ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
				ADD OPTION aRot TITLE 'Copiar'      ACTION 'VIEWDEF.GenToCad' OPERATION 9                      ACCESS 0 //OPERATION 9
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION aOpcoes[1][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[2][2] ACTION aOpcoes[2][3]      OPERATION aOpcoes[2][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[3][2] ACTION aOpcoes[3][3]      OPERATION aOpcoes[3][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[4][2] ACTION aOpcoes[4][3]      OPERATION aOpcoes[4][4]          ACCESS 0 //OPERATION X

			EndIf

			If aOpcoes[1][1] == 'OPER_CUS'

				ADD OPTION aRot TITLE 'Visualizar'  ACTION 'VIEWDEF.GenToCad' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 2
				ADD OPTION aRot TITLE aOpcoes[1][2] ACTION aOpcoes[1][3]      OPERATION aOpcoes[1][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[2][2] ACTION aOpcoes[2][3]      OPERATION aOpcoes[2][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[3][2] ACTION aOpcoes[3][3]      OPERATION aOpcoes[3][4]          ACCESS 0 //OPERATION X
				ADD OPTION aRot TITLE aOpcoes[4][2] ACTION aOpcoes[4][3]      OPERATION aOpcoes[4][4]          ACCESS 0 //OPERATION X

			EndIf

		EndIf	
		
	EndIf

Return aRot

/*---------------------------------------------------------------------*
| Func:  ModelDef                                                     |
| Autor: Carlos Henrique de Oliveira                                  |
| Data:  08/09/2017                                                   |
| Desc:  Cria��o do modelo de dados MVC                               |
| Obs.:  /                                                            |
*---------------------------------------------------------------------*/
Static Function ModelDef()

	//Cria o objeto do modelo de dados
	Local oModel := Nil

	//Cria a estrutura de dados utilizada na interface
	Local oStru  := FWFormStruct(1, cTabela)

	//Declara a vari�vel para o verifica��o de �ndice �nico
	Local lIndUnq:= .F.

	//Declara a vari�vel para o receber a chave prim�ria caso a tabela n�o tenha �ndice �nico
	Local aPrmKey:= {}

	//Declara a vari�vel para o t�tulo
	Local cTitulo:= ""

	//Declara o Array para verificar se existe �ndice �nico para a tabela
	Local aX2Unico := FwSX2Util():GetSX2data(cTabela, {"X2_UNICO"})

	//Verifica se existe �ndice �nico para o a tabela
	lIndUnq:= IIF(!Vazio(aX2Unico[1][2]),.T.,.F.)

	//Define o t�tulo da tela com base no SX2
	cTitulo:= Capital(FwSX2Util():GetX2Name(cTabela, .T.))

	//Instancia o modelo, n�o � recomendado colocar nome da user function (por causa do u_), respeitando 10 caracteres
	oModel := MPFormModel():New("Gen"+cTabela,/*bPre*/, /*bPos*/,/*bCommit*/,/*bCancel*/)

	//Atribui formul�rios para o modelo
	oModel:AddFields("FORM"+cTabela,/*cOwner*/,oStru)

	If !lIndUnq
		
		//Determina a PrimaryKey da entidade Modelo
		aPrmKey:= PrmKeyDef(cTabela)
		
		//Seta a chave prim�ria da rotina
		oModel:SetPrimaryKey(aPrmKey)
		
	EndIf

	//Adiciona a descri��o ao modelo
	oModel:SetDescription(cTitulo)

	//Seta a descri��o do formul�rio
	oModel:GetModel("FORM"+cTabela):SetDescription(cTitulo)

Return oModel

/*---------------------------------------------------------------------*
| Func:  ViewDef                                                      |
| Autor: Carlos Henrique de Oliveira                                  |
| Data:  08/09/2017                                                   |
| Desc:  Cria��o da vis�o MVC                                         |
| Obs.:  /                                                            |
*---------------------------------------------------------------------*/
Static Function ViewDef()

	//Cria o objeto do modelo de dados da Interface do Cadastro de Autor/Interprete
	Local oModel := FWLoadModel("GenToCad")

	//Cria a estrutura de dados utilizada na interface do cadastro de Autor
	Local oStru := FWFormStruct(2, cTabela,  { |cCampo| VldStruct(cCampo) } )

	//Cria o objeto oView como nulo
	Local oView := Nil

	//Declara��o de vari�vel para o t�tulo
	Local cTitulo := " "

	//Define o t�tulo da tela com base no SX2
	cTitulo:= Capital(FwSX2Util():GetX2Name(cTabela, .T.))

	//Cria a view que ser� o retorno da fun��o e setando o modelo da rotina
	oView := FWFormView():New()
	oView:SetModel(oModel)

	//Atribui formul�rios para interface
	oView:AddField("VIEW_"+cTabela, oStru, "FORM"+cTabela)

	//Cria um container com nome tela com 100%
	oView:CreateHorizontalBox("TELA",100)

	//Adiciona o t�tulo do formul�rio
	oView:EnableTitleView("VIEW_"+cTabela, cTitulo)

	//For�a o fechamento da janela na confirma��o
	oView:SetCloseOnOk({||.T.})

	//Aloca o formul�rio da interface dentro do container
	oView:SetOwnerView("VIEW_"+cTabela,"TELA")

Return oView

/*---------------------------------------------------------------------*
| Func:  PrmKeyDef                                                    |
| Autor: Carlos Henrique de Oliveira                                  |
| Data:  08/09/2017                                                   |
| Desc:  Fun��o para determinar a PrimaryKey da entidade Modelo       |
| Obs.:  /                                                            |
*---------------------------------------------------------------------*/
Static Function PrmKeyDef(cTab)

	//Declara o vetor para o receber a chave prim�ria da tabela
	Local aPrmKey := {}

	//Declara a vari�vel para o FOR
	Local nInd

	//Seta a vari�vel com os campo(s) para composi��o do �ndice da tabela
	aRet := FWSIXUtil():GetAliasIndexes(cTab)

	If Len(aRet) > 0

		For nInd:= 1 To Len(aRet[1])
			
			//Adiciona as informa��es em uma �nica linha do array
			aAdd(aPrmKey,aRet[1][nInd])
			
		Next nInd

	EndIf

Return aPrmKey

/*---------------------------------------------------------------------*
| Func:  VldStruct                                                    |
| Autor: Carlos Henrique de Oliveira                                  |
| Data:  08/09/2017                                                   |
| Desc:  Fun��o para determinar os campos da estrutura                |
| Obs.:  /                                                            |
*---------------------------------------------------------------------*/
Static Function VldStruct(cCampo)

Local lRet := .T.

	If !Vazio(cCmpRet)

		If !(AllTrim(cCampo) $ cCmpRet)

			lRet := .F.

		EndIf

	EndIf

Return lRet
