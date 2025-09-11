#include "protheus.ch"
#Include "tbiconn.ch"
#Include "topconn.ch"
#Include "APWEBEX.CH"
#INCLUDE "INNLIB.CH"

User Function wINNMail()

	oINNWeb:SetTitle("Consulta Mail")
	oINNWeb:SetIdPgn("wINNMail")

	cAcao := iif(Valtype(HttpGet->acao) == "C" .and. !empty(HttpGet->acao),HttpGet->acao,"")

	Do Case

		Case cAcao == "1"
			fDetalhe()

		Case cAcao == "2"
			HttpGet->chave := fTesta()
			HttpGet->inic  := dtoc(date())
			HttpGet->fim   := dtoc(date())
			fForca()
			fPesquisa()

		Case cAcao == "3"
			fForca()
			HttpGet->inic   := dtoc(date())
			HttpGet->fim    := dtoc(date())
			HttpGet->status := "0"
			fPesquisa()

		OtherWise
			fPesquisa()
									
	End Case
	
Return(.T.)// cHTML

Static Function fPesquisa()

    Local cQuery   := ""
    Local aOrigem     := {{"","Todos"}}
	//Local aStatus	:= {}
	//Local nI,nY
	//Local aFilSt	:= {}

	cid			:= iif(Valtype(HttpGet->id) == "C" .and. !empty(HttpGet->id),HttpGet->id,"")
	cChave		:= iif(Valtype(HttpGet->chave) == "C" .and. !empty(HttpGet->chave),HttpGet->chave,"")
	cOrigem		:= iif(Valtype(HttpGet->origem) == "C" .and. !empty(HttpGet->origem),HttpGet->origem,"")
	cStatus		:= iif(Valtype(HttpGet->status) == "C" .and. !empty(HttpGet->status),HttpGet->status,"")
	dInic		:= cTod(iif(Valtype(HttpGet->inic) == "C" .and. !empty(HttpGet->inic),HttpGet->inic,""))
	dFim 		:= cTod(iif(Valtype(HttpGet->fim) == "C" .and. !empty(HttpGet->fim),HttpGet->fim,""))
	cDest		:= iif(Valtype(HttpGet->dest) == "C" .and. !empty(HttpGet->dest),HttpGet->dest,"")
	cAssunto	:= iif(Valtype(HttpGet->assunto) == "C" .and. !empty(HttpGet->assunto),HttpGet->assunto,"")

	fListaOrigem(@aOrigem)

	//fListaStatus(@aStatus)

	/*aFilSt   := StrTokArr(cStatus,",")
	cStatus   := ""
	for nI := 1 to len(aFilSt)
		cStatus += iif(!Empty(cStatus),",","")
		cStatus += "'" +  Alltrim(aFilSt[nI]) + "'"
	next*/

	oINNWebParam := INNWebParam():New( oINNWeb )
	oINNWebParam:addHidden( {'acao','0'})

	oINNWebParam:addData( {'inic'	,'Data de'	,dInic	,.F.})
	oINNWebParam:addData( {'fim'	,'Data Ate'	,dFim	,.F.})

	oINNWebParam:addCombo( {'origem','Origem',cOrigem,aOrigem,.F.})
	oINNWebParam:addcBoxX3({'status','Status',cStatus,"ZP5_STAMAI",.F.})
	
	oINNWebParam:addText( {'chave'	,'Chave'		, 50,cChave		,.F.})
	oINNWebParam:addText( {'dest'	,'Destinatario'	,250,cDest		,.F.})
	oINNWebParam:addText( {'assunto','Assunto'		,250,cAssunto	,.F.})
	oINNWebParam:addText( {'id'		,'ID'			,  9,cid		,.F.})
	

	if 	!Empty(cOrigem) .or. ;
		!Empty(cChave) .or. ;
		!Empty(dInic) .or. ;
		!Empty(dFim) .or. ;
		!Empty(cDest) .or. ;
		!Empty(cAssunto) .or. ;
		!Empty(cStatus) .or. ;
		!Empty(cid) 

		oINNWebTable := INNWebTable():New( oINNWeb )
		oINNWebTable:AddHead({RetTitle("ZP5_ID"),"C","",.T.})
		oINNWebTable:AddHead({RetTitle("ZP5_STAMAI"),"C",""})
		oINNWebTable:AddHead({RetTitle("ZP5_DATA"),"C",""})
		oINNWebTable:AddHead({RetTitle("ZP5_DTMAIL"),"C",""})
		oINNWebTable:AddHead({RetTitle("ZP5_ORIGEM"),"C",""})
		oINNWebTable:AddHead({RetTitle("ZP5_EMAIL"),"C",""})
		oINNWebTable:AddHead({RetTitle("ZP5_ASSUNT"),"C",""})
		oINNWebTable:AddHead({RetTitle("ZP5_CHAVE"),"C",""})
		oINNWebTable:AddHead({RetTitle("ZP5_LOOP"),"N","",PesqPict("ZP5","ZP5_LOOP")})

		if !Empty(cid)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_ID LIKE '%"+Alltrim(cid)+"%' "
		endif

		if !Empty(dInic)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_DATA >= '"+dtos(dInic)+"' "
		endif

		if !Empty(dFim)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_DATA <= '"+dtos(dFim)+"' "
		endif


		if !Empty(cOrigem)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_ORIGEM = '"+Upper(Alltrim(cOrigem))+"' "
		endif

		if !Empty(cChave)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_CHAVE LIKE '%"+Upper(Alltrim(cChave))+"%' "
		endif

		if !Empty(cStatus)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_STAMAI = '"+cStatus+"' "
		endif

		if !Empty(cDest)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_EMAIL LIKE '%"+Alltrim(cDest)+"%' "
		endif

		if !Empty(cAssunto)
			cQuery += iif(!Empty(cQuery)," And ","")
			cQuery += " ZP5_ASSUNT LIKE '%"+Alltrim(cAssunto)+"%' "
		endif

		cQuery := "%"+cQuery+"%"

		BeginSql Alias "TMP"
			COLUMN ZP5_DATA AS DATE
			COLUMN ZP5_DTMAIL AS DATE
			SELECT 	ZP5_ID,
					ZP5_STAMAI,
					ZP5_DATA,
					ZP5_HORA,
					ZP5_DTMAIL,
					ZP5_HRMAIL,
					ZP5_ORIGEM,
					ZP5_EMAIL,
					ZP5_ASSUNT,
					ZP5_CHAVE,
					ZP5_LOOP,
				   R_E_C_N_O_ REC
			 FROM %Table:ZP5% ZP5
			WHERE ZP5.%NotDel% AND %Exp:cQuery%
			ORDER BY ZP5_DATA, ZP5_HORA
		EndSql

		oINNWeb:AddBody("<!-- "+GetLastQuery()[2]+" -->")
		
					
		WHILE (TMP->(!EOF()))
				
			oINNWebTable:AddCols({	TMP->ZP5_ID,;
									TMP->ZP5_STAMAI+fVOpcBox(TMP->ZP5_STAMAI,"","ZP5_STAMAI"),;
									dtoc(TMP->ZP5_DATA)+" - "+TMP->ZP5_HORA,;
									dtoc(TMP->ZP5_DTMAIL)+" - "+TMP->ZP5_HRMAIL,;
									TMP->ZP5_ORIGEM,;
									TMP->ZP5_EMAIL,;
									TMP->ZP5_ASSUNT,;
									TMP->ZP5_CHAVE,;
									TMP->ZP5_LOOP})
							
			oINNWebTable:SetLink(  , 1 , {"?x=wINNMail&acao=1&recno="+cValToChar(TMP->REC),"ZP5"+TMP->ZP5_ID} )
		
			TMP->(DbSkip())
	
		ENDDO
	
		TMP->(dbCloseArea())
			
	endif

	oINNWeb:AddHeadBtn( {"wINNMail","acao=2","Teste"} )
	oINNWeb:AddHeadBtn( {"wINNMail","acao=3","Forcar envio"} )

Return

Static Function fDetalhe()

	Local nRecno := Val(iif(Valtype(HttpGet->recno) == "C" .and. !empty(HttpGet->recno),HttpGet->recno,""))

	dbSelectArea("ZP5")
	ZP5->(dbGoTo(nRecno))

	oINNWeb:SetTitle("Consulta INN Mail - "+ZP5->ZP5_ID)

	oINNWebBrowse := INNWebBrowse():New( oINNWeb )
	oINNWebBrowse:SetTabela( "ZP5" )
	oINNWebBrowse:SetRec( nRecno )

Return

Static Function fListaOrigem(aOrigem)

	if select("TMP") <> 0
		TMP->(dbCloseArea())
	endif 

	BeginSql alias 'TMP'
		SELECT DISTINCT ZP5_ORIGEM 
		FROM %table:ZP5% 
		WHERE ZP5_ORIGEM != ' ' 
		  AND %NotDel%
		ORDER BY 1
	EndSql
		
	WHILE (TMP->(!EOF()))
	
		aadd(aOrigem,{Alltrim(TMP->ZP5_ORIGEM),Alltrim(TMP->ZP5_ORIGEM)})
		TMP->(DbSkip())

	ENDDO  

	if select("TMP") <> 0
		TMP->(dbCloseArea())
	endif 

Return

/*
Static Function fListaStatus(aStatus)

	if select("TMP") <> 0
		TMP->(dbCloseArea())
	endif 
	
	BeginSql alias 'TMP'
		SELECT DISTINCT ZP5_STAMAI 
		FROM %table:ZP5% 
		WHERE %NotDel%
		ORDER BY 1
	EndSql
		
	WHILE (TMP->(!EOF()))
	
		aadd(aStatus,{TMP->ZP5_STAMAI,TMP->ZP5_STAMAI+fVOpcBox(TMP->ZP5_STAMAI,"","ZP5_STAMAI")})

		TMP->(DbSkip())

	ENDDO  

	if select("TMP") <> 0
		TMP->(dbCloseArea())
	endif 

Return
*/

Static Function fTesta()

	Local oINNMail	:= ClsINNMail():New()
	Local cChave	:= fGetRam(10)
	Local cConta	:= padr("JOBINNMAIL",TamSx3("WF7_PASTA")[1])
	Local cDest		:= ""

	dbSelectArea('WF7')
	WF7->(dbSetOrder(1))
	if WF7->(dbSeek(xFilial('WF7')+cConta))
		cDest := alltrim(WF7->WF7_AUTUSU)
	else
		oINNWeb:AddAlert("Não foi possivel encontrar a configuração de workflow.","danger")
		Return("")
	EndIf 

	oINNMail:SetTitulo("Teste INN Mail")
	oINNMail:SetChave( cChave )
	oINNMail:AddBody("<h3>Teste INN Mail</h3>")
	oINNMail:AddNota("Nota adicional ao email")
	oINNMail:SetOrigem("wINNMail")
	oINNMail:ExeQuery( "Exemplo ExeQuery "," SELECT X3_CAMPO,X3_TITULO,X3_TIPO,X3_TAMANHO,X3_DECIMAL FROM "+RetSqlName("SX3")+" WHERE X3_ARQUIVO = 'ZP5' " )
	oINNMail:SetTable("Exemplo SetTable ",{ {"Campo 1", "C"} , {"Campo 2", "C"}},{{"Conteudo 1:1","Conteudo 1:2"},{"Conteudo 2:1","Conteudo 2:2"}})
	DbSelectArea("SA1")
	SA1->(dbGoTop())
	if !(SA1->(eof()))
		oINNMail:xCadastro("SA1",SA1->(Recno()))
	endif
	oINNMail:addTO("suporte@innovios.com.br")
	if oINNMail:Envia()
		oINNWeb:AddAlert("Email enviado com sucesso!","success")
	else
		oINNWeb:AddAlert("Não foi possivel enviar o email!","danger")
	endif

Return(cChave)

Static Function fForca()

	SmartJob(   "U_JobINNMail" ,;//Indica o nome do Job que será executado.
				getenvserver(),;//Indica o nome do ambiente em que o Job será executado.
				.F.,;//Indica se, verdadeiro (.T.), o processo será finalizado; caso contrário, falso (.F.).
				{SM0->M0_CODIGO,SM0->M0_CODFIL})//Parametros

	oINNWeb:AddAlert("Chamada na rotina U_JobINNMail em smartjob efetuada!","success")

Return
