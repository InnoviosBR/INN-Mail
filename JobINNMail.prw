#include "ap5mail.ch"  
#INCLUDE "XMLXFUN.CH"    
#include "fileio.ch"     
#INCLUDE "RWMAKE.CH"  
#INCLUDE "TBICONN.CH"  
#include "topconn.ch"  
#INCLUDE "INNLIB.CH"

/*/{Protheus.doc} JobINNMail
Inconsistencia de estoque
@author Clemilson Pena
@since 01/02/2013
@version 1.0
/*/

User Function JobINNMail(aParam) 

	//|*******************************************************|\\
	//|********************** Variaveis **********************|\\
	//|*******************************************************|\\	
	Local cTime			:= Time()
	Private cThreadId 	:= Alltrim(Str(ThreadId(),8,0))
	Private oINNLog		:= nil
	Private aServMail	:= {}
	Private TENTATIVAS	:= "10"
	Default aParam := {"01","01"}
	//|*******************************************************|\\
	//|********************** Variaveis **********************|\\
	//|*******************************************************|\\





	//|*******************************************************|\\
	//|******************* Abre o ambiente *******************|\\
	//|*******************************************************|\\		
	conout("************************************************************")
	conout("* JobINNMail EM EXECUCAO ["+cThreadId+"]")
	
	RPCClearEnv()
	RPCSetType(3)
	RPCSetEnv(aParam[1],aParam[2],"","","EST",,{"ZP5","WF7"}) 
	
	if Select("SM0") == 0 .or. Select("WF7") == 0
		conOut("* JobINNMail FINALIZADO ["+cThreadId+"] (Ambiente não configurado) ")
		RPCClearEnv()
		Return nil
	endif
	
	if ! LockByName("JobINNMail", .F., .F.)  // Semáforo, caso job esteja em execução finaliza esta chamada
		conOut("* JobINNMail FINALIZADO ["+cThreadId+"] (Outra thread em execucao) ")
		RPCClearEnv()
		Return nil
	endif 
	
	conout("* JobINNMail Data............: " + dtoc(date()) )
	conout("* JobINNMail Hora............: " + cTime )
	conout("* JobINNMail Environment.....: " + GetEnvServer() )
	conout("* JobINNMail Thread ID ......: " + cThreadId )
	conout("* JobINNMail Usuario.........: " + alltrim(upper(UsrRetName(RetCodUsr()))) )
	//|*******************************************************|\\
	//|******************* Abre o ambiente *******************|\\
	//|*******************************************************|\\


	oINNLog := INNLOG():New()
	oINNLog:SetTpServ("JOB")
	oINNLog:SetServico("JobINNMail")
	oINNLog:SetStatus(200,.F.)

	//|*******************************************************|\\
	//|*********************** PROCESSO **********************|\\
	//|*******************************************************|\\
	fProc()
	//|*******************************************************|\\
	//|*********************** PROCESSO **********************|\\
	//|*******************************************************|\\

	oINNLog:SetFim()


	//|*******************************************************|\\
	//|******************* Fecha o ambiente ******************|\\
	//|*******************************************************|\\	
	UnLockByName("JobINNMail", .F. , .F.) 
	dbCommitAll()
	//dbUnLockAll()
	DbCloseAll()
	RPCClearEnv()
	conout("* JobINNMail DEMOROU ["+ElapTime(cTime,Time())+"]")
	conout("* JobINNMail FINALIZADO ["+cThreadId+"]")
	conout("************************************************************")
	//|*******************************************************|\\
	//|******************* Fecha o ambiente ******************|\\
	//|*******************************************************|\\	
	
Return(.T.)

Static Function fProc()

	Local cConta	:= padr("JOBINNMAIL",TamSx3("WF7_PASTA")[1])
	Local oServer	:= nil
	Local lEnvia
	//Local lContinua := .T.
	Local aAnexo 	:= {}
	Local nLimite	:= SuperGetMV('IN_MAILLTD',.F.,100)
	Local nEnviado	:= 0
	Local nCntReg   := 0
	Local cLog

	//Local DATA_ATUAL := dtos(Date())//AND ZP5_DATA <= %Exp:DATA_ATUAL%
	//Local HORA_ATUAL := Time() //AND ZP5_HORA <= %Exp:HORA_ATUAL%

	oINNLog:AddItemLog("Iniciando o processo.")

	if select("TMP") <> 0
		TMP->(dbCloseArea())
	endif 

	BeginSql alias 'TMP'
		SELECT R_E_C_N_O_ REC FROM %table:ZP5%
		WHERE ZP5_FILIAL = %xfilial:ZP5% 
		  AND ZP5_STAMAI = '0'
		  AND ZP5_LOOP <= %Exp:TENTATIVAS%
		  AND CONVERT(datetime,CONCAT(ZP5_DATA , ' ', ZP5_HORA) ,21) <= GETDATE()
		  AND %NotDel%
		ORDER BY ZP5_DATA ,ZP5_HORA
	EndSql

	oINNLog:AddMemo("Query",GetLastQuery()[2])
	
	
	if !( TMP->(eof()) )
			
		While !( TMP->(eof()) )

			nCntReg += 1
		
			if nEnviado > nLimite
				TMP->(dbSkip())
				Loop
			endif

		
			dbSelectArea("ZP5")	
			ZP5->(dbSetOrder(1)) 
			ZP5->(dbGoTo(TMP->REC))
			
			IF ZP5->ZP5_STAMAI == "1"
				oINNLog:AddItemLog("Erro Recno...: " + cValToChar(ZP5->(Recno()))+" registro ja enviado")
				TMP->(dbSkip())
				Loop
			ENDIF

			IF ZP5->ZP5_STAMAI == "4"
				oINNLog:AddItemLog("Erro Recno...: " + cValToChar(ZP5->(Recno()))+" registro omitido")
				TMP->(dbSkip())
				Loop
			ENDIF

			oServer := nil
			if ZP5->(FieldPos("ZP5_CONTA")) > 0
				IF !fConecta(ZP5->ZP5_CONTA,@oServer)
					oINNLog:AddItemLog("Erro Recno...: " + cValToChar(ZP5->(Recno()))+" a conta não esta conectada")
					TMP->(dbSkip())
					Loop
				ENDIF
			else
				if !fConecta(cConta,@oServer)
					oINNLog:AddItemLog("Erro Recno...: " + cValToChar(ZP5->(Recno()))+" a conta não esta conectada")
					TMP->(dbSkip())
					Loop
				endif
			endif
			
			aAnexo := StrTokArr(Alltrim(ZP5->ZP5_ANEXO),CRLF)

			if ZP5->(FieldPos("ZP5_EMAILM")) > 0
				lEnvia := fEnvia(ZP5->ZP5_EMAILM,"","",ZP5->ZP5_ASSUNT,ZP5->ZP5_TEXTO,@oServer,@aAnexo)
			else
				lEnvia := fEnvia(ZP5->ZP5_EMAIL ,"","",ZP5->ZP5_ASSUNT,ZP5->ZP5_TEXTO,@oServer,@aAnexo)
			endif

			if lEnvia
				RecLock('ZP5',.F.)
					//Teve sucesso no envio
					Replace ZP5->ZP5_STAMAI  with "1"
					Replace ZP5->ZP5_DTMAIL  with date()
					Replace ZP5->ZP5_HRMAIL  with TimeFull()
					Replace ZP5->ZP5_LOOP    with ZP5->ZP5_LOOP + 1
				MsUnLock('ZP5')
				cLog := "Email enviado com sucesso!"
				fLogZP5(cLog)
				oINNLog:AddItemLog("Enviado Recno...: " + cValToChar(ZP5->(Recno())) )
				nEnviado += 1
			else
				//Falou no envio
				RecLock('ZP5',.F.)
					Replace ZP5->ZP5_LOOP  with ZP5->ZP5_LOOP + 1
				MsUnLock('ZP5')
				cLog := "Falha no envio do email, tentativa: "+cValtoChar(ZP5->ZP5_LOOP)
				if ZP5->ZP5_LOOP >= val( TENTATIVAS )
					RecLock('ZP5',.F.)
						Replace ZP5->ZP5_STAMAI  with "3"
					MsUnLock('ZP5')
					cLog := "Registro marcado com erro"
				endif
				fLogZP5(cLog)
				oINNLog:AddItemLog("Erro Recno......: " + cValToChar(ZP5->(Recno())),,,,,"ERRO")
				Sleep(15000)
			endif
				
			TMP->(dbSkip())
		
		EndDo
		
		fDesconecta()
		
	else
	
		oINNLog:AddItemLog("Nenhum email para enviar")
		
	endif
	
	if select("TMP") <> 0
		TMP->(dbCloseArea())
	endif

	oINNLog:AddItemLog("Tinhamos "+cValToChar(nCntReg)+" emails para enviar, enviamos "+cValToChar(nEnviado)+" no com sucesso.")
				
return()

Static Function fConecta(cConta,oServer)

	Local nRet := 0
	
	nPosSer := aScan(aServMail,{|x| x[1] == cConta })

	if nPosSer <= 0

		//Não existe essa conta conectada!
		aadd(aServMail,{cConta,nil})
		nPosSer := Len(aServMail)

		oINNLog:AddItemLog("Conectando...")
		
		dbSelectArea('WF7')
		WF7->(dbSetOrder(1))
		if !( WF7->(dbSeek(xFilial('WF7')+aServMail[nPosSer][1])) )
			oINNLog:AddItemLog("Conta de e-mail não encontrada!",,,,,"ERRO")
			oINNLog:SetStatus(400)
			Return(.F.)
		EndIf

		oINNLog:AddItemLog("cSmtpServer: "+alltrim(WF7->WF7_SMTPSR))
		oINNLog:AddItemLog("cAccount: "+alltrim(WF7->WF7_AUTUSU))
		oINNLog:AddItemLog("cPassword: "+alltrim (WF7->WF7_AUTSEN))
		oINNLog:AddItemLog("nMailPort: "+cValtoChar(WF7->WF7_SMTPPR))
		oINNLog:AddItemLog("nSmtpPort: "+cValtoChar(WF7->WF7_SMTPPR))
		oINNLog:AddItemLog("nTimeout: "+cValtoChar(WF7->WF7_TEMPO))
		oINNLog:AddItemLog("Criptografia: "+alltrim(WF7->WF7_SMTPSE))
											
		//Cria a conexão com o server STMP ( Envio de e-mail )	
		aServMail[nPosSer][2] := TMailManager():New()	
		if alltrim(WF7->WF7_SMTPSE) == "TLS" 
			aServMail[nPosSer][2]:SetUseTLS(.T.)
			oINNLog:AddItemLog("Ativando TLS")
		endif  
		if alltrim(WF7->WF7_SMTPSE) == "SSL" 
			aServMail[nPosSer][2]:SetUseSSL(.T.)
			oINNLog:AddItemLog("Ativando SSL")
		endif
		
		//< cMailServer >, < cSmtpServer >, < cAccount >, < cPassword >, [ nMailPort ], [ nSmtpPort ] 
		aServMail[nPosSer][2]:Init(	alltrim(WF7->WF7_SMTPSR)/*< cMailServer >*/,;
									alltrim(WF7->WF7_SMTPSR)/*< cSmtpServer >*/,;
									alltrim(WF7->WF7_AUTUSU)/*< cAccount >*/,;
									alltrim (WF7->WF7_AUTSEN)/*< cPassword >*/,;
									WF7->WF7_SMTPPR/*[ nMailPort ]*/,;
									WF7->WF7_SMTPPR/*[ nSmtpPort ] */)  

		//seta um tempo de time out com servidor de 1min	
		nRet := aServMail[nPosSer][2]:SetSmtpTimeOut( WF7->WF7_TEMPO )
		If nRet != 0		
			oINNLog:AddItemLog("Falha ao setar o time out! Erro: (" + str(nRet,6) + ") " + aServMail[nPosSer][2]:GetErrorString(nRet) )
			oINNLog:SetStatus(400)
			Return(.F.)
		EndIf    

		//realiza a conexão SMTP	
		nRet := aServMail[nPosSer][2]:SmtpConnect()
		If nRet != 0		
			oINNLog:AddItemLog("Falha ao conectar! Erro: (" + str(nRet,6) + ") " + aServMail[nPosSer][2]:GetErrorString(nRet))
			oINNLog:SetStatus(400)
			Return(.F.)
		EndIf	
		
		//autentica com a instancia	
		nRet := aServMail[nPosSer][2]:SMTPAuth(alltrim(WF7->WF7_AUTUSU),alltrim(WF7->WF7_AUTSEN))
		If nRet != 0		
			oINNLog:AddItemLog("Falha ao autenticar! Erro: (" + str(nRet,6) + ") " + aServMail[nPosSer][2]:GetErrorString(nRet),,,,,"ERRO")
			oINNLog:SetStatus(400)
			Return(.F.)
		EndIf

		oServer := aServMail[nPosSer][2]

	else

		oServer := aServMail[nPosSer][2]

	endif

Return(.T.)

Static Function fEnvia(cRetTo,cRetCc,cRetBCc,cRetSubject,cRetBody,oServer,aRetAnexo)

	//Apos a conexão, cria o objeto da mensagem	
	Local nX
	Local nRet := 0
	Local cLog
	Local oMessage := TMailMessage():New()	
	
	//Limpa o objeto	
	oMessage:Clear()	

	//Popula com os dados de envio	
	oMessage:cXMailer	:= "INN mail"
	oMessage:cFrom 		:= "INN mail <"+alltrim(WF7->WF7_AUTUSU)+">"
	oMessage:cTo 		:= Alltrim(cRetTo)
	oMessage:cCc 		:= Alltrim(cRetCc)
	oMessage:cBcc 		:= Alltrim(cRetBCc)
	oMessage:cSubject 	:= Alltrim(cRetSubject)
	oMessage:cBody 		:= cRetBody

	//Anexa o arquivo
	for nX := 1 To Len(aRetAnexo)
		if File(aRetAnexo[nX])
			nRet := oMessage:AttachFile(aRetAnexo[nX])
			If nRet < 0
				cLog := "Não foi possivel anexar o arquivo: "+aRetAnexo[nX]
				fLogZP5(cLog)
				oINNLog:AddItemLog(cLog)
				Return(.F.)
			EndIf
		else
			cLog := "Arquivo não encontrado: "+aRetAnexo[nX]
			fLogZP5(cLog)
			oINNLog:AddItemLog(cLog)
			Return(.F.)
		endif
	next nX
	
	//Envia o e-mail	
	nRet := oMessage:Send( oServer )
	If nRet != 0
		cLog := "Erro ao enviar o e-mail! Erro: (" + str(nRet,6) + ") " + oServer:GetErrorString(nRet)
		fLogZP5(cLog)
		oINNLog:AddItemLog(cLog)
		Return(.F.)
	EndIf	
	
Return(.T.)

Static Function fDesconecta() 
	
	Local nRet := 0
	Local oServer
	Local nY

	for nY := 1 To Len(aServMail)

		oServer := aServMail[nY][2]
	
		oINNLog:AddItemLog("Desconectando")
			
		//Desconecta do servidor	
		nRet := oServer:SmtpDisconnect()
		If nRet != 0		
			oINNLog:AddItemLog("Erro ao disconectar do servidor SMTP! Erro: (" + str(nRet,6) + ") " + oServer:GetErrorString(nRet))
			Return(.F.)
		EndIf

	next
	
Return(.T.)

/*Static Function Scheddef()

	Local aParam := {	"P",;//Tipo R para relatorio P para processo   
						"ParamDef",;// Pergunte do relatorio, caso nao use passar ParamDef            
						"",;  // Alias            
						{},;   //Array de ordens   
						"Envia e-mails"}    

Return aParam*/

Static Function fLogZP5(cRetLog)

    Local xTemp := ""

	if ZP5->(FieldPos("ZP5_LOG"))>0

		xTemp := Alltrim(ZP5->ZP5_LOG)
		xTemp += iif(Empty(xTemp),"",CRLF)
		xTemp += dToc(date()) + " " + TimeFull() + " -> " + cRetLog

		RecLock("ZP5",.F.)
			Replace ZP5->ZP5_LOG with xTemp
		MsUnlock("ZP5")

	endif

Return
