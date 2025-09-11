#include "protheus.ch"
#Include "tbiconn.ch"
#Include "topconn.ch"
#Include "APWEBEX.CH"
#INCLUDE "INNLIB.CH"

CLASS ClsINNMail

	data cVersao
	data cConta
	data cTitulo
	data cNota
	data cMAIL
	//data cMAILSub
	data cHead
	data cBody
	//data cRodape
	data cFoot
	data lVersao
	data lAmbiente
	data lHead
	data lFoot
	data aDestinos
	data aDescCp
	data cOrigem
	data cID
	data lEnvInd
	data aAnexo
	data cCompany
	data cChave
	data lPronto
	data cNomeEmpresa

	data DATA_ATUAL
	data HORA_ATUAL
		
	METHOD New() Constructor

	METHOD ExecHead()
	METHOD ExecFoot()
	METHOD ExecMonta()
	METHOD NomeEmpresa()

	METHOD SetTitulo()
	METHOD SetChave()
	METHOD AddBody()
	METHOD SetBody()
	METHOD GetBody()
	METHOD AddNota()
	METHOD SetNota()
	METHOD SetOrigem()
	METHOD SetTable()
	METHOD ExeQuery()
	METHOD xCadastro()
	METHOD xModelo3()
	METHOD xBrowse()
	METHOD HidVersao()
	METHOD HidAmbiente()
	METHOD HidHead()
	METHOD HidFoot()
	METHOD SetIndividual()
	METHOD RetID()
	METHOD Envia()
	METHOD addTO()
	METHOD addTOGrupo()
	METHOD SetTo()
	METHOD addAnexo()

	METHOD SetEnvFuturo()
	METHOD SetConta()
          
ENDCLASS

METHOD New() Class ClsINNMail
	
	::cVersao	:= "3.5"
	::cConta    := padr("JOBINNMAIL",TamSx3("WF7_PASTA")[1])
	::cTitulo	:= ""
	::cNota		:= "&nbsp;"
	::cMAIL		:= ""
	//::cMAILSub	:= ""
	::cHead		:= ""
	::cBody		:= ""
	//::cRodape	:= "&nbsp;"
	::cFoot		:= ""
	::lVersao	:= .T.
	::lAmbiente	:= .T.
	::lHead		:= .T.
	::lFoot		:= .T.
	::aDestinos	:= {}
	::aDescCp	:= {}
	::cOrigem	:= FunName()//ListPilha()//FunName()//"ClsINNMail"
	::lEnvInd	:= .F.
	::aAnexo	:= {}
	::cCompany	:= ""
	::cChave	:= ""
	::lPronto	:= .F.

	::DATA_ATUAL := Date()
	::HORA_ATUAL := TimeFull()

	::cNomeEmpresa := "INN Mail"

	::cID := GetSxeNum("ZP5","ZP5_ID")
	ConfirmSx8()
	
Return Self

METHOD AddBody(cBody) Class ClsINNMail
	::cBody += "<tr>" + CRLF
	::cBody += "  <td colspan='3' style='padding: 10px;'>"+cBody+"</td>" + CRLF
	::cBody += "</tr>" + CRLF
Return

METHOD SetEnvFuturo(dData,cTime) Class ClsINNMail

	::DATA_ATUAL := dData
	::HORA_ATUAL := cTime

Return

METHOD SetConta(cConta) Class ClsINNMail

	::cConta := cConta

Return


METHOD SetTitulo(cTitulo) Class ClsINNMail
	::cTitulo := Alltrim(cTitulo)
Return

METHOD SetChave(cChave) Class ClsINNMail
	::cChave := Alltrim(cChave)
Return

METHOD SetBody(cBody,lBorda) Class ClsINNMail
	Default lBorda := .T.
	if lBorda
		::cBody := "<tr>" + CRLF
		::cBody += "  <td colspan='3' style='padding: 10px;'>"+cBody+"</td>" + CRLF
		::cBody += "</tr>" + CRLF
	endif
Return

METHOD GetBody() Class ClsINNMail
Return( ::cBody )


METHOD AddNota(cNota) Class ClsINNMail
	::cNota += cNota	
Return

METHOD SetNota(cNota) Class ClsINNMail
	::cNota := cNota	
Return

METHOD SetOrigem(cOrigem) Class ClsINNMail
	::cOrigem := cOrigem	
Return

METHOD SetIndividual(lEnvInd) Class ClsINNMail
	::lEnvInd := lEnvInd	
Return

METHOD HidVersao() Class ClsINNMail
	::lVersao := .F.
Return

METHOD HidAmbiente() Class ClsINNMail
	::lAmbiente := .F.
Return

METHOD HidHead() Class ClsINNMail
	::lHead := .F.
Return

METHOD HidFoot() Class ClsINNMail
	::lFoot := .F.
Return

METHOD RetID() Class ClsINNMail
Return(::cID)

METHOD SetTable(cTitulo,aHead,aCols,lHead,lInclui,lSimples) Class ClsINNMail

	//Local nCount := 0
	Local nLinha := 0
	Local nCol	 := 0
	Local cCampo := ""
	Local cTipo  := ""
	Local cAlign := ""
	Local nY 	 := 0
	Local cTable := ""
	
	//Local nReg := Len(aCols)
	
	Default lHead    := .T.
	Default lInclui  := .T.
	Default lSimples := .F.
	//Default lStriped := .T.
	
	//iif(empty(cTitulo),cTitulo := "",nil)			

	if !lSimples
		if !Empty(cTitulo)
			cTable += "<tr>" + CRLF
			cTable += "  <td colspan='3' style='padding: 10px 10px 0px 10px ;'><b>"+cTitulo+"</b></td>" + CRLF
			cTable += "</tr>" + CRLF
		endif
		
		cTable += "<tr>" + CRLF
		cTable += "  <td colspan='3' style='padding: 10px;'>" + CRLF
	endif
	cTable += "    <table width='100%' border='0' cellspacing='0' cellpadding='5'>" + CRLF

	if lHead
		cTable += "    <tr>" + CRLF
		For nY := 1 To Len(aHead)
			cTable += "      <td style='border: 1px solid #ddd;border-bottom-width: 2px;'>"+aHead[nY][1]+"</td>" + CRLF
		next
		cTable += "    </tr>" + CRLF
	endif	

	For nLinha := 1 To Len(aCols)
		//cTable += "    <tr id='"+cValToChar(nLinha)+"' "+iif(lStriped,"style='background-color: "+iif(mod(nLinha,2)==0,"#FFF","#F9F9F9"),"")+"'>" + CRLF
		cTable += "    <tr id='"+cValToChar(nLinha)+"' style='background-color: "+iif(mod(nLinha,2)==0,"#FFF","#F9F9F9")+"'>" + CRLF
		For nCol := 1 To Len(aHead)
			cCampo := aCols[nLinha][nCol]
			cTipo  := aHead[nCol][2]
			cAlign := ""
			if cTipo == "D"
				if Empty(cCampo)
					cCampo := ""
				else
					cCampo := dToc(cCampo)
					cAlign := "center"
				endif
			elseif cTipo == "N"
				cCampo := Alltrim(Transform(cCampo,aHead[nCol][3]))
				cAlign := "right"
			elseif cTipo == "L"
				cCampo := iif(cCampo,"Verdadeiro","Falso")
				cAlign := "center"
			else
				cCampo := Alltrim(cCampo)
				cCampo := StrTran(cCampo,"<br>",CRLF)
				cCampo := StrTran(cCampo,CRLF,"<br>")
			endif			
			cTable += "      <td style='border: 1px solid #ddd;'"+iif(!empty(cAlign)," align='"+cAlign+"'","")+">"+cCampo+"</td>" + CRLF
		next nCol
		cTable += "    </tr>" + CRLF
	Next nLinha
	
	if Len(aCols) == 0
		cTable += "    <tr>" + CRLF
		cTable += "      <td style='border: 1px solid #ddd;line-height: 2;background-color: #F9F9F9' colspan='"+cValToChar(Len(aHead)+1)+"'>Nenhum registro encontrado</td>" + CRLF
		cTable += "    </tr>" + CRLF
	endif
			
	cTable += "    </table>" + CRLF

	if !lSimples
		cTable += "  </td>" + CRLF
		cTable += "</tr>" + CRLF
		
		cTable += "<tr><td colspan='3'>&nbsp;</td></tr>" + CRLF
		cTable += "<tr><td colspan='3'>&nbsp;</td></tr>" + CRLF
	endif
	
	if lInclui
		::cBody += cTable
	endif
			
Return(cTable)

METHOD ExeQuery(cTitulo,cQuery) Class ClsINNMail

	Local cAlias 	:= GetNextAlias()
	Local aStrut 	:= {}
	Local aHead 	:= {}
	Local aCols		:= {}
	Local aLinha	:= {}
	Local nY		:= 0
	
	if select(cAlias) <> 0
		(cAlias)->(dbCloseArea())
	endif 
				 
	dbUseArea(.T.,"TOPCONN",TCGenQry(,,cQuery),cAlias,.F.,.T.)

	DbSelectArea(cAlias)
	(cAlias)->(dbGoTop())
	
	aStrut	:= (cAlias)->(dbStruct()) 
	
	for nY := 1 To Len(aStrut)

		Aadd(aHead,{aStrut[nY,1],aStrut[nY,2],"",.T.})
	
	next
	
	dbSelectArea("SX3")
	SX3->(dbSetOrder(2))
	
	for nY := 1 To Len(aHead)
		if SX3->(MSSeek(aHead[nY,1]))
			aHead[nY,1] := SX3->X3_TITULO
			aHead[nY,2] := SX3->X3_TIPO
			aHead[nY,3] := SX3->X3_PICTURE
		else
			IF aHead[nY,2] == "N"
				aHead[nY,3] := "@E 99,999,999,999.999"
			ENDIF
		endif
	next
			
	WHILE ((cAlias)->(!EOF()))
	
		aLinha := {}	
	
		for nY := 1 To Len(aStrut)
			
			Do Case
				Case aStrut[nY,2] == "C" .and. aHead[nY,2] == "D"
					Aadd(aLinha,sTod((cAlias)->&(aStrut[nY,1])))
				Case aStrut[nY,2] == "C"
					Aadd(aLinha,(cAlias)->&(aStrut[nY,1]))
				Case aStrut[nY,2] == "D"
					Aadd(aLinha,(cAlias)->&(aStrut[nY,1]))
				Case aStrut[nY,2] == "N"
					Aadd(aLinha,(cAlias)->&(aStrut[nY,1]))
				Case aStrut[nY,2] == "L"
					Aadd(aLinha,iif((cAlias)->&(aStrut[nY,1]),"Verdadeiro","Falso"))
				OtherWise
					Aadd(aLinha,"")				
			End Case
		
		next
		
		Aadd(aCols,aLinha)

		(cAlias)->(DbSkip())

	ENDDO  
				
	if select(cAlias) <> 0
		(cAlias)->(dbCloseArea())
	endif 
				
	Self:SetTable(cTitulo,aHead,aCols)

Return

METHOD xCadastro(cTabela,nRec,cCampos,aCPOObr,lSX3,aHead,lCria) Class ClsINNMail

	Local aDicionario 	:= {{cTabela + "X","Outros",{}}}
	Local aCampo		:= {}
	Local cTitulo	 	:= "Consulta padrão"
	Local cErro			:= ""
	Local cPasta
	Local nPasta
	Local nY
	Local nX
	
	Private ALTERA   := .F.
	Private DELETA   := .F.
	Private INCLUI   := .F.
	Private VISUAL   := .T.
	
	Default cCampos := ""
	Default aCPOObr := ""
	Default lSX3    := .T.
	Default lCria   := .F. 

	if !lCria
	
		if !Empty(cTabela) .and. nRec > 0
		    
			if lSX3
			
				DbSelectArea("SX2")
				SX2->(DbSetOrder(1))
				if !( SX2->(DbSeek(cTabela)) )
					cErro += "Tabela não encontrada!"
				else
					cTitulo := Alltrim(SX2->X2_NOME)
				endif
				
			endif		
			
			dbSelectArea(cTabela)
			(cTabela)->(dbGoTo(nRec))  
			
			if nRec != (cTabela)->(Recno())
				cErro += "Registro não encontrado!"
			endif
			
		else
			
			cErro += "Informe corretamente os parametros!"
				
		endif
	endif
	
	if Empty(cErro)	
        
		if lSX3
		
			if lCria
				RegToMemory(cTabela,.T.,.T.)
			else
				RegToMemory(cTabela,.F.,.T.)			
			endif
			
			//OpenSxs(,,,,,"SXA","SXA",,.F.)
			
			
			DbSelectArea("SXA")
			SXA->(DbSetOrder(1))
							
			DbSelectArea("SX3")
			SX3->(DbSetOrder(1))
			SX3->(dbSeek(cTabela))
			
			WHILE !SX3->(EOF()) .and. ALLTRIM(SX3->X3_ARQUIVO) == cTabela 
		
				IF !( X3USO(SX3->X3_USADO)) .and. Empty(cCampos) .and. !(Alltrim(SX3->X3_CAMPO) $ aCPOObr)
					SX3->(dbSkip())
					Loop
				ENDIF
				
				if ("_FILIAL" $ Alltrim(SX3->X3_CAMPO) ) .and. Empty(cCampos)
					SX3->(dbSkip())
					Loop
				ENDIF
	
				if !Empty(cCampos) .and.  !(Alltrim(SX3->X3_CAMPO) $ cCampos)
					SX3->(dbSkip())
					Loop
				ENDIF
				
				cPasta := cTabela + iif(Empty(SX3->X3_FOLDER),"X",Alltrim(SX3->X3_FOLDER))
				nPasta := aScan(aDicionario,{|x|  Alltrim(x[1]) == cPasta })
				if nPasta < 1
					if SXA->(dbSeek(SX3->X3_ARQUIVO+SX3->X3_FOLDER))
						aadd(aDicionario,{cPasta,alltrim(SXA->XA_DESCRIC),{}})
					else
						aadd(aDicionario,{cPasta,"Outros",{}})
					endif				
					nPasta := Len(aDicionario)
				endif
				
				aCampo := {Alltrim(SX3->X3_CAMPO),Alltrim(SX3->X3_TITULO),nil,0,"T",SX3->X3_ORDEM}
				
				IF SX3->X3_TIPO ==  "D"
					aCampo[3] := dToc(M->&(SX3->X3_CAMPO))
				ELSEIF SX3->X3_TIPO ==  "N"
					aCampo[3] := Alltrim(TRANSFORM(M->&(SX3->X3_CAMPO),SX3->X3_PICTURE))
				ELSEIF SX3->X3_TIPO ==  "L"
					aCampo[3] := IIF(M->&(SX3->X3_CAMPO),"VERDADEIRO","FALSO")
				ELSEIF SX3->X3_TIPO ==  "C" .and. empty(SX3->X3_CBOX)
					aCampo[3] := Alltrim(M->&(SX3->X3_CAMPO))
				ELSEIF SX3->X3_TIPO ==  "C" .and. !empty(SX3->X3_CBOX)
					aCampo[3] := Alltrim(M->&(SX3->X3_CAMPO)) + fVOpcBox(M->&(SX3->X3_CAMPO),SX3->X3_CBOX,SX3->X3_CAMPO)
				ELSEIF SX3->X3_TIPO ==  "M"
					aCampo[3] := Alltrim(M->&(SX3->X3_CAMPO))
					aCampo[5] := "M"
					aCampo[6] := "ZZ"
				ELSE
					aCampo[3] := "DESPREPARADO PARA O TIPO: " + SX3->X3_TIPO
				ENDIF    
				
				aCampo[4] := cValToChar(Len(aCampo[3]))
				aCampo[4] := iif(Val(aCampo[4])<=100,aCampo[4],"100")
				aCampo[4] := iif(Val(aCampo[4])<=10,"10",aCampo[4])
				aadd(aDicionario[nPasta][3],aCampo)
							
				SX3->(dbSkip())
				
			ENDDO 
			
		else

			aDicionario := {}
			aadd(aDicionario,{cTabela + "X","Outros",{}})		
			nPasta := 1
			
			for nY := 1 To Len(aHead)
				aCampo := {aHead[nY][1],aHead[nY][2],nil,0,"T",strzero(len(aCampo),2)}

				IF aHead[nY][3] ==  "D"
					aCampo[3] := dToc((cTabela)->&(aHead[nY][1]))
				ELSEIF aHead[nY][3] ==  "N"
					aCampo[3] := Alltrim(TRANSFORM((cTabela)->&(aHead[nY][1]),"@E 99,999,999,999.999999"))
				ELSEIF aHead[nY][3] ==  "L"
					aCampo[3] := IIF((cTabela)->&(aHead[nY][1]),"VERDADEIRO","FALSO")
				ELSEIF aHead[nY][3] ==  "C"
					aCampo[3] := Alltrim((cTabela)->&(aHead[nY][1]))
				ELSEIF aHead[nY][3] ==  "M"
					aCampo[3] := LEFT(Alltrim((cTabela)->&(aHead[nY][1])),100)
					aCampo[5] := "M"
					aCampo[6] := "ZZ"
				ELSE
					aCampo[3] := "DESPREPARADO PARA O TIPO: " + aHead[nY][3]
				ENDIF    
				
				aCampo[4] := cValToChar(Len(aCampo[3]))
				aCampo[4] := iif(Val(aCampo[4])<=100,aCampo[4],"100")
				aCampo[4] := iif(Val(aCampo[4])<=10,"10",aCampo[4])
				aadd(aDicionario[nPasta][3],aCampo)
				
			next nY
		
		endif
		
		//VARINFO("dicionario",aDicionario)
		aHead := {{"","C",""},{"","C",""}}
		aCols := {}
		aDicionario := aSort(aDicionario,,, { |x, y| x[1] < y[1] })

		for nY := 1 To Len(aDicionario)
			for nX := 1 To Len(aDicionario[nY][3])
				aadd(aCols, { aDicionario[nY][3][nX][2] , aDicionario[nY][3][nX][3] })
			next nX
		next nY


		Self:SetTable(cTitulo,aHead,aCols,.F.)

		
	else
							
		Self:AddBody(cErro)  
		
	endif

Return

// --------------------------------------------------------------------------
METHOD xModelo3(cTabela1,nRec1,cTabela2,nIndex,bRegra) Class ClsINNMail

	Private ALTERA   := .F.
	Private DELETA   := .F.
	Private INCLUI   := .F.
	Private VISUAL   := .T.

	Self:xCadastro(cTabela1,nRec1)		
	Self:xBrowse(cTabela2,nIndex,bRegra)
				                        		
Return

// --------------------------------------------------------------------------
METHOD xBrowse(cTabela2,nIndex,bRegra,cCampos) Class ClsINNMail

	Local aCampo		:= {}
	Local nY
	
	Local aHead 	:= {}
	Local aCols		:= {}
	Local aArea		:= GetArea()
	Local aAreaEsp	:= (cTabela2)->(GetArea())
	
	Private ALTERA   := .F.
	Private DELETA   := .F.
	Private INCLUI   := .F.
	Private VISUAL   := .T.
	
	Default cCampos := ""
				
	aCampo := {}
	
	DbSelectArea("SX3")
	SX3->(DbSetOrder(1))
	SX3->(dbSeek(cTabela2))
	
	WHILE !SX3->(EOF()) .and. ALLTRIM(SX3->X3_ARQUIVO) == cTabela2 

		IF ( X3USO(SX3->X3_USADO) .or.  "_FILIAL" $ Alltrim(SX3->X3_CAMPO) )  .or. Alltrim(SX3->X3_CAMPO) $ cCampos
						
			aadd(aCampo ,{Alltrim(SX3->X3_CAMPO),Alltrim(SX3->X3_TITULO),SX3->X3_TIPO,SX3->X3_PICTURE,SX3->X3_CBOX})
			
		ENDIF
					
		SX3->(dbSkip())
		
	ENDDO 				

	for nY := 1 To len(aCampo)
	
		aadd(aHead,{aCampo[nY,2],aCampo[nY,3],aCampo[nY,4]})
	
	next nY
				
	(cTabela2)->(DBClearFilter())
	(cTabela2)->( DBSetFilter( { || &bRegra } , bRegra ) )	
	(cTabela2)->(dbSetOrder(nIndex)) 
	(cTabela2)->(dbGoTop()) 
		
	WHILE !( (cTabela2)->(EOF()) )// .and. {|| bRegra }
	
		RegToMemory(cTabela2,.F.,.T.)
		aLinha := {}
		
		for nY := 1 To len(aCampo)
		
			IF aCampo[nY,3] ==  "D" // Data
				aadd(aLinha, M->&(aCampo[nY,1]) )
				
			ELSEIF aCampo[nY,3] ==  "N" // Numerico
				aadd(aLinha, M->&(aCampo[nY,1]) )
				
			ELSEIF aCampo[nY,3] ==  "L" // Logico
				aadd(aLinha, M->&(aCampo[nY,1]) )
				
			ELSEIF aCampo[nY,3] ==  "C" .and. Empty(aCampo[nY,5]) // Caracter
				aadd(aLinha, Alltrim(M->&(aCampo[nY,1])) )
				
			ELSEIF aCampo[nY,3] ==  "C" .and. !Empty(aCampo[nY,5]) // Caracter
				aadd(aLinha, Alltrim(M->&(aCampo[nY,1])) + fVOpcBox(M->&(aCampo[nY,1]),"",aCampo[nY,1]) )
				
			ELSEIF aCampo[nY,3] ==  "M" // Memo
				aadd(aLinha, LEFT(Alltrim(M->&(aCampo[nY,1])),100) )
				
			ELSE
			
				aadd(aLinha, "DESPREPARADO PARA O TIPO: " + aCampo[nY,3] )
				
			ENDIF 
		
		next nY
		
		aadd(aCols,aLinha)
	
		(cTabela2)->(dbSkip())
		
	ENDDO 
	
	Self:SetTable(cTabela2 + " - " + Alltrim(POSICIONE("SX2",1,cTabela2,"X2_NOME")),aHead,aCols)
	
	(cTabela2)->(DBClearFilter())
	(cTabela2)->(RestArea(aAreaEsp))
	RestArea(aArea)
		
Return

// --------------------------------------------------------------------------
METHOD ExecHead() Class ClsINNMail
	
	::cHead += "<!DOCTYPE html>" + CRLF
	::cHead += "<html>" + CRLF
	::cHead += "<head>" + CRLF
  	::cHead += "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" + CRLF
  	::cHead += "<title>"+::cTitulo+"</title>" + CRLF  	
	::cHead += "</head>" + CRLF
	
Return

// --------------------------------------------------------------------------
METHOD ExecFoot() Class ClsINNMail

	::cFoot += "</body>" + CRLF
	::cFoot += "</html>"
	
Return

// --------------------------------------------------------------------------
METHOD ExecMonta() Class ClsINNMail
	
	::cMAIL := ""

	Self:NomeEmpresa()	
	Self:ExecHead()
	Self:ExecFoot()

	::cMAIL += ::cHead	
	::cMAIL += "<body style='margin:0px; font-family: Helvetica, Arial, sans-serif; font-size: 12px;' topmargin='0' leftmargin='0' rightmargin='0'>" + CRLF
	::cMAIL += "<table width='100%' border='0' cellspacing='0' cellpadding='0' style='border: 1px solid #fff;' >" + CRLF
	if ::lHead
		if Empty(::cNota)
			::cMAIL += "  <tr style='width: 100%; border-color: #e7e7e7;padding: 15px;'>" + CRLF
			::cMAIL += "    <td width='25%' align='left' style='line-height: 20px; text-align: left; color: #5e5e5e; font-size: 18px;'>"+::cNomeEmpresa+"</td>" + CRLF
			::cMAIL += "    <td colspan='2' width='75%' align='center' style='text-align: center; font-size: 18px;'>"+::cTitulo+"</td>" + CRLF
			::cMAIL += "  </tr>" + CRLF
		else
			::cMAIL += "  <tr style='width: 100%; border-color: #e7e7e7;padding: 15px;'>" + CRLF
			::cMAIL += "    <td width='25%' align='left' style='line-height: 20px; text-align: left; color: #5e5e5e; font-size: 18px;'>"+::cNomeEmpresa+"</td>" + CRLF
			::cMAIL += "    <td width='50%' align='center' style='text-align: center; font-size: 18px;'>"+::cTitulo+"</td>" + CRLF
			::cMAIL += "    <td width='25%' align='right' style='text-align: right;font-size: 12px;'>"+::cNota+"</td>" + CRLF
			::cMAIL += "  </tr>" + CRLF
		endif
	endif
	::cMAIL += ::cBody
	if ::lFoot
		::cMAIL += "  <tr>" + CRLF
		::cMAIL += "    <td colspan='3'><hr></td>" + CRLF
		::cMAIL += "  </tr>" + CRLF
		::cMAIL += "  <tr>" + CRLF
		::cMAIL += "    <td colspan='2'>" + CRLF
		::cMAIL += "      <strong>Notificação automatica enviada pelo Protheus</strong><br>" + CRLF
		::cMAIL += "      <strong>Não responda este e-mail</strong><br>" + CRLF
		::cMAIL += "    </td>" + CRLF
		::cMAIL += "    <td width='25%' align='right'>"+iif(::lVersao,"Versão: "+::cVersao+"<br>","")+iif(::lAmbiente,"Ambiente: "+upper(ALLTRIM(GetEnvServer())),"")+"</td>" + CRLF
		::cMAIL += "  </tr>" + CRLF
	endif
	::cMAIL += "</table>" + CRLF
	::cMAIL += ::cFoot
	
	::cMAIL := StrToHtml(::cMAIL)

	::lPronto := .T.
				                                           
Return

// --------------------------------------------------------------------------
METHOD Envia(lInfoAut) Class ClsINNMail

	Local aArea		:= GetArea()
	Local lRet 		:= .F.
	Local aDest 	:= {}
	Local aDestTemp := {}
	Local nY
	Local nEnviados := 0
	Local cAnexo 	:= ""
	
	Default lInfoAut := .F.
	
	if !::lPronto
		Self:ExecMonta()
	endif
	/*======================== DEBUG =======================*/
	::cMAIL += CRLF + CRLF + "<!-- ID: " + ::cId + "-->"

	::cMAIL += CRLF + CRLF + "<!-- LEN: " + cValToChar(len(::cMAIL)) + "-->"	
	/*======================== DEBUG =======================*/
	
	if !(::lEnvInd) //NAO IRA ENVIAR ENDIVIDUALMENTE
		aDestTemp := aClone(::aDestinos)
		aSort(aDestTemp, , , { | x,y | x[2] < y[2] } )
		aDest := {{"",""}}
		For nY  := 1 To Len(aDestTemp)
			aDest[1][1] += iif(Empty(aDest[1][1]),"",", ") + Alltrim(aDestTemp[nY][1])
			aDest[1][2] += iif(Empty(aDest[1][2]),"",", ") + Alltrim(aDestTemp[nY][2])
		Next nY
	else
		aDest := aClone(::aDestinos)
	endif

	For nY  := 1 To Len(::aAnexo)
		cAnexo += iif(Empty(cAnexo),"",CRLF)
		cAnexo += ::aAnexo[nY]
		::cMAIL += CRLF + "<!-- anexo: " + ::aAnexo[nY] + "-->"	
	Next nY
	
	DbSelectArea("ZP5")
	For nY  := 1 To Len(aDest)
		IF !Empty(aDest[nY,2]) .and. RecLock('ZP5',.T.)//NOTIFICACOES                  
			Replace ZP5_FILIAL with xFilial('ZP5') // Filial (C,  2,  0)* Campo não usado
			Replace ZP5_EMAIL  with aDest[nY,2] // Email (C,100,  0)
			if ZP5->(FieldPos("ZP5_EMAILM")) > 0
				Replace ZP5_EMAILM with aDest[nY,2]
			endif
			if ZP5->(FieldPos("ZP5_CONTA")) > 0
				Replace ZP5_CONTA with ::cConta
			endif
			Replace ZP5_ASSUNT with ::cTitulo // Assunto (C,100,  0)
			Replace ZP5_TEXTO  with ::cMAIL // Texto (M, 10,  0)
			Replace ZP5_ORIGEM with ::cOrigem // Origem (C, 10,  0)
			Replace ZP5_ID     with ::cID
			Replace ZP5_DATA   with ::DATA_ATUAL // Data (D,  8,  0)
			Replace ZP5_HORA   with ::HORA_ATUAL // Hora (C,  8,  0)
			Replace ZP5_ANEXO  with cAnexo 
			Replace ZP5_CHAVE  with ::cChave
			Replace ZP5_STAMAI with "0"
			Replace ZP5_LOOP   with 0
			Replace ZP5_MSFIL  with SM0->M0_CODFIL
			if ZP5->(FieldPos("ZP5_LOG")) > 0
				Replace ZP5_LOG with dtoc(::DATA_ATUAL)+" "+::HORA_ATUAL+" -> Email criado"+CRLF+dtoc(::DATA_ATUAL)+" "+::HORA_ATUAL+" -> Inserido na lista de envio"
			endif
			MsUnLock('ZP5')
			nEnviados += 1
		endif	
	Next nY
	
	if len(aDest) == nEnviados
		lRet := .T.
	endif
	
	if lInfoAut
		if lRet
			FWAlertInfo("Notificação criada!","ClsINNMail")
		else
			FWAlertError("Não foi possive enviar o e-mail para todos os destinatarios!","ClsINNMail")
		Endif
	endif

	RestArea(aArea)
	
Return(lRet)

// --------------------------------------------------------------------------
METHOD addTO(xRetTo) Class ClsINNMail

	Local cMai := ""
	//Local nPos := 0
	Local nY

	IF Empty(xRetTo)
		Return(.F.)
	Endif
	
	xRetTo := strtran(xRetTo,";",",")
	if Valtype(xRetTo) == "C" //.and. "," $ xRetTo
		xRetTo := StrTokArr(xRetTo,",")
	endif
	
	For nY  := 1 To Len(xRetTo)
			
		xRetTo[nY] := Alltrim(xRetTo[nY])
		PswOrder(1)//1 - ID do usuário/grupo
		
		if PswSeek( xRetTo[nY] , .T. )
			if !PSWRET()[01][17] //verificar se esta bloqueado
				cMai := UsrRetMail(xRetTo[nY])
				if !Empty(cMai)
					if aScan(::aDestinos,{|x| x[2] == cMai}) <= 0
						if IsEmail(cMai)
							aadd(::aDestinos,{xRetTo[nY],cMai})
						endif
					endif
				endif
			endif
		elseif "@" $ xRetTo[nY]	
			if aScan(::aDestinos,{|x| x[2] == xRetTo[nY]}) <= 0
				if IsEmail(xRetTo[nY])
					aadd(::aDestinos,{"",xRetTo[nY]})
				endif
			endif
		elseif left(xRetTo[nY],1) == "G"
			//descobre todos os usuarios do grupo e chama esse mesmo metodo passando os e-mails da galera do grupo
			Self:addTOGrupo(xRetTo[nY])
		endif

	next

	U_aUnique(@::aDestinos,2)
	
Return(.T.)

METHOD addTOGrupo(xRetGrupo) Class ClsINNMail

	Local cAliasUSR := GetNextAlias()
	Local cUrs := ""

	xRetGrupo := Alltrim(xRetGrupo)
	xRetGrupo := Right(xRetGrupo,6)

	BeginSql Alias cAliasUSR
		SELECT
			DISTINCT su.USR_ID ,
			su.USR_EMAIL
		FROM
			SYS_USR_GROUPS sug
		INNER JOIN SYS_USR su ON
			sug.USR_ID = su.USR_ID
		WHERE
			sug.USR_GRUPO =  %Exp:xRetGrupo%
			AND su.USR_MSBLQL != '1'
			AND sug.%NotDel%
			AND su.%NotDel%
		ORDER BY 1
	EndSql 
	
	While !( (cAliasUSR)->(eof()) )
		cUrs += iif(Empty(cUrs),"",",")
		cUrs += (cAliasUSR)->USR_ID
		(cAliasUSR)->(dbskip())
	EndDo
	
	(cAliasUSR)->(dbCloseArea())
	Self:addTO(cUrs)

Return()

// --------------------------------------------------------------------------
METHOD SetTo(xRetTo) Class ClsINNMail

	::aDestinos := {}
	Self:addTO(xRetTo)
	
Return()

METHOD addAnexo(cAnexo) Class ClsINNMail
	aadd(::aAnexo,cAnexo)
Return

METHOD NomeEmpresa() Class ClsINNMail

	Local cArquivo := "/system/innmail_nomeempresa.txt"
    Local cConteudo := ""
    Local oFile := FWFileReader():New(cArquivo)
    Local aLinhas
    Local cLinAtu
 
    //Se o arquivo pode ser aberto
    If (oFile:Open())
 
        //Se não for fim do arquivo
        If ! (oFile:EoF())
            //Definindo o tamanho da régua
            aLinhas := oFile:GetAllLines()
 
            //Método GoTop não funciona, deve fechar e abrir novamente o arquivo
            oFile:Close()
            oFile := FWFileReader():New(cArquivo)
            oFile:Open()
 
            //Enquanto houver linhas a serem lidas
            While (oFile:HasLine())
                //Buscando o texto da linha atual
                cLinAtu := oFile:GetLine()
                cConteudo += cLinAtu
            EndDo
        EndIf
 
        //Fecha o arquivo e finaliza o processamento
        oFile:Close()

    EndIf

	if !Empty(cConteudo)
		::cNomeEmpresa := cConteudo
	endif

Return

Static Function StrToHtml(cTexto)

	cTexto := StrTran(cTexto,"ç","&ccedil;")
	cTexto := StrTran(cTexto,"á","&aacute;")
	cTexto := StrTran(cTexto,"à","&agrave;")
	cTexto := StrTran(cTexto,"â","&acirc;" )
	cTexto := StrTran(cTexto,"ã","&atilde;")
	cTexto := StrTran(cTexto,"ä","&auml;" )
	cTexto := StrTran(cTexto,"ó","&oacute;")
	cTexto := StrTran(cTexto,"ò","&ograve;")
	cTexto := StrTran(cTexto,"ô","&ocirc;" )
	cTexto := StrTran(cTexto,"õ","&otilde;")
	cTexto := StrTran(cTexto,"é","&eacute;")
	cTexto := StrTran(cTexto,"è","&egrave;")
	cTexto := StrTran(cTexto,"ê","&ecirc;" )
	cTexto := StrTran(cTexto,"í","&iacute;")
	cTexto := StrTran(cTexto,"ì","&igrave;")
	cTexto := StrTran(cTexto,"î","&icirc;" )
	cTexto := StrTran(cTexto,"ú","&uacute;")
	cTexto := StrTran(cTexto,"ù","&ugrave;")
	cTexto := StrTran(cTexto,"û","&ucirc;" )

	cTexto := StrTran(cTexto,"Ç","&Ccedil;")
	cTexto := StrTran(cTexto,"Á","&Aacute;")
	cTexto := StrTran(cTexto,"À","&Agrave;")
	cTexto := StrTran(cTexto,"Â","&Acirc;" )
	cTexto := StrTran(cTexto,"Ã","&Atilde;")
	cTexto := StrTran(cTexto,"Ä","&Auml;" )
	cTexto := StrTran(cTexto,"Ó","&Oacute;")
	cTexto := StrTran(cTexto,"Ò","&Ograve;")
	cTexto := StrTran(cTexto,"Ô","&Ocirc;" )
	cTexto := StrTran(cTexto,"Õ","&Otilde;")
	cTexto := StrTran(cTexto,"É","&Eacute;")
	cTexto := StrTran(cTexto,"È","&Egrave;")
	cTexto := StrTran(cTexto,"Ê","&Ecirc;" )
	cTexto := StrTran(cTexto,"Í","&Iacute;")
	cTexto := StrTran(cTexto,"Ì","&Igrave;")
	cTexto := StrTran(cTexto,"Î","&Icirc;" )
	cTexto := StrTran(cTexto,"Ú","&Uacute;")
	cTexto := StrTran(cTexto,"Ù","&Ugrave;")
	cTexto := StrTran(cTexto,"Û","&Ucirc;" )
	
	cTexto := StrTran(cTexto,"°","&deg;"   )
	cTexto := StrTran(cTexto,"º","&ordm;"  )
	                                        
	cTexto := StrTran(cTexto,"‡","?"	   )
	cTexto := StrTran(cTexto,"–","?"	   )
	cTexto := StrTran(cTexto,"’","?"	   )	
	cTexto := StrTran(cTexto,"£","&#163"   )	
	
Return(cTexto)
