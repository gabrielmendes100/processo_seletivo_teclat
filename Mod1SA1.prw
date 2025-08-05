#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
#Include 'TopConn.ch'

Static cTitulo   := "Cadastro de Clientes"
Static cCamposSA1:= "A1_COD;A1_LOJA;A1_NOME;A1_CEP;A1_EST;A1_MUN;A1_END;A1_BAIRRO;A1_COMPLEM"

User Function Mod1SA1()
    Local aArea     := GetArea()
    Local oBrowse   := FWMBrowse():New()
    Local cCepVazio := "Empty(SA1->A1_CEP)"
    Local cEndVazio := "!Empty(SA1->A1_CEP) .AND. (Empty(SA1->A1_EST) .OR. Empty(SA1->A1_MUN) .OR. Empty(SA1->A1_END) .OR. Empty(SA1->A1_BAIRRO))"
    Local cOk       := "!Empty(SA1->A1_CEP) .AND. !Empty(SA1->A1_EST) .AND. !Empty(SA1->A1_MUN) .AND. !Empty(SA1->A1_END) .AND. !Empty(SA1->A1_BAIRRO)"
    Local aStruSA1  :={}
    Local aColumns  := {}
    Local nUltCol   := nAtual := 0

    Private lMsErroAuto :=.F. 
    Private lAutoErrNoFile := .T.

    DBSelectArea("SA1")
    aStruSA1 := SA1->(DbStruct())

    oBrowse:SetAlias("SA1")
    
    For nAtual := 1 To Len(aStruSA1)
        If    Alltrim(aStruSA1[nAtual][1]) $ cCamposSA1
            aAdd(aColumns,FWBrwColumn():New())
            nUltCol := Len(aColumns)
            
            aColumns[nUltCol]:SetData( &("{||"+aStruSA1[nAtual][1]+"}") )
            aColumns[nUltCol]:SetTitle(RetTitle(aStruSA1[nAtual][1])) 
            aColumns[nUltCol]:SetSize(aStruSA1[nAtual][3]) 
            aColumns[nUltCol]:SetDecimal(aStruSA1[nAtual][4])
            aColumns[nUltCol]:SetPicture(PesqPict("SA1",aStruSA1[nAtual][1])) 
        EndIf     
    Next nAtual 
    oBrowse:AddLegend(cCepVazio, "BR_VERMELHO", "CEP em Branco")
    oBrowse:AddLegend(cEndVazio, "BR_AMARELO" , "CEP preenchido, mas os demais campos vazios")
    oBrowse:AddLegend(cOk      , "BR_VERDE"   , "Campos de Endereco preenchidos")
    oBrowse:SetColumns(aColumns)
    oBrowse:SetDescription(cTitulo)
    oBrowse:Activate()
     
    RestArea(aArea)
Return Nil
 
Static Function MenuDef()
    Local aRot := {}
    
    ADD OPTION aRot TITLE 'Visualizar'                            ACTION 'VIEWDEF.Mod1SA1' OPERATION MODEL_OPERATION_VIEW ACCESS 0
    ADD OPTION aRot TITLE 'Legendas'                              ACTION 'U_LegMod1()' OPERATION 7 ACCESS 0
    ADD OPTION aRot TITLE 'Atualizar Endereco por CEP (manual)  ' ACTION 'U_ModAtuWS()' OPERATION 4 ACCESS 0
    ADD OPTION aRot TITLE 'Atualizar Endereco via CSV (em Massa)' ACTION 'U_ModAtuCsv()' OPERATION 4 ACCESS 0
 
Return aRot

Static Function ModelDef()
    Local oModel := Nil
    Local oStruSA1 := FWFormStruct(1, "SA1")

    oModel := MPFormModel():New("Mod1SA1M") 
     
    oModel:AddFields("FORMSA1",,oStruSA1)
          
    oModel:SetDescription("Modelo de Dados do Cadastro "+cTitulo)
     
    oModel:GetModel("FORMSA1"):SetDescription("Formulario do Cadastro "+cTitulo)
Return oModel
 
Static Function ViewDef()
    Local oModel := FWLoadModel("Mod1SA1")
    Local oStruSA1 := FWFormStruct(2, "SA1",{|x| Alltrim(x) $  cCamposSA1})  
    Local oView := Nil
 
    oView := FWFormView():New()
    oView:SetModel(oModel)
    oView:AddField("VIEW_SA1", oStruSA1, "FORMSA1")
    oView:CreateHorizontalBox("TELA",100)
    oView:EnableTitleView('VIEW_SA1', 'Dados - ' + cTitulo )  
    oView:SetCloseOnOk({||.T.})
    oView:SetOwnerView("VIEW_SA1","TELA")
     
Return oView

User Function LegMod1()
    Local aLegenda := {}

    aAdd(aLegenda,{"BR_VERMELHO",     "CEP em Branco"})
    aAdd(aLegenda,{"BR_AMARELO",      "CEP preenchido, mas os demais campos vazios"})
    aAdd(aLegenda,{"BR_VERDE",        "Campos de Endereco preenchidos"})
     
    BrwLegenda("Legendas", "Legendas - Clientes", aLegenda)
Return

UsER Function ModAtuCsv()
    Local cTitulo  := "Atualizar dados - CSV"
    Local oDlg
    Local oFont
    DEFINE MSDIALOG oDlg TITLE cTitulo FROM 000, 000  TO 160, 600 COLORS 0, 16777215 PIXEL

        oFont := TFont():New('Courier new',,-18,.T.)

        TSay():New(010, 010, {|| "Rotina de Atualizacao de dados de Endereco via csv;"}                             , oDlg, , , , , , .T., , , 200, 10)
        TSay():New(020, 010, {|| "A rotina irá atualizar os dados de Endereco dos clientes;"}                       , oDlg, , , , , , .T., , , 200, 10)
        
        TButton():New(060, 150, "Baixar CSV"    , oDlg, {|| BaixarCsv()            }  , 40, 012, , , , .T., , , , , ,)
        TButton():New(060, 200, "Importar"      , oDlg, {|| ImpCsv()   , oDlg:End()}  , 40, 012, , , , .T., , , , , ,)
        TButton():New(060, 250, "Sair"          , oDlg, {|| oDlg:End()}               , 40, 012, , , , .T., , , , , ,)
        
    ACTIVATE MSDIALOG oDlg CENTERED

Return 
Static Function BaixarCsv()
    Local cLinha     := ""
    Local cArquivo   := GetTempPath() + "clientes.csv"
    Local nHandle    := 0
    Local cQuery     := ""
    
    if MsgNoYes("Deseja baixar o CSV de modelo para a Atualizacao de dados?")

        cLinha := "CODIGO;LOJA;CEP;ENDERECO;COMPLEMENTO;BAIRRO;CIDADE;UF" + CRLF

        cQuery := " SELECT A1_COD, A1_LOJA, A1_CEP, A1_END, A1_COMPLEM, A1_BAIRRO, A1_MUN, A1_EST "
        cQuery += " FROM " + RetSqlName("SA1") + " SA1 "
        cQuery += " WHERE SA1.D_E_L_E_T_ = ' ' "
        cQuery += " AND A1_FILIAL = '" + xFilial("SA1") + "' "
        cQuery += " ORDER BY A1_COD, A1_LOJA "

        nHandle := FCreate(cArquivo)

        If nHandle < 0
            MsgStop("Erro ao criar arquivo: " + Str(FError()))
            Return
        EndIf
        
        if Select("QRYSA1") > 0
            QRYSA1->(DbCloseArea())
        Endif   
        
        TcQuery cQuery New Alias "QRYSA1"

        While !QRYSA1->(EOF())
            cLinha += QRYSA1->A1_COD + ";"
            cLinha += QRYSA1->A1_LOJA + ";"
            cLinha += QRYSA1->A1_CEP + ";"
            cLinha += QRYSA1->A1_END + ";"
            cLinha += QRYSA1->A1_COMPLEM + ";"
            cLinha += QRYSA1->A1_BAIRRO + ";"
            cLinha += QRYSA1->A1_MUN + ";"
            cLinha += QRYSA1->A1_EST
            cLinha += CRLF

            QRYSA1->(DbSkip())
        EndDo

        MemoWrite(cArquivo,cLinha)
        If File(cArquivo)
            MsgInfo("Arquivo gerado com sucesso em: " +  CHR(13) + CHR(10) + cArquivo)
            ShellExecute("open", cArquivo, "", "", 1)
        EndIf
    endif
Return

Static Function ImpCsv()
    Local i,j
    Local cArquivo    := ""
    Local aCabecalho  := {}
    Local aFields     := {}
    Local aLinhas     := {}
    Local aExecAuto   := {}
    Local cDado       := ""

    cArquivo := cGetFile( 'Arquivo CSV|*.csv','Selecao de Arquivos') 

    if  MsgNoYes("Deseja Prosseguir com a Atualizacao?")

        aAdd(aFields, {"CODIGO"     , "A1_COD"    })
        aAdd(aFields, {"LOJA"       , "A1_LOJA"   })
        aAdd(aFields, {"CEP"        , "A1_CEP"    })
        aAdd(aFields, {"ENDERECO"   , "A1_END"    })
        aAdd(aFields, {"COMPLEMENTO", "A1_COMPLEM"})
        aAdd(aFields, {"BAIRRO"     , "A1_BAIRRO" })
        aAdd(aFields, {"CIDADE"     , "A1_MUN"    })
        aAdd(aFields, {"UF"         , "A1_EST"    })

        If !File(cArquivo)
            MsgStop("Arquivo CSV nao encontrado!")
            Return
        EndIf

        aLinhas := FileToArr(cArquivo)

        If Len(aLinhas) <= 1
            MsgStop("Arquivo CSV vazio ou contem apenas cabecalho!")
            Return
        EndIf

        aCabecalho := Separa(aLinhas[1], ";")

        For i := 2 to Len(aLinhas)
            aExecAuto := {}
            aLinha := Separa(aLinhas[i], ";", .T.)

            For j:= 1 to Len(aLinha)
                if  aFields[j][2] == "A1_LOJA"
                    cDado := Padl(Alltrim(Decodeutf8(aLinha[j])),2,"0")
                else
                    cDado := Alltrim(Decodeutf8(aLinha[j]))
                    
                    if  aFields[j][2] == "A1_MUN"
                        aAdd(aExecAuto, {"A1_COD_MUN", Posicione("CC2",2,xFilial("CC2")+Upper(cDado),"CC2_CODMUN"),NIL })
                    Endif

                endif       

                aAdd(aExecAuto, { aFields[j][2],Upper(cDado),NIL })
            Next

            lMsErroAuto := .F.
            
            aExecAuto := FWVetByDic( aExecAuto, 'SA1' )

            If GetMv("MV_MVCSA1")
                MSExecAuto({|x,y| CRMA980(x,y)}, aExecAuto, 4)
            else
                MSExecAuto({|x,y| MATA030(x,y)}, aExecAuto, 4)
            EndIf

            IF lMsErroAuto
                AutoGrLog("Erro na linha " + cValToChar(i) + " do arquivo")
                cLog := StrTran(ArrTokStr(GetAutoGrLog()),"|",CRLF)
                nOpcAviso := Aviso("Erro na importacao",;
                cLog + CRLF + ;
                'Esta mensagem fechara em 5 segundos. Para interromper o timer, selecione "Timer Off". ',;
                {"OK"},,, 1,,,5)

            ENDIF
        Next
    endif
Return 

User Function ModAtuWs()
    Local cTitulo    := "Atualizar dados - Via CEP"
    Local oDlg
    Local oRest      := FWRest():New("https://viacep.com.br")
    Local cCep       := SA1->A1_CEP
    Local cEnd       := SA1->A1_END
    Local cComp      := SA1->A1_COMPLEM
    Local cBairro    := SA1->A1_BAIRRO  
    Local cCidade    := SA1->A1_MUN
    Local cUF        := SA1->A1_EST 

    DEFINE MSDIALOG oDlg TITLE cTitulo FROM 000, 000  TO 500, 700  PIXEL
        @ 010, 010 SAY "CEP:"         SIZE 050, 007 OF oDlg PIXEL 
        @ 040, 010 SAY "Endereco:"    SIZE 050, 007 OF oDlg PIXEL
        @ 070, 010 SAY "Complemento:" SIZE 050, 007 OF oDlg PIXEL
        @ 100, 010 SAY "Bairro:"      SIZE 050, 007 OF oDlg PIXEL
        @ 130, 010 SAY "Cidade:"      SIZE 050, 007 OF oDlg PIXEL
        @ 160, 010 SAY "UF:"          SIZE 050, 007 OF oDlg PIXEL
        
        @ 020, 010 MSGET cCep     PICTURE "@R 99999-999" SIZE 050, 011 OF oDlg PIXEL VALID ConsultaCEP(oRest, @cEnd, @cComp, @cBairro, @cCidade, @cUF, cCep)
        @ 050, 010 MSGET cEnd     WHEN .F. SIZE 270, 011 OF oDlg PIXEL
        @ 080, 010 MSGET cComp    WHEN .F. SIZE 270, 011 OF oDlg PIXEL
        @ 110, 010 MSGET cBairro  WHEN .F. SIZE 270, 011 OF oDlg PIXEL
        @ 140, 010 MSGET cCidade  WHEN .F. SIZE 270, 011 OF oDlg PIXEL
        @ 170, 010 MSGET cUF      WHEN .F. SIZE 020, 011 OF oDlg PIXEL
        
        TButton():New(200, 010, "Atualizar Cliente", oDlg, {||AtuSA1WS( cEnd, cComp, cBairro, cCidade, cUF, cCep),oDlg:End()}, 080, 012, , , , .T., , , , , ,)
        TButton():New(200, 100, "Sair"    , oDlg, {|| oDlg:End()}, 080, 012, , , , .T., , , , , ,)
        
    ACTIVATE MSDIALOG oDlg CENTERED
Return

Static Function ConsultaCEP(oRest, cEnd, cComp, cBairro, cCidade, cUF, cCep)
    Local cJson := ""
    Local oJson := Nil

    cCep := StrTran(cCep,"-","")
    
    oRest:SetPath("/ws/"+cCep+"/json/")
    
    cEnd    := ""
    cComp   := ""
    cBairro := ""
    cCidade := ""
    cUF     := ""

    If oRest:Get() 
        cJson := oRest:GetResult()
        oJson := JsonObject():New()
        oJson:FromJson(cJson)

        if AlltoChar(oJson:GetJsonObject("erro")) != "true"        
            cEnd    := Upper(Decodeutf8(Alltochar(oJson:GetJsonObject("logradouro"))))
            cComp   := Upper(Decodeutf8(Alltochar(oJson:GetJsonObject("complemento"))))
            cBairro := Upper(Decodeutf8(Alltochar(oJson:GetJsonObject("bairro"))))
            cCidade := Upper(Decodeutf8(Alltochar(oJson:GetJsonObject("localidade"))))
            cUF     := Upper(Decodeutf8(Alltochar(oJson:GetJsonObject("uf"))))
        else
            MsgStop("CEP nao encontrado!")
        endif
    Else
        MsgStop("CEP nao encontrado!")
    EndIf
Return .t.
Static Function AtuSA1WS(cEnd, cComp, cBairro, cCidade, cUF, cCep)
    Local aExecAuto := {}

    If Empty(SA1->A1_COD)
        MsgStop("Selecione um cliente primeiro!")
        Return
    EndIf

    aAdd(aExecAuto, {"A1_COD"    , SA1->A1_COD   , Nil})
    aAdd(aExecAuto, {"A1_LOJA"   , SA1->A1_LOJA  , Nil})
    aAdd(aExecAuto, {"A1_CEP"    , cCep          , Nil})
    aAdd(aExecAuto, {"A1_END"    , cEnd          , Nil}) 
    aAdd(aExecAuto, {"A1_COMPLEM", cComp         , Nil})
    aAdd(aExecAuto, {"A1_BAIRRO" , cBairro       , Nil})
    aAdd(aExecAuto, {"A1_EST"    , cUF           , Nil})
    aAdd(aExecAuto, {"A1_COD_MUN", Posicione("CC2",2,xFilial("CC2")+Upper(cCidade),"CC2_CODMUN"), Nil})
    aAdd(aExecAuto, {"A1_MUN"    , cCidade       , Nil})

    aExecAuto := FWVetByDic( aExecAuto, 'SA1' )

    lMsErroAuto := .F.
    If GetMv("MV_MVCSA1") 
        MSExecAuto({|x,y| CRMA980(x,y)}, aExecAuto, 4)
    else
        MSExecAuto({|x,y| MATA030(x,y)}, aExecAuto, 4)
    EndIf

    If lMsErroAuto
        cLog := StrTran(ArrTokStr(GetAutoGrLog()),"|",CRLF)
        Aviso("Erro na importacao", cLog )
    Else
        FwAlertSuccess("Cliente atualizado com sucesso!")
    EndIf

Return
