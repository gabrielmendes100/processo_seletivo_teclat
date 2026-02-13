#Include 'Protheus.ch'
#Include 'TopConn.ch'

/*/{Protheus.doc} QuerySB2SD4
    Query para buscar dados de SB2 com empenhamento de SD4 e SDD
    @type  Function
    @author Gabriel Mendes
    @since 13/02/2026
/*/
User Function QuerySB2SD4()
    Local cQuery := ""
    
    cQuery := "SELECT SB2.R_E_C_N_O_ AS B2RECCNO,                                           " + CRLF
    cQuery += " SD4.D4_FILIAL,                                                              " + CRLF
    cQuery += " SD4.D4_COD,                                                                 " + CRLF
    cQuery += " SD4.D4_OP,                                                                  " + CRLF
    cQuery += " SD4.D4_TRT,                                                                 " + CRLF
    cQuery += " SD4.D4_LOTECTL,                                                             " + CRLF
    cQuery += " SD4.D4_NUMLOTE,                                                             " + CRLF
    cQuery += " SD4.D4_LOCAL,                                                               " + CRLF
    cQuery += " SD4.D4_ORDEM,                                                               " + CRLF
    cQuery += " SD4.D4_OPORIG,                                                              " + CRLF
    cQuery += " SD4.D4_SEQ,                                                                 " + CRLF
    cQuery += " ISNULL(SUM(SD4.D4_QUANT), 0) + ISNULL(SUM(SDD.DD_QUANT), 0) AS QTDEMP       " + CRLF
    cQuery += " FROM " + RetSQLName("SB2") + " SB2                                          " + CRLF
    cQuery += "LEFT JOIN " + RetSQLName("SD4") + " SD4                                      " + CRLF
    cQuery += " ON SD4.D4_FILIAL    = SB2.B2_FILIAL                                         " + CRLF
    cQuery += " AND SD4.D4_COD      = SB2.B2_COD                                            " + CRLF
    cQuery += " AND SD4.D4_LOCAL    = SB2.B2_LOCAL                                          " + CRLF
    cQuery += " AND SD4.D4_QUANT   <> 0                                                     " + CRLF
    cQuery += " AND SD4.D_E_L_E_T_  = ' '                                                   " + CRLF
    cQuery += "LEFT JOIN " + RetSQLName("SDD") + " SDD                                      " + CRLF
    cQuery += " ON SDD.DD_FILIAL    = SB2.B2_FILIAL                                         " + CRLF
    cQuery += " AND SDD.DD_PRODUTO  = SB2.B2_COD                                            " + CRLF
    cQuery += " AND SDD.DD_LOCAL    = SB2.B2_LOCAL                                          " + CRLF
    cQuery += " AND SDD.DD_QUANT   <> 0                                                     " + CRLF
    cQuery += " AND SDD.D_E_L_E_T_  = ' '                                                   " + CRLF
    cQuery += "WHERE SB2.D_E_L_E_T_ = ' '                                                   " + CRLF
    cQuery += " AND SB2.B2_FILIAL  = '"+xFilial("SB2")+"'                                   " + CRLF
    If !(IsBlind())
        cQuery += " AND SB2.B2_COD     >= '"+MV_PAR01+"'                                    " + CRLF
        cQuery += " AND SB2.B2_COD     <= '"+MV_PAR02+"'                                    " + CRLF
        cQuery += " AND SB2.B2_LOCAL   >= '"+MV_PAR03+"'                                    " + CRLF
        cQuery += " AND SB2.B2_LOCAL   <= '"+MV_PAR04+"'                                    " + CRLF
    EndIf
    cQuery += " GROUP BY SB2.R_E_C_N_O_,                                                    " + CRLF
    cQuery += " SD4.D4_FILIAL,                                                              " + CRLF
    cQuery += " SD4.D4_COD,                                                                 " + CRLF
    cQuery += " SD4.D4_OP,                                                                  " + CRLF
    cQuery += " SD4.D4_TRT,                                                                 " + CRLF
    cQuery += " SD4.D4_LOTECTL,                                                             " + CRLF
    cQuery += " SD4.D4_NUMLOTE,                                                             " + CRLF
    cQuery += " SD4.D4_LOCAL,                                                               " + CRLF
    cQuery += " SD4.D4_ORDEM,                                                               " + CRLF
    cQuery += " SD4.D4_OPORIG,                                                              " + CRLF
    cQuery += " SD4.D4_SEQ                                                                  " + CRLF

Return cQuery
