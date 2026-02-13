# Modificações na Query SQL

## Descrição
Este documento descreve as modificações realizadas na query SQL conforme solicitado no processo seletivo.

## Objetivo
Adicionar os seguintes campos tanto no SELECT quanto no GROUP BY:
- D4_FILIAL
- D4_COD
- D4_OP
- D4_TRT
- D4_LOTECTL
- D4_NUMLOTE
- D4_LOCAL
- D4_ORDEM
- D4_OPORIG
- D4_SEQ

## Query Original

```sql
SELECT SB2.R_E_C_N_O_ AS B2RECCNO,
 ISNULL(SUM(SD4.D4_QUANT), 0) + ISNULL(SUM(SDD.DD_QUANT), 0) AS QTDEMP
 FROM SB2070 SB2
LEFT JOIN SD4070 SD4
 ON SD4.D4_FILIAL    = SB2.B2_FILIAL
 AND SD4.D4_COD      = SB2.B2_COD
 AND SD4.D4_LOCAL    = SB2.B2_LOCAL
 AND SD4.D4_QUANT   <> 0
 AND SD4.D_E_L_E_T_  = ' '
LEFT JOIN SDD070 SDD
 ON SDD.DD_FILIAL    = SB2.B2_FILIAL
 AND SDD.DD_PRODUTO  = SB2.B2_COD
 AND SDD.DD_LOCAL    = SB2.B2_LOCAL
 AND SDD.DD_QUANT   <> 0
 AND SDD.D_E_L_E_T_  = ' '
WHERE SB2.D_E_L_E_T_ = ' '
 AND SB2.B2_FILIAL  = '01'
 AND SB2.B2_COD     >= 'CR19           '
 AND SB2.B2_COD     <= 'CR19           '
 AND SB2.B2_LOCAL   >= '  '
 AND SB2.B2_LOCAL   <= 'ZZ'
 GROUP BY SB2.R_E_C_N_O_
```

## Query Modificada

```sql
SELECT SB2.R_E_C_N_O_ AS B2RECCNO,
 SD4.D4_FILIAL,
 SD4.D4_COD,
 SD4.D4_OP,
 SD4.D4_TRT,
 SD4.D4_LOTECTL,
 SD4.D4_NUMLOTE,
 SD4.D4_LOCAL,
 SD4.D4_ORDEM,
 SD4.D4_OPORIG,
 SD4.D4_SEQ,
 ISNULL(SUM(SD4.D4_QUANT), 0) + ISNULL(SUM(SDD.DD_QUANT), 0) AS QTDEMP
 FROM SB2070 SB2
LEFT JOIN SD4070 SD4
 ON SD4.D4_FILIAL    = SB2.B2_FILIAL
 AND SD4.D4_COD      = SB2.B2_COD
 AND SD4.D4_LOCAL    = SB2.B2_LOCAL
 AND SD4.D4_QUANT   <> 0
 AND SD4.D_E_L_E_T_  = ' '
LEFT JOIN SDD070 SDD
 ON SDD.DD_FILIAL    = SB2.B2_FILIAL
 AND SDD.DD_PRODUTO  = SB2.B2_COD
 AND SDD.DD_LOCAL    = SB2.B2_LOCAL
 AND SDD.DD_QUANT   <> 0
 AND SDD.D_E_L_E_T_  = ' '
WHERE SB2.D_E_L_E_T_ = ' '
 AND SB2.B2_FILIAL  = '01'
 AND SB2.B2_COD     >= 'CR19           '
 AND SB2.B2_COD     <= 'CR19           '
 AND SB2.B2_LOCAL   >= '  '
 AND SB2.B2_LOCAL   <= 'ZZ'
 GROUP BY SB2.R_E_C_N_O_,
 SD4.D4_FILIAL,
 SD4.D4_COD,
 SD4.D4_OP,
 SD4.D4_TRT,
 SD4.D4_LOTECTL,
 SD4.D4_NUMLOTE,
 SD4.D4_LOCAL,
 SD4.D4_ORDEM,
 SD4.D4_OPORIG,
 SD4.D4_SEQ
```

## Mudanças Realizadas

### No SELECT
Adicionados os seguintes campos após `B2RECCNO` e antes do campo calculado `QTDEMP`:
1. SD4.D4_FILIAL
2. SD4.D4_COD
3. SD4.D4_OP
4. SD4.D4_TRT
5. SD4.D4_LOTECTL
6. SD4.D4_NUMLOTE
7. SD4.D4_LOCAL
8. SD4.D4_ORDEM
9. SD4.D4_OPORIG
10. SD4.D4_SEQ

### No GROUP BY
Adicionados os mesmos campos após `SB2.R_E_C_N_O_`:
1. SD4.D4_FILIAL
2. SD4.D4_COD
3. SD4.D4_OP
4. SD4.D4_TRT
5. SD4.D4_LOTECTL
6. SD4.D4_NUMLOTE
7. SD4.D4_LOCAL
8. SD4.D4_ORDEM
9. SD4.D4_OPORIG
10. SD4.D4_SEQ

## Implementação
A query modificada foi implementada no arquivo `QuerySB2SD4.prw` na função `User Function QuerySB2SD4()`.

## Observações
- A query agora retorna mais colunas, permitindo uma visão detalhada dos registros de empenhamento (SD4)
- O GROUP BY foi ajustado para incluir todos os campos não agregados, conforme requerido pela sintaxe SQL
- A estrutura da query mantém a lógica original de agregação através da função SUM
