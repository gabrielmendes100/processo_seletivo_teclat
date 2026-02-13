# SQL Query Modifications

## Overview
This document describes the modifications made to the SQL query for tables SB2070, SD4070, and SDD070.

## Changes Made

### Original Query
The original query selected only:
- `SB2.R_E_C_N_O_` (as B2RECCNO)
- `QTDEMP` (calculated field)

And grouped by:
- `SB2.R_E_C_N_O_`

### Modified Query
The modified query now includes additional fields from the SD4 table in both the SELECT statement and GROUP BY clause:

#### Added Fields (from SD4 table):
1. `D4_FILIAL` - Branch/Filial
2. `D4_COD` - Product Code
3. `D4_OP` - Production Order
4. `D4_TRT` - Treatment
5. `D4_LOTECTL` - Lot Control
6. `D4_NUMLOTE` - Lot Number
7. `D4_LOCAL` - Warehouse Location
8. `D4_ORDEM` - Order
9. `D4_OPORIG` - Original Production Order
10. `D4_SEQ` - Sequence

### Purpose
These fields are required to provide more detailed information about the production order requirements and inventory allocation in the query results. The fields follow the TOTVS Protheus standard indexing pattern:

```
D4_FILIAL + D4_COD + D4_OP + D4_TRT + D4_LOTECTL + D4_NUMLOTE + D4_LOCAL + D4_ORDEM + D4_OPORIG + D4_SEQ
```

### File Location
The modified query can be found in: `query_sb2_sd4_sdd.sql`

### Query Structure
- **Tables**: SB2070 (Stock by Warehouse), SD4070 (Production Order Requirements), SDD070 (Warehouse Allocations)
- **Joins**: LEFT JOIN between SB2 and SD4, LEFT JOIN between SB2 and SDD
- **Aggregation**: SUM of D4_QUANT and DD_QUANT to calculate QTDEMP (Total Quantity Allocated)
- **Filters**: Filial '01', Product 'CR19', Warehouse '02'
