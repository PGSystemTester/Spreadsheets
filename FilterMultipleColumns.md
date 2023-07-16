# Filter Multiple Columns in Formula
Uses Excel Array Features


## Start Data
| ColumnA | ColumnB | ColumnC |
|---|---|---|
| Alpha | Donkey | Dog |
| Bravo | Tango | Cat |
| charlie | Alpha | Mivo |
| Delta | Charlie | Tango |
| Tango | Slice | Tango |
| Donkey | Bravo | Whale |

*Objective to filter for rows with TANGO in any column*

## Formula

`````
=LET(sRng,A1:C5,findText,"Tango",
       myFilter,BYROW(sRng,LAMBDA(r,  OR(BYCOL(r,LAMBDA(c,c=findText))))),FILTER(sRng,myFilter))


`````
## Result
| ColumnA | ColumnB | ColumnC |
|---|---|---|
| Bravo | Tango | Cat |
| Delta | Charlie | Tango |
| Tango | Slice | Tango |

## Other Usages

### Exclude Members
*Excludes rows where donkey is in any cell*
`````
=LET(sRng,A1:C5,findText,"Donkey", 
   myFilter,BYROW(sRng,LAMBDA(r,OR(BYCOL(r,LAMBDA(c,c=findText))))),FILTER(sRng,NOT(myFilter)))
`````

### Cell Contains Text
Filters on rows where certain text exists in a cell in any column.
