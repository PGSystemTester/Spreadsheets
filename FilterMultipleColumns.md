# Filter Multiple Columns in Formula
Applies filter to an array of rows by checking conditions in EACH column. Differs from standard filter function which filters one column.

## Formula

### Lambda
`````
=Lambda(sRng,findText,
       Let(myFilter,BYROW(sRng,LAMBDA(r,OR(BYCOL(r,LAMBDA(c,c=findText))))),FILTER(sRng,myFilter)))
`````

### Let
`````
=LET(sRng,A1:C5,findText,"Tango",
       myFilter,BYROW(sRng,LAMBDA(r,  OR(BYCOL(r,LAMBDA(c,c=findText))))),FILTER(sRng,myFilter))
`````


## Example

*Objective to filter for rows with `Tango` in any column*

### Start Data

| ColumnA | ColumnB | ColumnC |
|---|---|---|
| Alpha | Donkey | Dog |
| Bravo | Tango | Cat |
| charlie | Alpha | Mivo |
| Delta | Charlie | Tango |
| Tango | Slice | Tango |
| Donkey | Bravo | Whale |


### Result After Formula
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

## Requirements
- Lambda Function
- [Byrow function](https://bettersolutions.com/excel/functions/byrow-function.htm)
