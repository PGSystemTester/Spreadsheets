# Excel Formulas in Google Sheets
These formulas are currently available in Excel, but not google sheets. By adding these to a google sheet **TOOLS → Named Functions** similar functionality can be created. At the time of these posted, there is no parameter for optional variables.



# Take
[Take](https://exceljet.net/functions/take-function) is currently not available in Google sheets. There are several workarounds.



## KeepColumns
Similar to Excel's TAKE function, but scoped specifically to Columns

### Parameters
 - **ARRAY**: Range/Array to modify
 - **ColumnsToInclude**: Negative value implies reversed direction (i.e. `-1` would be final column)


### Function in Named Formulas
```
=if(ColumnsToInclude>0,index(ARRAY,1,1):index(ARRAY,rows(ARRAY),ColumnsToInclude),
index(ARRAY,rows(ARRAY),columns(ARRAY)):index(ARRAY,1,columns(ARRAY)+ColumnsToInclude+1))
```

### Examples
- `=KeepColumns(A1:D9,3)` returns array of `A1:C9`
- `=KeepColumns(A1:D9,-3)` returns array of `B1:D9`

## KeepRows
Similar to Excel's TAKE function, but scoped specifically to Rows

### Parameters
 - **ARRAY**: Range/Array to modify
 - **RowsToInclude**: Negative value implies reversed direction (i.e. `-1` would be final row)


### Function in Named Formulas
```
=if(RowsToInclude>0,index(ARRAY,1,1):index(ARRAY,RowsToInclude,columns(ARRAY)),
index(ARRAY,rows(ARRAY),columns(ARRAY)):index(ARRAY,rows(array)+1+RowsToInclude,1))
```


### Examples
- `=KeepRows(A1:D9,3)` returns array of `A1:D3`
- `=KeepRows(A1:D9,-3)` returns array of `A7:D9`


## Combined KeepCol and KeepRow
- Can combined both to resemble TAKE function in excel
- `=keepCols(keepRows(A1:D9,3),-3)` returns `B1:D3`
