# Excel Custom Lambda Functions

## KeyCompare

Provides a fast, flexible, and formula-only solution for comparing structured datasets in Excel — no Power Query, VBA, or external tools required.

---

### Overview

`KeyCompare` is a reusable Excel LAMBDA function designed to compare two datasets (typically master data or dimension tables) and identify:

- New records  
- Removed records  
- Changed records (based on full row comparison)  
- Optional column-level change highlighting  

It is optimized for structured datasets where a **key column uniquely identifies each record**.

### 🧾 Parameters

| Parameter | Required | Description |
|----------|--------|-------------|
| `startData` | ✅ | Original dataset (baseline) |
| `endData` | ✅ | Updated dataset to compare against |
| `keyIdCol` |  | Column index of the unique key (default = 1) |
| `displayOnlyChanges` |  | If `TRUE`, masks unchanged column values |
| `txtNoChange` |  | Text used to represent unchanged values (default = `"....."`) |


### Formula
```
=LAMBDA(startData,endData,[keyIdCol],[displayOnlyChanges],[txtNoChange],
LET(idCol,IF(ISOMITTED(keyIdCol),1,keyIdCol),showChangesOnly,displayOnlyChanges,txtRemove,"Removed",txtChanged,"Changed",txtNew,"New",zSplitter,"→",txtNoChanges,"No Changes",
fxScopData,LAMBDA(_allData,LET(startRange,TRIMRANGE(_allData),FILTER(startRange,INDEX(startRange,,idCol)<>""))),sPure,fxScopData(startData),ePure,fxScopData(endData),idStartdata,INDEX(sPure,,idCol),idEndData,INDEX(ePure,,idCol),oldMembers,BYROW(sPure,LAMBDA(_,IF(ISNA(MATCH(INDEX(_,1,idCol),idEndData,0)),TRUE,TEXTJOIN(zSplitter,FALSE,_)))),
newMembers,BYROW(ePure,LAMBDA(_,
LET(existsInStart,ISNUMBER(MATCH(INDEX(_,1,idCol),idStartdata,0)),
        hasExactMatch,IF(existsInStart,LET(zJoinValues,TEXTJOIN(zSplitter,FALSE,_),ISNUMBER(MATCH(zJoinValues,oldMembers,0)))),
        IF(NOT(existsInStart),txtNew,
            IF(hasExactMatch,0,txtChanged)
        )))),fullBlock,VSTACK(HSTACK(IF(oldMembers=TRUE,txtRemove,0),sPure),HSTACK(newMembers,ePure)),zRay,FILTER(fullBlock,INDEX(fullBlock,,1)<>0,txtNoChanges),boolChoiceDisplay,IF(displayOnlyChanges,INDEX(zRay,1,1)<>txtNoChanges),IF(boolChoiceDisplay,
LET(txtNoChange,IF(ISOMITTED(txtNoChange),".....",txtNoChange),deltaRay,MAKEARRAY(ROWS(zRay),COLUMNS(zRay),LAMBDA(r,c,LET(newValue,INDEX(zRay,r,c),iType,INDEX(zRay,r,1),zKeyId,INDEX(zRay,r,2),IF(OR(c<3,iType<>txtChanged),newValue,LET(oldValue,XLOOKUP(zKeyId,INDEX(sPure,,1),INDEX(sPure,,c-1),newValue),IF(oldValue=newValue,txtNoChange,newValue)))))),deltaRay),zRay)))
```

### 📸 Before vs After Example

#### Input Data

**Start Data**

| ID | Name  | Region | Amount |
|----|------|--------|--------|
| A1 | Alpha | West   | 100    |
| A2 | Beta  | East   | 200    |
| A3 | Gamma | North  | 300    |

**End Data**

| ID | Name  | Region | Amount |
|----|------|--------|--------|
| A1 | Alpha | West   | 150    |
| A2 | Beta  | East   | 200    |
| A4 | Delta | South  | 400    |

---

#### Output (Standard)

| Status  | ID | Name  | Region | Amount |
|--------|----|------|--------|--------|
| Changed | A1 | Alpha | West   | 150    |
| Removed | A3 | Gamma | North  | 300    |
| New     | A4 | Delta | South  | 400    |

---

#### Output (displayOnlyChanges = TRUE)

| Status  | ID | Name  | Region | Amount |
|--------|----|------|--------|--------|
| Changed | A1 | ..... | ..... | 150    |
| Removed | A3 | Gamma | North | 300    |
| New     | A4 | Delta | South | 400    |


### Design Notes

- Uses `TEXTJOIN` to create row-level fingerprints for fast comparison  
- Avoids full dataset joins for performance  
- Built entirely with native Excel dynamic array functions  
- Handles:
  - Missing keys  
  - Blank rows  
  - Variable dataset sizes  

### Assumptions

- Key column must uniquely identify records  
- Column structure between datasets must match  
- Data types should be consistent between datasets  

### Use Cases

- SharePoint Lists, Masterdata Validation
- Pre/post load reconciliation  
- Data migration validation  
- Financial dimension auditing  
- Comparing exports from different systems  

####  Optional Feature: Change Highlighting

When `displayOnlyChanges = TRUE`:

- Only changed column values are shown  
- Unchanged columns are replaced with a placeholder (default: `"....."`)


### Behavior

#### ✔ New Records
- Present in `endData` but not in `startData`  
- Status = `New`

#### ✔ Removed Records
- Present in `startData` but not in `endData`  
- Status = `Removed`

#### ✔ Changed Records
- Same key exists in both datasets  
- At least one column differs  
- Status = `Changed`

#### ✔ Unchanged Records
- Excluded from output  

### 📄 License

Use freely. Modify as needed. Attribution appreciated.


----
## ReverseArray
Reverses an array. If more than one column is selected the reverse order is by row and then column.

### Parameters
- **array**: The array to be reversed.

### Example

#### Before

| col A | col B | col C |
|---|---|---|
| 1 | 2 | 3 |
| 4 | 5 | 6 |
| 7 | 8 | 9 |
| 10 | 11 | 12 |

#### After
| col A | col B | col C |
|---|---|---|
| 12 | 11 | 10 |
| 9 | 8 | 7 |
| 6 | 5 | 4 |
| 3 | 2 | 1 |


### Formula
````
=LAMBDA(array,LET(rCount,ROWS(array),iFlat,TOCOL(array),vRowCount,ROWS(iFlat),vRay,
SEQUENCE(vRowCount,1),zSeq,SEQUENCE(rCount,COLUMNS(array),vRowCount,-1),XLOOKUP(zSeq,vRay,iFlat)))
````


## Get Last Row
Gets max last row from column or columns ranges.

### Parameters
- **col_Range**: The range of columns to be searched.


### Formula
````
'=LAMBDA(col_Range,ROWS(TRIMRANGE(col_Range,2)))'
````

## UnpivotData

Moves dynamic array of column axis to row axis. The row and column axis are not bound to the data set, thus users can select an entire row(s)/column.

### Parameters
- **iRows**: The row axis of the data. This is not bound by the data range, thus selecting an entire column(s) will work.
- **iCols**: Column axis of the data, also not bound by the data range, thus selecting an entire row(s) will work.
- **iData**: The data axis. This must be mapped accurately to match where the data will exist, however using zero suppression empowers the user to extend this with limited concern.
- **suppressZero**: Optional boolean variable. Setting to `true` will remove all blanks and zero values from output. By default all values will be included.


### Formula
`````
=LAMBDA(iRows,iCols,iData,[suppressZero],LET(
    tCol, COLUMNS(iData),dColAx,ROWS(iCols),
    tRow, ROWS(iData),dRowAx,COLUMNS(iRows),
    cPad, COLUMN(INDEX(iData,1,1))-COLUMN(INDEX(iCols,1,1)),
    rPad, ROW(INDEX(iData,1,1))-ROW(INDEX(iRows,1,1)),
    aRow, INDEX(iRows,1+rPad,1):INDEX(iRows,tRow+rPad,dRowAx),
    aCol, INDEX(iCols,1,1+cPad):INDEX(iCols,dColAx,tCol+cPad),
   tDims, dRowAx+dColAx,
   tCols, tDims+1,
    zSeq, SEQUENCE(tRow*tCol,tCols,0),
    zDiv, tCols*tCol,
   zAddr, dRowAx-1,
    zMod, MOD(zSeq,tCols),
       c, MOD(INT(zSeq/tCols),tCol)+1,
       r, INT(zSeq/zDiv)+1,iResult,
   IF(zMod<dRowAx,INDEX(aRow,r,zMod+1),
      IF(zMod<tDims,INDEX(aCol,zMod-zAddr,c),INDEX(iData,r,c))),
IF(suppressZero,FILTER(iResult,INDEX(iResult,,tCols)<>0),iResult)))
`````

### Example

#### Before

|  |  |  |  | Jan | Feb | Mar | Apr | May |
|---|---|---|---|---|---|---|---|---|
|  |  |  |  | Actual | Actual | Actual | Budget | Budget |
|  |  |  |  |  |  |  |  |  |
| Disney | Tickets | Revenue |  | 507 | 607 | 707 | 807 | 907 |
| Disney | Movies | Costs |  | 508 | 608 | 708 | 808 | 908 |
| LucasFilm | Promo | Revenue |  | 509 | 609 | 709 | 809 | 909 |
| LucasFilm | Vader | Taxes |  | 510 | 610 | 710 | 810 | 910 |
| Marvel | HR | Costs |  | 511 | 611 | 711 | 811 | 911 |

#### After

| Disney | Tickets | Revenue | Jan | Actual | 507 |
|---|---|---|---|---|---|
| Disney | Tickets | Revenue | Feb | Actual | 607 |
| Disney | Tickets | Revenue | Mar | Actual | 707 |
| Disney | Tickets | Revenue | Apr | Budget | 807 |
| Disney | Tickets | Revenue | May | Budget | 907 |
| Disney | Movies | Costs | Jan | Actual | 508 |
| Disney | Movies | Costs | Feb | Actual | 608 |
| Disney | Movies | Costs | Mar | Actual | 708 |
| Disney | Movies | Costs | Apr | Budget | 808 |
| Disney | Movies | Costs | May | Budget | 908 |
| LucasFilm | Promo | Revenue | Jan | Actual | 509 |
| LucasFilm | Promo | Revenue | Feb | Actual | 609 |
| LucasFilm | Promo | Revenue | Mar | Actual | 709 |
| LucasFilm | Promo | Revenue | Apr | Budget | 809 |
| LucasFilm | Promo | Revenue | May | Budget | 909 |
| LucasFilm | Vader | Taxes | Jan | Actual | 510 |
| LucasFilm | Vader | Taxes | Feb | Actual | 610 |
| LucasFilm | Vader | Taxes | Mar | Actual | 710 |
| LucasFilm | Vader | Taxes | Apr | Budget | 810 |
| LucasFilm | Vader | Taxes | May | Budget | 910 |
| Marvel | HR | Costs | Jan | Actual | 511 |
| Marvel | HR | Costs | Feb | Actual | 611 |
| Marvel | HR | Costs | Mar | Actual | 711 |
| Marvel | HR | Costs | Apr | Budget | 811 |
| Marvel | HR | Costs | May | Budget | 911 |


## StackSplit
Splits array members based on text values in each element.

### Parameters
- **dataCol**: Single column with text values that are to be split
- **rowAxis**: (Optional) Datarange that will be repeat/pivoted with new elements
- **splitChar**: (Optional) Text character(s) to be used to split elements
- **includeIndex**: (Optional) Adds column to the left with the lookup index (1-based).

### Example

#### Before
| ColA                  |
|-----------------------|
| echo,charlie          |
| bravo,zulu,kilo,hotel |
| alpha,tango           |
| juliet,sierra         |

#### After
| ColA    |
|---------|
| echo    |
| charlie |
| bravo   |
| zulu    |
| kilo    |
| hotel   |
| juliet  |
| sierra  |
| alpha   |
| tango   |


### Formula
````F#
=LAMBDA(dataCol,[rowAxis],[splitChar],[includeIndex],LET(Θ,IF(ISOMITTED(splitChar),",",splitChar),zvals,
TRANSPOSE(TEXTSPLIT(TEXTJOIN(Θ,FALSE,dataCol),Θ)),iRef,VSTACK(0,SCAN(0,DROP(dataCol,IF(ROWS(dataCol)=1,0,-1)),
LAMBDA(oldval,eRow,oldval+1+LEN(eRow)-LEN(SUBSTITUTE(eRow,Θ,""))))),iRow,MATCH(SEQUENCE(ROWS(zvals),1,0,1),iRef,1),
testRowAxis,IF(ISOMITTED(rowAxis),zvals,HSTACK(INDEX(rowAxis,iRow,SEQUENCE(1,COLUMNS(rowAxis))),zvals)),
IF(includeIndex=TRUE,HSTACK(iRow,testRowAxis),testRowAxis)))
````


## Reconcile
- Takes two datasets and lists numberical differences
- Must have same number of columns, rows are dynamic
- Data is assumed to be far right column
- _Optional Paremeters:_
     -  **maxTolerance**: sets maximum amount of delta to not include (defaults to zero if ommited)
     -  **excludeHeader**: set to `TRUE` to exclude header row
     -  **lableStartData**: term to show for starting data
     -  **lableEndData**: term to show for ending data


### Formula
````F#
=LAMBDA(startData,endData,[maxTolerance],[excludeHeader],[lableStartData],[lableEndData],[lableDelta],LET(termWrongColCount,"# cols differ",termStart,IF(ISOMITTED(lableStartData),"Starting",lableStartData),termEnd,IF(ISOMITTED(lableEndData),"Ending",lableEndData),termChanges,IF(ISOMITTED(lableDelta),"Changes",lableDelta),termNoChanges,"No Changes",zColumnCount,COLUMNS(startData),IF(zColumnCount<>COLUMNS(endData),termWrongColCount,LET(pStart,TRIMRANGE(endData),pEnd,TRIMRANGE(startData),limitValueDelta,ABS(IF(ISOMITTED(maxTolerance),0,maxTolerance)),combinedData,VSTACK(HSTACK(pStart,MAKEARRAY(ROWS(pStart),1,LAMBDA(r,c,termStart))),HSTACK(pEnd,MAKEARRAY(ROWS(pEnd),1,LAMBDA(r,c,termEnd)))),rowAxis,DROP(combinedData,,-2),numData,INDEX(combinedData,,zColumnCount), zPvt,DROP(PIVOTBY(rowAxis,TAKE(combinedData,,-1),numData,SUM,0,0,,0),1),zStartDataCol,INDEX(zPvt,,zColumnCount),zEndDataCol,INDEX(zPvt,,zColumnCount+1),zDelta,N(zEndDataCol)-N(zStartDataCol),boolDelta,ABS(zDelta)>limitValueDelta, dataWithDelta,HSTACK(zPvt,zDelta),zFilteredData,FILTER(dataWithDelta,boolDelta,termNoChanges),zTopHeader,IF(COUNTA(zFilteredData)>1, MAKEARRAY(1,zColumnCount+2,LAMBDA(r,c,IF(c=zColumnCount,termStart,IF(c=zColumnCount+1,termEnd,IF(c=zColumnCount+2,termChanges,""))))),""),IF(excludeHeader,zFilteredData,IF(COUNTA(zFilteredData)=1,zFilteredData,VSTACK(zTopHeader,zFilteredData)))))))
````
----

## excludeColumns
Excudes columns as a numeric array. Exactly the opposite of ChooseCols.

### Parameters
- **allData**: Data to apply exclusion on .
- **columnsToExclude**: Column(s) to exclude. Can be a single integer or an array. 1-based (consistent with excel)
    - For array use `{1,3,5}` format
    - Can also use sequence as `Sequence(1,3,1,2)`

### Formula
````F#
 =LAMBDA(allData,columnsToExclude,LET(idx,SEQUENCE(1,COLUMNS(allData)),
 chkList,hstack(columnsToExclude),keepMask,ISNA(MATCH(idx,chkList, 0)),FILTER(allData,keepMask)))
````

## FilterMultiple
Applies the same filter to more than one column.

### Parameters
- **rng2Filter**: Data to be evaluated on a row by row basis. Note that this will also be returned unless third parameter `returnRange` is provided.
- **ƒeachCell**: Lambda function that must have one input that will evaluate each cell and return `true` or `false`.
    - Example `Lambda(eachCell,eachCell="Actual")`
- **returnRange**: Optional parameter to provide a different range to return after filter. Note if row count of `rng2Filter` and `returnRange` are not equal, value of `#RowCountMismatch` will be returned.


### Formula
````
=LAMBDA(rng2Filter,ƒeachCell,[returnRange],
    IF(IF(NOT(ISOMITTED(returnRange)),ROWS(returnRange)<>ROWS(rng2Filter)),"#RowCountMismatch",
    LET(zReturnRng,IF(ISOMITTED(returnRange),rng2Filter,returnRange), zFilter,BYROW(rng2Filter,LAMBDA(_,
    REDUCE(FALSE,_,LAMBDA(zOld,zNew,
    IF(zOld=TRUE,TRUE,ƒeachCell(zNew)))))),
    FILTER(zReturnRng,zFilter,"#noMatches"))))
````
### Example
The following examples are based on this starting dataset beginning in cell `a1`.

#### Starting Data
| colA | colB | colC | colD |
|---|---|---|---|
| Item1 | 4/30/2024 | Fines | 510 |
| Item2 | 6/15/2024 | Insurance | 443 |
| Item3 | 4/20/2025 | Insurance | 405 |
| Item4 | 7/25/2024 | Fees | 512 |
| Item5 | 7/16/2024 | Cogs | 216 |
| Item6 | 8/29/2025 | Cogs | 218 |
| Item7 | 4/2/2025 | Cogs | 357 |
| Item8 | 9/15/2024 | Fines | 271 |
| Item9 | 1/6/2025 | Fines | 323 |


#### Result A
`=multiFilter(C1:D9,LAMBDA(α,IF(α="Fines",TRUE,IF(ISNUMBER(α),α>500))),A1:D9)`

Returns all columns checking in alast two columnsif fines OR amount is over 500.


|  |  |  |  |
|---|---|---|---|
| Item1 | 4/30/2024 | Fines | 510 |
| Item4 | 7/25/2024 | Issues | 512 |
| Item8 | 9/15/2024 | Fines | 271 |
| Item9 | 1/6/2025 | Fines | 323 |

#### Result B
`=multiFilter(C1:D9,LAMBDA(α,α="Cogs"),HSTACK(B1:B9,D1:D9))`

Returns columns b and d checking if any cells in c:d have 'cogs'.


|  |  |
|---|---|
| 7/16/2024 | 216 |
| 8/29/2025 | 218 |
| 4/2/2025 | 357 |

