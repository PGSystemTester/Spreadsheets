# Excel Custom Lambda Functions

## ReverseArray
Reverses an array, capable of doing by columns


### Formula
````
=LAMBDA(iRay,LET(rCount,ROWS(iRay),iFlat,TOCOL(iRay),vRowCount,ROWS(iFlat),vRay,
SEQUENCE(vRowCount,1),zSeq,SEQUENCE(rCount,COLUMNS(iRay),vRowCount,-1),XLOOKUP(zSeq,vRay,iFlat)))
````


## Get Last Row
Gets max last row from column or columns ranges.

### Formula
````
=LAMBDA(pColumn,LET(ƒ,LAMBDA(xcol,LET(textVal,"Θ",numVal,9.99999999999999E+307,
finalCellInRange,INDEX(xcol,ROWS(xcol),0),firstSearch,
IFERROR(MATCH(textVal,xcol,1),0),firstSrchRng,
INDEX(xcol,firstSearch+1,1):finalCellInRange,IF(COUNTA(firstSrchRng)=0,firstSearch,
LET(secondSearch,IFERROR(MATCH(numVal,firstSrchRng,1),0)+firstSearch,secondSearchRng,
INDEX(xcol,secondSearch+1,1):finalCellInRange,IF(COUNTA(secondSearchRng)=0,secondSearch,
XMATCH(FALSE,ISBLANK(xcol),0,-1)))))),totCol,COLUMNS(pColumn),IF(totCol=1,ƒ(pColumn),
LET(S,SEQUENCE(1,totCol),REDUCE(0,S,LAMBDA(old,eachCol,LET(iCol,DROP(INDEX(pColumn,,eachCol),
old),old+ƒ(iCol))))))))
````



## UnpivotData

Moves dynamic array of column axis to row axis

### Formula
`````
=LAMBDA(iRows,iCols,iData,LET(
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
       r, INT(zSeq/zDiv)+1,
   IF(zMod<dRowAx,INDEX(aRow,r,zMod+1),
      IF(zMod<tDims,INDEX(aCol,zMod-zAddr,c),INDEX(iData,r,c)))))
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
