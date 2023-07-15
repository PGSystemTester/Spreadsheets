# Unpivot Data Function
This can be used to bring a two dimensional set of data into a flat file (i.e. a load file).
[Example File](https://onedrive.live.com/edit.aspx?resid=7D665EED73FFBB23!75536)


## Excel Without Lambda
`````
=LET(dataValues,$F$10:$J$14,AxisRows,$B$10:$D$14,  AxisColumns,$F$7:$J$8,
      DimsInColumnAxis, ROWS(AxisColumns),
      DimsRowAxis, COLUMNS(AxisRows),
      totalDims,   DimsRowAxis+DimsInColumnAxis,
      totNewCols,  DimsRowAxis+DimsInColumnAxis+1,
      rowCount,    ROWS(dataValues)*COLUMNS(dataValues),
      zSeq,        SEQUENCE(rowCount,totNewCols,0),
      colCount,    COLUMNS(AxisColumns),
      zDiv,        totNewCols*colCount,
      zAddr,      DimsRowAxis-1,
      modResult,   MOD(zSeq,totNewCols),
      c,           MOD(INT(zSeq/totNewCols),colCount)+1,
      r,           INT(zSeq/zDiv)+1,
      zFinalResult,IF(modResult<DimsRowAxis,
                      INDEX(AxisRows,r,modResult+1),
                   IF(modResult<totalDims,
                      INDEX(AxisColumns,modResult-zAddr,c),

                      INDEX(dataValues,r,c))),zFinalResult)
`````
## Excel With Lambda
`````
=LET(dataRng,F10:J14, rowAxis, B10:D14, colAxis,F7:J8,
          amountCol,   TOCOL(dataRng),
          iCol,        COLUMN(TAKE(rowAxis,,1)),
          colCount,    COLUMNS(colAxis),
          amountCount, ROWS(amountCol),
          rAxis,  INDEX(rowAxis,
                      INT(SEQUENCE(amountCount,1,0,1)/colCount)+1,
                      BYCOL(TAKE(rowAxis,1),
                          LAMBDA(eachColumn,COLUMN(eachColumn)-iCol+1))),
          yAxis,  INDEX(colAxis,
                       SEQUENCE(1,ROWS(colAxis)),
                       MOD(SEQUENCE(amountCount,1,0),colCount)+1),
          HSTACK(rAxis,yAxis,amountCol))
`````

## Example

### Before

|  |  |  |  | Jan | Feb | Mar | Apr | May |
|---|---|---|---|---|---|---|---|---|
|  |  |  |  | Actual | Actual | Actual | Budget | Budget |
|  |  |  |  |  |  |  |  |  |
| Disney | Tickets | Revenue |  | 507 | 607 | 707 | 807 | 907 |
| Disney | Movies | Costs |  | 508 | 608 | 708 | 808 | 908 |
| LucasFilm | Promo | Revenue |  | 509 | 609 | 709 | 809 | 909 |
| LucasFilm | Vader | Taxes |  | 510 | 610 | 710 | 810 | 910 |
| Marvel | HR | Costs |  | 511 | 611 | 711 | 811 | 911 |

### After


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
