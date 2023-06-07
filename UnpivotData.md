# Unpivot Data Function
This can be used to bring a two dimensional set of data into a flat file (i.e. a load file).


## Excel Without Lambda
`````
=LET(dataValues,D4:G7,rAxis,B4:C7,cAxis,D2:G3,
           countCaxis,ROWS(cAxis),CountRAxis,COLUMNS(rAxis),totalDims,CountRAxis+countCaxis,totNewCols,totalDims+1,
           zseq,SEQUENCE(ROWS(dataValues)*COLUMNS(dataValues),totalDims+1,0,1),
           colCount,COLUMNS(cAxis),
           modResult,MOD(zseq,totNewCols),
           c,MOD(INT(zseq/totNewCols),colCount)+1,
           r,INT(zseq/(totNewCols*colCount))+1,
                   IF(modResult<CountRAxis,
                       INDEX(rAxis,r,modResult+1),
                       IF(modResult<totalDims,INDEX(cAxis,modResult-CountRAxis+1,c),
                       INDEX(dataValues,r,c))))
`````
## Excel With Lambda
`````
=LAMBDA(dataRng,rowAxis,colAxis,
      LET(iCol,COLUMN(INDEX(rowAxis,1,1)),   amountCol,TOCOL(dataRng),  totalCells,COUNTA(amountCol),
          HSTACK(
              INDEX(rowAxis,
                     INT(SEQUENCE(totalCells,1,0,1)/COLUMNS(dataRng))+1,
                     BYCOL(INDEX(rowAxis,1,),  LAMBDA(aCol,COLUMN(aCol) -iCol +1))),
              INDEX(colAxis,
                      SEQUENCE(1,ROWS(colAxis),1,1),
                      MOD(SEQUENCE(totalCells,1,0,1),COLUMNS(dataRng))+1),
               amountCol
                      )))(D4:G7,B4:C7,D2:G3)
`````
