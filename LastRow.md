## Function To Get Last Row In Excel
Still a work in progress, testing a couple methods to get last row as quickly as possible.

```
=LET(xcol,A:A,tText,"Θ",tNumb,1E+303,iLastRow,LAMBDA(iType,xRng,[currentLeader],IF(ISOMITTED(currentLeader),IFERROR(MATCH(iType,xRng),0),
LET(testRange,DROP(xRng,currentLeader),IF(COUNTA(testRange)=0,currentLeader,IFERROR(MATCH(iType,testRange),0)+currentLeader)))),
textResult,iLastRow(tText,xcol,COUNTA(xcol)),numResult,iLastRow(tNumb,xcol,textResult),
remainingRange,DROP(xcol,numResult),iFinal,IF(COUNTA(remainingRange)=0,numResult,MATCH(1,1/(remainingRange<>""))+numResult),iFinal)
````
