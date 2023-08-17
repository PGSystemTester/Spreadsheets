# Misc Excel Routines For Typescript

## getRangeIntersect
Similar to VBA `intersect` function. Below returns intersections of two ranges as new range.

 - Returns a range or `null` if no intersection
 - Cannot do more than two ranges (just repeat function)

```*.js
//	const ws = workbook.getActiveWorksheet();
//	let newRng: ExcelScript.Range = getRangeIntersect(ws.getRange("A12:G19"),ws.getRange("A:C"));
function getRangeIntersect(aRng: ExcelScript.Range, bRng: ExcelScript.Range): ExcelScript.Range {
    const noIntersectionValue: null = null;
    //const noIntersectionValue:string = "no intersection";
    if (aRng.getWorksheet().getName() != bRng.getWorksheet().getName()) {
        return noIntersectionValue;
    }
    const ws = aRng.getWorksheet();

    const startRow = Math.max(aRng.getCell(0, 0).getRowIndex(), bRng.getCell(0, 0).getRowIndex());
    const endRow = Math.min(aRng.getLastCell().getRowIndex(), bRng.getLastCell().getRowIndex());
    if (startRow > endRow) {
        return noIntersectionValue;
    }

    const startCol = Math.max(aRng.getCell(0, 0).getColumnIndex(), bRng.getCell(0, 0).getColumnIndex());
    const endCol = Math.min(aRng.getLastCell().getColumnIndex(), bRng.getLastCell().getColumnIndex());

    if (startCol > endCol) {
        return noIntersectionValue;
    }

    const newRng: ExcelScript.Range = ws.getRangeByIndexes(startRow, startCol,
        endRow - startRow + 1, endCol - startCol + 1);

    return newRng;
}
```

## getLastUsedRow
Similar to this [infamous Stackoverflow debate](https://stackoverflow.com/questions/11169445/find-last-used-cell-in-excel-vba/59081657#59081657). Defines `""` as empty row vs. `isblank()`. Leverages above **getRangeIntersect()** function. Can be scoped to a single range or be applied to entire worksheet using `true` in second variable.

 - Returns a number
 - First parameter must be a range
 - Second parameter is optional, if true, applies to entires sheet that first parameter range exists on

```*.js
function getLastUsedRow(aRng: ExcelScript.Range,doEntireSheet?:boolean){
	const ws:ExcelScript.Worksheet = aRng.getWorksheet();
	const zUsedRng: ExcelScript.Range = aRng.getWorksheet().getUsedRange();
	const noValues = 'No Values';//value if no values
	if(zUsedRng!==undefined){
		let tangoRng: ExcelScript.Range;
		if(doEntireSheet){
			tangoRng = zUsedRng;
		}else{
			tangoRng = getRangeIntersect(zUsedRng,aRng);
		}

		let allValues: (string | number | boolean)[][];
		allValues = tangoRng.getValues();

		for(let r=allValues.length-1; r>=0; r--){
			for(let c=0; c<allValues[0].length;c++){
				let zSingleValue = allValues[r][c];
				if (zSingleValue!=''){
					let tangoCell:ExcelScript.Range = tangoRng.getCell(r,c);
					let finalValue:number = tangoCell.getRowIndex()+1;
					return finalValue;
				}
			}
		}
	}
	return noValues;
}
```


