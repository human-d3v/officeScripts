function main(wb: ExcelScript.Workbook) {
  //define variables to use in later functions
  var ws:ExcelScript.Worksheet = wb.getActiveWorksheet();
  const lRow:number = (ws.getUsedRange().getLastRow().getRowIndex()+1) //offset 
  let arr: object[] = dateArrBuilder(ws.getRange(`C2:C${lRow + 1}`));
  let lMonth:number = getMonthFromStr(arr,'max'); //returns the index of the month starting from zero
  const fMonth: number = getMonthFromStr(arr,'min');
  resolveDate(ws); //creates the resolved date column in ws
  createSheets(fMonth,lMonth);

  
  //this is the master function that creates the individual sheets based on the number of months in the range. It then creates a caldendar adds a function to pull the data from column B of the raw data sheet.

  function createSheets(startMonth: number, endMonth: number) {
		let stopNumber:number[] = [];
		let startNumber:number[] = [];
		let years:number[]= [];
		if(startMonth>endMonth){
			stopNumber = [11,endMonth];
			startNumber = [startMonth, 0];
			years = [new Date().getFullYear(),new Date().getFullYear() + 1];
		}else{
			stopNumber = [endMonth];
			startNumber = [startMonth];
			years = [new Date().getFullYear()];
		}
		for(let j:number = 0; j<startNumber.length; j++ ){
			for (let i: number = startNumber[j]; i <= stopNumber[j]; i++) {
    	  let name: string = getMonthName(i);
    	  wb.addWorksheet(name);
    	  let wksht:ExcelScript.Worksheet = wb.getWorksheet(name);
    	  wksht.getRange("B2").setValue(name); //allows us to use the Name Value in CalendarDay
    	  makeCalendar(i,wksht,years[j]);
    	};
		};
	};
    
  function dateArrBuilder(rng: ExcelScript.Range):object[]{
      const arr:object[] = [];
      const stopRow:number = rng.getLastRow().getRowIndex();
      for(let i:number = 2;i<=stopRow;i++){
          let nRng:ExcelScript.Range = ws.getRange(`C${i}`);
          arr.push(new Date(nRng.getText()));
      };
    return arr;
  };

  function getMonthFromStr(arr:object[],minOrMax:string):number{
    if (minOrMax.toLowerCase() === 'min'){
      let m = new Date(Math.min.apply(null,arr));
      return m.getMonth();
    }
    else if (minOrMax.toLowerCase() === 'max'){
      let m = new Date(Math.max.apply(null,arr));
      return m.getMonth();
    }
    else {
      throw new Error(`${minOrMax} is not accepted syntax. Please enter 'min' or 'max'.`);
    };
  };

  function getMonthName(n:number):string{
    const months:string[] = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return months[n];
  };

  function getLastDay(monthNumber:number,year:number):number{
    let arr:number[] = [31,28,31,30,31,30,31,31,30,31,30,31];
    if (year % 4 === 0){
      arr[1] = 29;
    };
    return arr[monthNumber];
  };

  function weekDayHeader(wksht:ExcelScript.Worksheet){
    let a:string[] = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    let c:number = 2;
    for (let s of a){
      wksht.getCell(2,c).setValue(s);
      c += 2;
    };
  }

  function getFirstWeekday(month:number, year:number):number{
    return new Date(`${month+1}/01/${year}`).getDay();
  };

  /* this function creates the grid(s) for calendar and accepts the month number to use in the getLastDay function.*/
    function makeCalendar(month:number,wksht:ExcelScript.Worksheet,year:number){ 
      let rNum:number = 3;
      let cNum:number = (getFirstWeekday(month,year)*2)+1; //(x*2)+1 -> the offset to apply the correct day to the correct column. 
      const lastD:number = getLastDay(month,year);
      weekDayHeader(wksht); //adds weekday headers
      for (let d:number = 1;d<=lastD;d++){
        let cell:ExcelScript.Range = wksht.getCell(rNum,cNum);
        allBorders(cell);
        cell.setValue(d);
        let dRng:ExcelScript.Range = wksht.getRangeByIndexes(rNum,cNum,10,2);
        allBorders(dRng); //creates the cell for the calendar
        d === lastD ? lastCalendarDay(rNum,cNum,wksht,month,year) : calendarDay(rNum,cNum,wksht,year);
        if((cNum + 2)>14){
          cNum = 1;
          rNum += 10;
        }
        else{
          cNum +=2;
        };
      };
    };


  //accepts a range (one cell or multiple) and places a medium border around it in black. Used in makeCalendar to facilitate readability.
    function allBorders(rng:ExcelScript.Range){
      let format = rng.getFormat();
        format.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.medium);
        format.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.medium);
        format.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.medium);
        format.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.medium);
    };

  //receives a cell reference from createSheets in the loop, offsets those cell references, and adds a formula that captures the name of the original sheet, and what day of the month it is. 
    function calendarDay(r:number,c:number,wsName:ExcelScript.Worksheet,year:number){
      let name: string = ws.getName();
      wsName.getCell((r + 1), (c + 1)).setFormulaR1C1(`=FILTER('${name}'!C2, (('${name}'!C12>=DATEVALUE(R[-1]C[-1]&" "&R2C2&" ${year}"))*('${name}'!C12<DATEVALUE((R[-1]C[-1]+1)&" "&R2C2&" ${year}"))),"")`);
    };

    function lastCalendarDay(r: number, c: number, wsName: ExcelScript.Worksheet,month:number,year:number) {
      let name: string = ws.getName();
      let mName: string = getMonthName(month!==11?month+1:0)
      wsName.getCell((r + 1), (c + 1)).setFormulaR1C1(`=FILTER('${name}'!C2, (('${name}'!C12>=DATEVALUE(R[-1]C[-1]&" "&R2C2&" ${year}"))*('${name}'!C12<DATEVALUE("1 ${mName} ${month==11?year+1:year}"))),"")`);
    };

  //this funtion creates an array of arrays, each subarray returned has a relative copy of provided function. This is used to interact with the .setFormulasR1C1() method. 
    function formArrBuilder(formula:string,num:number):string[][]{
      let arr:string[][]=[];
      for (let i = 0;i<=num;i++){
        arr[i] = [];
        for(let j = 0;j<1;j++){
          arr[i][j] = formula;
        };
      };
      return arr;
    };

    function resolveDate(wksht:ExcelScript.Worksheet){
      let lCol:number = wksht.getUsedRange().getLastColumn().getColumnIndex() +1; //offset to create a new lastColumn
      let rng:ExcelScript.Range = wksht.getRangeByIndexes(1,lCol,lRow-1,1);
      rng.setFormulasR1C1(formArrBuilder(`=R[0]C3 + (NUMBERVALUE(LEFT(RC[-1],LEN(RC[-1])-1))*IFS(RIGHT(RC[-1],1)="d",1,RIGHT(RC[-1],1)="w",7,RIGHT(RC[-1],1)="m",30))`,lRow - 2));
    };
};
