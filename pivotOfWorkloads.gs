function clearPivot(){
  try{
  svodSheet.getRange(7, 1, svodSheet.getLastRow()-6, svodSheet.getLastColumn()).clear();
  svodSheet.getRange(5, 11, 1, svodSheet.getLastColumn()-10).clear();
  }catch(e){}
}

function start(){ // Обход по всем листам главспецов и сбор данных в сводную таблицу
  clearPivot();
  main.getSheets()
  .map( sheet => sheet.getName())
  .filter(sheetName => !!~sheetName.indexOf("Главспец"))
  .forEach( sheetName => fromSpec(sheetName))
}

function fromSpec(sheetName){

  const specSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  let objects = getObjects(specSheet, objectColumn, specSheet.getLastRow()-5, 1)
                .map(arr => [arr[0], arr[1], 
  specSheet.getRange((arr[1]+1), 3, specSheet.getLastRow()-(arr[1]+2), 1).getValues()
           .map(arr => arr[0]).indexOf("")]);

  objects.forEach( object => {    

    const info = specSheet.getRange((object[1]+1), objectColumn, object[2], 13).getValues()
                         .map( (arr, i) => [object[0], arr[4], arr[5], arr[6], arr[7], arr[8], arr[9], arr[11], specSheet.getName(), arr[12], (object[1]+1+i)]);

    info.map(e=>e[1])
        .filter((e, position, array) => {return array.lastIndexOf(e) === position})
        .map(stage => 
        {  
          let a = fillArray(specSheet.getLastColumn()-15, 4, 0);
          let tempArr = [0, 0, 0, new Date(2025, 0, 1), new Date(1970, 0, 1), 0, 0, 0, 0].concat(a);
          let days = 0;
          let rows = info.filter(arr => arr[1] == stage);
          rows.forEach(item => 
          {
              tempArr[0] = item[0];
              tempArr[1] = item[1];
              tempArr[2] = item[2];
              tempArr[3] > item[3] ? tempArr[3] = item[3] : tempArr[3];
              tempArr[4] < item[4] ? tempArr[4] = item[4] : tempArr[4];
              days += item[5];
              tempArr[5] += item[6]; // ТРЗ
              tempArr[7] = item[8];
              tempArr[8] += item[9];

              specSheet.getRange(item[10], startDateLine, 1, specSheet.getLastColumn()-15).getValues()[0]
              .forEach((jtem, j) => jtem != "" || jtem != 0 ? tempArr[j+9] += jtem : tempArr[j+9] += 0)
          });
          tempArr[6] = (tempArr[5]/((days * 8)/100))/100;
          tempArr[8] /= rows.length;
          tempArr = tempArr.map( num => {return num == 0 ? "" : num})
          svodSheet.getRange(svodSheet.getLastRow()+1, 2, 1, tempArr.length).setValues([tempArr]);
        })    
  })  

  let workload = svodSheet.getRange(5, 11, 1, svodSheet.getLastColumn()-10).getValues()[0]
                          .map(num => {return num ? num : 0});
  findLastNumber(totalAssignedEmployes(specSheet)).forEach((num, index) => workload[index] += num);
  svodSheet.getRange(5, 11, 1, svodSheet.getLastColumn()-10).setValues([workload.map( num => {return num == 0 ? "" : num})]);
}

function totalAssignedEmployes(specSheet){

  let row = specSheet.getRange("O:O").getValues()
                     .map((cell, index) => [cell[0], index+1])
                     .filter( obj => obj[0] == "Общее число спец.")[0][1];
  
  return specSheet.getRange(row, 16, 1, specSheet.getLastColumn()-15).getValues()[0].map(num => {return num == 0 ? "" : num});

}

function findLastNumber(array){
  for(let a = array.length; a>0 ; a--)
  {
    if(array[a])
    {
      return array.slice(0, a+1);      
    }
  }
}
