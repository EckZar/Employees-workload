function onOpen(){
  SpreadsheetApp.getUi().createMenu("__MENU__")
    .addItem("Собрать для текущего Главспеца", "unite")
    .addItem("Собрать свод", "fromSpec")
  .addToUi();
}

function unite(){

  let name = main.getActiveSheet().getName();

  if(!!~name.indexOf("Главспец"))
  {
    countForTasks(name);
    countForObjects(name);
    countForAll(name);
    countTotalWorkers(name);
  } else {
    throw Error("Скрипт можно запускать только из листов Главспецов")
  }



}

function countForTasks(name){ // Подсчет для задач  

  let specSheet = main.getSheetByName(name);

  let datesLine = specSheet.getRange(3, 16, 1, specSheet.getLastColumn()-15).getValues()[0].map(date => new Date(date)).map(date => date.getDate() + "." + date.getMonth() + "." + date.getFullYear());
  let firstDate = specSheet.getRange(3, 16).getValue();


  let dates = getObjects(specSheet, startDateCol, specSheet.getLastRow()-5, 2);
 
  dates.forEach(date => 
  {
    let dateStart = new Date(date[0])
    let posOne, dayStartNum;
    dateStart = dateStart.getDate() + "." + dateStart.getMonth() + "." + dateStart.getFullYear();
    
    let dateEnd = new Date(date[1])
    dateEnd = dateEnd.getDate() + "." + dateEnd.getMonth() + "." + dateEnd.getFullYear();
  
    !!~datesLine.indexOf(dateStart) ? (
                                        posOne = datesLine.indexOf(dateStart), 
                                        dayStartNum = new Date(date[0]).getDay()
                                      ) : (
                                        posOne = 0, 
                                        dayStartNum = new Date(firstDate).getDay()
                                      );
    
   
    let posTwo = !!~datesLine.indexOf(dateEnd) ? datesLine.indexOf(dateEnd) : specSheet.getLastColumn()-15;

    let arr = fillArray(((posTwo-posOne)+1), dayStartNum);

    specSheet.getRange(date[2], 16 + posOne, 1, arr.length).setValues([arr]);
  })
  
}



function countForObjects(name){ // Подсчет для объектов

  let specSheet = main.getSheetByName(name);
  let objects = getObjects(specSheet, objectColumn, specSheet.getLastRow()-5, 1)
                .map(arr => [arr[1], specSheet.getRange((arr[1]+1), 3, specSheet.getLastRow()-(arr[1]+2), 1).getValues().map(arr => arr[0]).indexOf("")]);

  objects.forEach( object => {
    let arr = specSheet.getRange(3, startDateLine, 1, specSheet.getLastColumn()-15).getValues()[0].fill(0);
    
    let range = specSheet.getRange((object[0]+1), startDateLine, object[1], specSheet.getLastColumn()-15).getValues();
    
    range.forEach((item, i) => {
      item.forEach((jtem, j)=>{
        jtem == 1 ? arr[j] += jtem : arr[j] += 0;
      })
    })

    specSheet.getRange(object[0], startDateLine, 1, specSheet.getLastColumn()-15).setValues([arr]);
  })  

}

function countForAll(name){ // Общая сумма всех занятых

  let specSheet = main.getSheetByName(name);
  let objects = getObjects(specSheet, objectColumn, specSheet.getLastRow()-5, 1);  
  let arr = specSheet.getRange(4, startDateLine, 1, specSheet.getLastColumn()-15).getValues()[0].fill(0);    
  
  objects.forEach(object => {
    let range = specSheet.getRange(object[1], startDateLine, 1, specSheet.getLastColumn()-15).getValues();
    range[0].forEach((item, i)=>{
      item != 0 ? arr[i] += item : arr[i] += 0;
    })
  })

  specSheet.getRange(4, startDateLine, 1, specSheet.getLastColumn()-15).setValues([arr]);
}

function countTotalWorkers(name){ // Подсчет только тех тех задач, на которые назначен специалист

  let specSheet = main.getSheetByName(name);

  let commonLine = fillArray(specSheet.getLastColumn()-15, 4, 0);

  let lastRow = specSheet.getRange("C6:C").getValues()
                .map(function(arr, i)
                { 
                  return arr.concat(i+6);
                })
                .filter(function(e)
                {
                  return e[0] != "";
                })
                .pop()[1];

  let workers = getObjects(specSheet, startDateCol, lastRow, 6);
  getObjects(specSheet, namesCol, lastRow, 1)
    .map(arr => arr[0])
    .filter(function (e, position, array) 
    {
        return array.lastIndexOf(e) === position && e != "" && e != "Гл.спец" && e != "Ведущий" && e != "Категория"; // вернём уникальные элементы и отсеим строки без специалиста
    })
    .map(arr => JSON.stringify({
                  "name":arr, 
                  "dates":workers.filter(e => e[5]===arr).map(arr => [arr[0], arr[1]])
                }))
    .forEach(function(item, i){

      let tempArr = fillArray(specSheet.getLastColumn()-15, 4, 0);   

      JSON.parse(item).dates.forEach(function(jtem, j){
          
        let dateStart = new Date(jtem[0])
        let dayStartNum = dateStart.getDay();
        dateStart = dateStart.getDate() + "." + dateStart.getMonth() + "." + dateStart.getFullYear();
        
        let dateEnd = new Date(jtem[1])
        dateEnd = dateEnd.getDate() + "." + dateEnd.getMonth() + "." + dateEnd.getFullYear();

        let posOne = datesLine.indexOf(dateStart)+1;
        let posTwo = datesLine.indexOf(dateEnd)+1;

        let arr = fillArray(((posTwo-posOne)+1), dayStartNum);
        
        arr.forEach((qtem, q) => 
        {
          qtem != "" ? tempArr[q + (posOne-1)] = 1 : qtem = "";          
        })
      })

      for(let q = 0; q<commonLine.length; q++)
      {
        commonLine[q] += tempArr[q];
      }
    })
                      

  specSheet.getRange(specSheet.getLastRow()-2, startDateLine, 1, commonLine.length).setValues([commonLine]);

}
