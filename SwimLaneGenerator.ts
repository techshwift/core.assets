// Data format TaskID	Stage	Task	Type	Input	Output	OwnerPersona	DependsOn, end with ###END_OF_DATA###
interface DATARANGE {
  startIndex: Number;
  endIndex: Number;
}

function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("Sheet1");
  const datarange: DATARANGE = detectStartAndFinish(sheet)
  const personaColumnIndex = 6
  const uniquePersonas = getUniquePersonas(sheet, personaColumnIndex, datarange)
  createSwimlaneDiagram(sheet, datarange, uniquePersonas)
}

function detectStartAndFinish(sheet: ExcelScript.Worksheet) {
  var startString = "TaskID";
  const endString = "###END_OF_DATA###";

  var cell = sheet.getRange("A:A").find(startString, {
    completeMatch: true,
    matchCase: false,
    searchDirection: ExcelScript.SearchDirection.forward
  });
  var startIndex = cell.getRowIndex();
  cell = sheet.getRange("A:A").find(endString, {
    completeMatch: true,
    matchCase: false,
    searchDirection: ExcelScript.SearchDirection.forward
  });
  var endIndex = cell.getRowIndex();
  return {startIndex, endIndex}
}

function getUniquePersonas(sheet: ExcelScript.Worksheet, colIndex: number, datarange: DATARANGE) {
  const personasRangeStart = +datarange.startIndex + 1
  const personasRangeEnd = +datarange.endIndex - personasRangeStart
  const personasRange = sheet.getRangeByIndexes(personasRangeStart, colIndex, personasRangeEnd, 1)
  const personas = personasRange.getValues()
  var uniquePersonas: Array<string> = []
  personas.forEach(persona => {
    if(! uniquePersonas.includes(persona.toString())) uniquePersonas.push(persona.toString())
  })
  return uniquePersonas
}


function createSwimlaneDiagram(sheet: ExcelScript.Worksheet, datarange: DATARANGE, uniquePersonas: Array<string>) {

  // Cleanup shapes
  let shapes = sheet.getShapes()
  shapes.forEach(shape => {
    shape.delete();
  });

  // Master Data
  const colorPalette = {
    "Planning": "#C04F15", // Orange
    "Prepare Target": "#223861", // Dark Teal
    "Prepare Source": "#6d7178", // Gray
    "Migrate": "#c00000", // Red
    "Cutover": "#43186e", // Dark purple
    "Closure": "#508021" // Green
  }
  


  sheet.setShowGridlines(false);
  const swimLaneBeginRow = +datarange.endIndex + 4
  const swimLaneEndRow = swimLaneBeginRow + uniquePersonas.length
  const swimLaneRange = sheet.getRangeByIndexes(swimLaneBeginRow, 0, uniquePersonas.length, 10)
  swimLaneRange.getFormat().setRowHeight(75);

  drawLanes(swimLaneRange)
  let i = 0;
  let personaLaneMap = {}
  let lastTopLeftMap = {}
  for(i = swimLaneBeginRow; i < swimLaneEndRow; i++) {
    personaLaneMap[uniquePersonas[i - swimLaneBeginRow]] = i
    let firstShapeCell = sheet.getCell(i,1)
    lastTopLeftMap[i] = {
      "top": firstShapeCell.getTop(),
      "left": firstShapeCell.getLeft()
    }
    sheet.getCell(i, 0).setValue(uniquePersonas[i - swimLaneBeginRow])
    sheet.getCell(i, 0).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)
    sheet.getCell(i, 0).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center)
    sheet.getCell(i, 0).getFormat().setWrapText(false)
    sheet.getCell(i, 0).getFormat().getFont().setBold(true)
  }

  let row = +datarange.startIndex + 1
  let shapesMap = {}
  for (row = +datarange.startIndex + 1; row < datarange.endIndex; row++) {
    let taskId = sheet.getCell(row, 0).getValue()
    let stage = sheet.getCell(row, 1).getValue()
    let task = sheet.getCell(row, 2).getValue()
    let taskType = sheet.getCell(row, 3).getValue()
    let persona = sheet.getCell(row, 6).getValue()
    let dependsOn = sheet.getCell(row, 7).getValue()
    
    let shapeType = ExcelScript.GeometricShapeType.rectangle
    let taskLaneIndex: number = personaLaneMap[persona.toString()]
    switch(taskType) {
      case "Task": {
        shapeType = ExcelScript.GeometricShapeType.rectangle
        break
      }
      case "Decision": {
        shapeType = ExcelScript.GeometricShapeType.diamond
        break
      }
      default : {
        shapeType = ExcelScript.GeometricShapeType.rectangle
        break;
      }
    }
    let width = 150
    let height = 50
    let top = +lastTopLeftMap[taskLaneIndex]["top"]
    let left = +lastTopLeftMap[taskLaneIndex]["left"] + 15

    let shape = sheet.addGeometricShape(shapeType);
    shape.setLeft(left);
    shape.setTop(top);
    shape.setWidth(width);
    lastTopLeftMap[taskLaneIndex]["left"] += 175
    if(shapeType == ExcelScript.GeometricShapeType.diamond) {
      shape.setHeight(height * 1.5)
      shape.setTop(top)
    }
    else {
      shape.setHeight(height);
      shape.setTop(top+15)
    }
    shape.getTextFrame().getTextRange().setText(task.toString());
    shape.getTextFrame().getTextRange().getFont().setSize(12)
    shape.getTextFrame().setHorizontalAlignment(ExcelScript.ShapeTextHorizontalAlignment.center);
    shape.getTextFrame().setVerticalAlignment(ExcelScript.ShapeTextVerticalAlignment.middle);
    
    shape.getFill().setSolidColor(colorPalette[stage.toString()])

    shapesMap[taskId.toString()] = shape

    
    //Add Links
    let dependsOnString = dependsOn.toString()
    if (isNumber(dependsOnString)) {
      if (shapesMap[dependsOnString]) {
        let fromShape: ExcelScript.Shape = shapesMap[dependsOnString]        
        connectShapes(sheet, fromShape, 3, shape, 1, "")
      }
    }
    if (isDecisionString(dependsOnString)){
      let arr = dependsOnString.split(':')
      let fromShape:ExcelScript.Shape = shapesMap[arr[0]]
      let text = arr[1]
      connectShapes(sheet, fromShape, 3, shape, 1, text)
    }
    if (isList(dependsOnString)) {
      let connections = dependsOnString.split(',')
      parseDependencies(sheet, connections, shapesMap, shape)
      // connections.forEach(conn => {
      //   if(isNumber(conn)) {
      //     if(shapesMap[conn]) {
      //       let fromShape: ExcelScript.Shape = shapesMap[conn]
      //       connectShapes(sheet, fromShape, 3, shape, 1, "")
      //     }
      //   }
      //   if(isDecisionString(conn)) {
      //     let arr = dependsOnString.split(':')
      //     let fromShape: ExcelScript.Shape = shapesMap[arr[0]]
      //     let text = arr[1]
      //     connectShapes(sheet, fromShape, 3, shape, 1, text)
      //   }
      // });
    }
    
  }
}

function drawLanes(range: ExcelScript.Range) {
  let rowCount = range.getRowCount()
  let i:number = 0
  for(i=0; i < rowCount; i++) {
    let currentRange = range.getRow(i).getEntireRow()
    let edgeTop = currentRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop);
    edgeTop.setStyle(ExcelScript.BorderLineStyle.continuous);
    edgeTop.setWeight(ExcelScript.BorderWeight.thin);
    let edgeBottom = currentRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
    edgeBottom.setStyle(ExcelScript.BorderLineStyle.continuous);
    edgeBottom.setWeight(ExcelScript.BorderWeight.thin);
  }
  let firstColumn = range.getColumn(0)
  let edgeRight = firstColumn.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight)
  edgeRight.setStyle(ExcelScript.BorderLineStyle.continuous)
  edgeRight.setWeight(ExcelScript.BorderWeight.thin);
  
  firstColumn.getFormat().setColumnWidth(150);
}

function isNumber(str: string): boolean {
  let re = new RegExp('^\\d+$')
  return re.test(str)
}

function isDecisionString(str: string): boolean {
  let re = new RegExp('^\\d+:\\w+$')
  return re.test(str)
}

function isList(str: string): boolean {
  let re = new RegExp(',')
  return re.test(str)
}

function connectShapes(sheet:ExcelScript.Worksheet, fromShape: ExcelScript.Shape, fromSite: number, toShape: ExcelScript.Shape, toSite: number, text: string) {
  let arrow = sheet.addLine(0, 0, 10, 10, ExcelScript.ConnectorType.elbow)
  arrow.getLine().setEndArrowheadStyle(ExcelScript.ArrowheadStyle.open)
  arrow.getLine().connectBeginShape(fromShape, fromSite)
  arrow.getLine().connectEndShape(toShape, toSite)
  if(text != "") {
    let txtTop = +arrow.getLine().getShape().getTop() + 5
    let txtLeft = arrow.getLine().getShape().getLeft() + 15
    console.log(location)
    let txtBox = sheet.addTextBox(text)
    txtBox.setTop(txtTop)
    txtBox.setLeft(txtLeft)
  }
}

function parseDependencies(sheet: ExcelScript.Worksheet, connections: Array<string>, shapesMap:{}, shape:ExcelScript.Shape){
     connections.forEach(conn => {
        if(isNumber(conn)) {
          if(shapesMap[conn]) {
            let fromShape: ExcelScript.Shape = shapesMap[conn]
            connectShapes(sheet, fromShape, 3, shape, 1, "")
          }
        }
        if(isDecisionString(conn)) {
          let arr = conn.split(':')
          let fromShape: ExcelScript.Shape = shapesMap[arr[0]]
          let text = arr[1]
          connectShapes(sheet, fromShape, 3, shape, 1, text)
        }
      });
}
