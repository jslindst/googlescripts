/**
 * Returns right most available value(s) from range.
 * 
 * @param {range} range The range which contains the values
 * @return Right most values for each row of the range
 * @customfunction
 */
function RIGHTMOST_VALUES(range) {
  var returnValues = [[]]; 
  for (var i = 0; i < range.length; i++) {
      for (var j = 0; j < range[i].length; j++) {
        if (range[i][j] > 0) {
          returnValues[i][0] = range[i][j];
        }
      }
  }
  return returnValues;
}

function TEST_CACHE(range) {
  if (range.map) {
    return range.map(function(y) {return TEST_CACHE(y) });
  } else {
    var value = CacheService.getDocumentCache().get("test");
    if (value == null) {
      CacheService.getDocumentCache().put("test", "WAIT");
      value = Math.random();
      CacheService.getDocumentCache().put("test", value, 5);
    } else {
      return "got cached";
    }
    return value;
  }
}

/**
 * Returns Targets-To-Improve for a specific initiative and number.
 * 
 * @param {sbuMatrixRange} sbuMatrixRange The range which contains X-matrix
 * @param {number} number The ordinal number (1, 2, 3 etc..) of the TTI for the specific initiative name 
 * @param {initiativeName} initiativeName The initiative name for which to fetch the TTI
 * @return Name of the Target-to-Improve
 * @customfunction
 */
function SBU_TTI_FOR_INITIATIVE(sbuMatrixRange, number, initiativeName) {
  var args = extractArguments(3);
  var sbuMatrixA1notation = new a1Notation(SpreadsheetApp.getActiveSheet(),args[1]);  
  var sbuMatrix = XMatrix.GET_XMATRIX_FOR(sbuMatrixA1notation.sheet());

  var initiativeNumber = sbuMatrix.initiativeNumberForName(initiativeName);
  if (initiativeNumber == -1) return "(initiative '" + initiativeName + "' not found)";  
  
  var ttis = sbuMatrix.ttisForInitiative(initiativeNumber);
  if (number > ttis.length) return "(no tti)";
  return sbuMatrix.tti(ttis[number-1]);  
}

function _XMATRIX_TITLE_TEST() {
  var sbuMatrixA1notation = new a1Notation(SpreadsheetApp.getActiveSheet(),"'Level-0 Cargotec X-Matrix'!1:21");  
  var sbuMatrix = XMatrix.GET_XMATRIX_FOR(sbuMatrixA1notation.sheet());
  Logger.log(sbuMatrix.data[sbuMatrix.xrow-1][sbuMatrix.xcolumn-1] + " " + sbuMatrix.xrow + " " + sbuMatrix.xcolumn);
}

/**
 * Returns the center title of the X-Matrix in the range.
 * @param {sbuMatrixRange} sbuMatrixRange The range which contains X-matrix
 * @return Value of the Center Cell
 * @customfunction
 */
function XMATRIX_TITLE(sbuMatrixRange) {
  var args = extractArguments(1);
  var sbuMatrixA1notation = new a1Notation(SpreadsheetApp.getActiveSheet(),args[1]);  
  var sbuMatrix = XMatrix.GET_XMATRIX_FOR(sbuMatrixA1notation.sheet());
  return sbuMatrix.title;
}

function BOWLER_SBU_TTI(sbuMatrixName, number) {
  var args = extractArguments(2, arguments);
  var sbuMatrixA1notation = new a1Notation(SpreadsheetApp.getActiveSheet(),args[1]);
  var sbuMatrix = XMatrix.GET_XMATRIX_FOR(sbuMatrixA1notation.sheet());

  var tti = sbuMatrix.ttis[number];

  var solid = sbuMatrix.solid_priorities_for_tti[tti];
  var supporting = sbuMatrix.supporting_priorities_for_tti[tti];

  var bullets = false;
  var count_priorities = 0;
  count_priorities += (solid != null ? solid.length : 0);
  count_priorities += (supporting != null ? supporting.length : 0);
  if (count_priorities > 1) bullets = true; 

  var text = "";
  if (solid != null) solid.forEach(function (val) { text += (bullets ? "* " : "") + val + "\n"; })
  if (supporting != null) supporting.forEach(function (val) { text += "(" + val + ")\n"; })
      
  return [[tti, text.trim()]];
}

function _testBOWLER_SBU_TTI() {
  var timer = new StopWatch().start();
  var values = BOWLER_SBU_TTI("'Level-0 Cargotec - BASED ON STRATEGY FORUM'!A:AD",1);
  Logger.log(values);
  assertEquals(values[0][0], "# Number of Filing Approvals\n\
Number of Consultant/Lawyer hours spent :D\n\
(What would be leading metrics?)");
  assertEquals(values[0][1], "Run Clover and ensure sufficient M&A capabilities, processes & resources to manage, close and integrate future and current active deals");
  Logger.log(timer.toString());
}


function SBU_Q4_OBJECTIVE(sbuMatrixName, number) {
  var args = extractArguments(2);
  var sbuMatrixA1notation = new a1Notation(SpreadsheetApp.getActiveSheet(),args[1]);
  var sbuMatrix = XMatrix.GET_XMATRIX_FOR(sbuMatrixA1notation.sheet());

  return sbuMatrix.objectives[number];
}

function SBU_STRATEGIC_PRIORITY(sbuMatrixRange, number) {
  var args = extractArguments(2);
  var sbuMatrixA1notation = new a1Notation(SpreadsheetApp.getActiveSheet(),args[1]);  
  var sbuMatrix = XMatrix.GET_XMATRIX_FOR(sbuMatrixA1notation.sheet());
  
  return sbuMatrix.priorities[number];
}



function SBU_STRATEGIC_INITIATIVEX(sbuMatrixRange, number, resourceName) {
  var args = extractArguments(3, arguments);
  var sbuMatrixA1notation = new a1Notation(SpreadsheetApp.getActiveSheet(),args[1]);  
  var sbuMatrix = XMatrix.GET_XMATRIX_FOR(sbuMatrixA1notation.sheet());

  var resource = sbuMatrix.resources.indexOf(resourceName);
  if (resource == -1) return "(resource '" + resourceName + "' not found)";  
    
  var initiatives = sbuMatrix.priorities_for_resource[resourceName];
  Logger.log(initiatives);
  if (number >= initiatives.length) return "(no initiative)";
  return initiatives[number];
}

function testStrategicPriorityForResource() {
  var val = SBU_STRATEGIC_INITIATIVEX("'Level-0 Cargotec - BASED ON STRATEGY FORUM'!$1:$25", 0, "SVP IM & Digi, Soili Mäkinen");
  assertEquals(val, "2021 Cargotec Strategic Improvement Priorities");
  val = SBU_STRATEGIC_INITIATIVEX("'Level-0 Cargotec - BASED ON STRATEGY FORUM'!$1:$25", 1, "SVP IM & Digi, Soili Mäkinen");
  assertEquals(val, "Run Clover and ensure sufficient M&A capabilities, processes & resources to manage, close and integrate future and current active deals");

}




function SHEET_NAME(argument) {
  var args = extractArguments(1, arguments);
  var a1n = new a1Notation(SpreadsheetApp.getActiveSheet(),args[1]);  
  return a1n.sheet().getSheetName();
}

function testExtractWithLocal() {  
  assertEquals(SHEET_NAME("'Level-0 Cargotec - BASED ON STRATEGY FORUM'!$1:$25"), "Level-0 Cargotec - BASED ON STRATEGY FORUM");
}

function testXXX() {
  var timer = new StopWatch();

  Logger.log("ACTIVE SHEET READ TEST, reading 10 times");
  timer.start();
  for (var i = 0; i < 10; i++) {
    var data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  }
  timer.stop();
  Logger.log(timer);


  var dataString = JSON.stringify(data);
  CacheService.getDocumentCache().put("data", dataString, 30);


  Logger.log("CACHE SERVICE TEST, reading 10 times");
  timer.clear();
  timer.start();
  for (var i = 0; i < 10; i++) {
     data = JSON.parse(CacheService.getDocumentCache().get("data"));
  }  
  timer.stop();
  Logger.log(timer);

  // Data length

  Logger.log("Data Length in both tests: " + dataString.length);
//  Logger.log(dataString);


  TEST_CACHE("asdf");
  var matrix = new XMatrix(SpreadsheetApp.getActiveSheet());
  CacheService.getDocumentCache().put("val", "test");
  Logger.log(CacheService.getDocumentCache().get("val"));
  

  
  var values = {
    'foo': 'bar',
    'x':'y',
    'key': 'value'
  }
  values['how'] = 'dihow!';
  Logger.log(values);
  Logger.log(values['foo']);
  Logger.log(values['how']);

  CacheService.getDocumentCache().put("map", JSON.stringify(values));
  Logger.log(CacheService.getDocumentCache().get("map"));
  Logger.log(JSON.parse(CacheService.getDocumentCache().get("map"))['how']);


  CacheService.getDocumentCache().put("matrix", JSON.stringify(matrix));
  Logger.log(JSON.stringify(matrix));
  var newX = JSON.parse(CacheService.getDocumentCache().get("matrix"));
  Logger.log(newX.xcolumn);
  Logger.log(newX.initiativesForResource(1))
  
}

function testReadingXmatrix() {
  var matrix = new XMatrix(SpreadsheetApp.getActiveSheet());
  assertEquals(matrix.objectives[0], "2021 Cargotec Objectives");
  assertEquals(matrix.objectives[1], "Running the Business & Continuity: Turnover from XX to YY, Operating Profit from XX to YY%");
  assertEquals(matrix.objectives[2], "Merger: Successful merger approval and excellent readiness for integration in 2022");
  assertEquals(matrix.objectives[3], "Inspiration: Employee Climate from XX to YY% - Direction, Pride & Hope");

  assertEquals(matrix.resources.indexOf("SVP IM & Digi, Soili Mäkinen"), 9);

}

function _testXMatrixCache() {
  var timer = new StopWatch().start();
  var matrix = XMatrix.GET_XMATRIX_FOR(SpreadsheetApp.getActiveSheet());
  Logger.log(timer.toString());
  var matrix = XMatrix.GET_XMATRIX_FOR(SpreadsheetApp.getActiveSheet());
  Logger.log(timer.toString());

}

function placeNamedXCentersOnAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var invalidSheets = [];
  spreadsheet.getSheets().forEach(function (sheet) {
    try {
      locateXCenter(sheet);
    } catch (e) {
      invalidSheets.push(sheet.getName());
    }
  });
  Logger.log("Invalid sheets: " + invalidSheets);
  return invalidSheets;
}


/**
 * Locates the XMatrixCenter first from named ranges, if not finding it, it reads it 
 * from the data and then sets the named range for future.
 */
function locateXCenter(sheet) {
  var val = locateXCenterFromNamedRange(sheet);
  if (val != null) return val;
  val = locateXCenterFromContent(sheet);

  // Store the XCENTER if possible from this scope.
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  if (authInfo.getAuthorizationStatus() == ScriptApp.AuthorizationStatus.NOT_REQUIRED) {
    var rangeToBeNamed = sheet.getRange(val[0], val[1]);
    var spreadsheet = sheet.getParent();
    spreadsheet.setNamedRange("XCENTER" + sheet.getSheetId(), rangeToBeNamed);
  } 

  return val;
}

/**
 * Returns, if exist, the XMatrixCenter from named ranges having in "XCENTER"
 */
function locateXCenterFromNamedRange(sheet) {
  var namedRanges = sheet.getNamedRanges();
  var matches = namedRanges.filter(function (named) { return named.getName().match(/.*XCENTER*/); });
  if (matches == 0) return null;
  if (matches > 2) throw new Error("The sheet " + sheet.getName() + " has several named ranges ending in 'XCENTER', could not locate center");
  //Logger.log(matches[0].getRange().getRow() + ", " + matches[0].getRange().getColumn());
  return [matches[0].getRange().getRow(), matches[0].getRange().getColumn()];
}

// TODO: To make this a bit more elaborate. Current it tries to locate "targets to improve"
/**
 * Locate XMatrix Center from sheet data.
 */
function locateXCenterFromContent(sheet) {
  this.data = sheet.getDataRange().getValues(); 
  for (var i = 4; i < data.length; i++) { // should be ok to skip first four rows
    for (var j = 4; j < data[i].length; j++) { // should be ok to skip first four columns
      var regexp = new RegExp(".*targets to improve","gi");
      if (regexp.exec(data[i][j])) {
        return [i + 1, (j - 1) + 1];
      }
    }

  }  
  throw new Error(`Could not find X-Matrix on sheet ${sheet.getSheetName()}. There was no 'Targets To Improve' item (which is used to locate the X-matrix) in the sheet.`);
} 

class XMatrix {
  
  constructor(sheet) {
    this.sbuMatrixSheet = sheet;
    Logger.log(`Reading sheet ${sheet.getSheetName()}`);

    // read the whole data matrix    
    this.data = this.sbuMatrixSheet.getDataRange().getValues();

    Logger.log (this.sbuMatrixSheet.getRange(1,1).getNote());

    // The center of X-matrix needs to identified correctly (read from properties, if exists)
    var values = locateXCenter(this.sbuMatrixSheet);
    this.xrow = values[0];
    this.xcolumn  = values[1];
    this.title = this.data[this.xrow][this.xcolumn]; 

    this.breakthroughs = [];
    this.objectives = [];
    this.priorities = [];
    this.ttis = [];
    this.resources = [];
    var i = 1;  
    var val = this.data[this.xrow-1+i][this.xcolumn-1];
    while (val != undefined && new String(val).trim().length > 0) {
      this.breakthroughs.push(new String(val).trim());
      i++;
      val = this.data[this.xrow-1+i][this.xcolumn-1];
    }
    i = 1; 
    val = this.data[this.xrow-1][this.xcolumn-1-i];
    while (val != undefined && new String(val).trim().length > 0) {
      this.objectives.push(new String(val).trim());
      i++;
      val = this.data[this.xrow-1][this.xcolumn-1-i]
    }
    i = 1; 
    val = this.data[this.xrow-1-i][this.xcolumn-1];
    while (val != undefined && new String(val).trim().length > 0) {
      this.priorities.push(new String(val).trim());
      i++;
      val = this.data[this.xrow-1-i][this.xcolumn-1]
    }

    // TTIs
    i = 1; 
    this.solid_priorities_for_tti = new Object();
    this.supporting_priorities_for_tti = new Object();
    val = this.data[this.xrow-1][this.xcolumn-1+i];
    while (val != undefined && new String(val).trim().length > 0) {
      this.ttis.push(new String(val).trim());

      var solid_matches = [];
      var supporting_matches = [];
      for (var p_index = 0; p_index < this.priorities.length; p_index++) {
        var dependency = this.data[this.xrow-1-(p_index+1)][this.xcolumn-1+i];
        if (dependency == '●') solid_matches.push(this.priorities[p_index]);
        if (dependency == '○') supporting_matches.push(this.priorities[p_index]);
      }
      this.solid_priorities_for_tti[new String(val).trim()] = solid_matches;
      this.supporting_priorities_for_tti[new String(val).trim()] = supporting_matches;
    
      i++;
      val = this.data[this.xrow-1][this.xcolumn-1+i]
      if (val == "Resourcing") break;
    }

    // RESOURCING
    this.priorities_for_resource = new Object();
    val = this.data[this.xrow-1][this.xcolumn-1+i];
    while (val != undefined && new String(val).trim().length > 0) {
      this.resources.push(new String(val).trim());      

      var solid_matches = [];
      for (var p_index = 0; p_index < this.priorities.length; p_index++) {
        if (p_index == 0) solid_matches.push(this.priorities[p_index]);
        var dependency = this.data[this.xrow-1-(p_index+1)][this.xcolumn-1+i];
        if (dependency == '●' || dependency == '○') solid_matches.push(this.priorities[p_index]);
      }
      this.priorities_for_resource[new String(val).trim()] = solid_matches;

      i++;
      val = this.data[this.xrow-1][this.xcolumn-1+i]
    }

/*    Logger.log(this.breakthroughs);
    Logger.log(this.objectives);
    Logger.log(this.priorities);
    Logger.log(this.ttis);
    Logger.log(this.resources);
*/
    Logger.log(this.priorities_for_resource);
    
    this.ximprovements_start = this.xrow - 1;
    
    this.xtti_start = this.xcolumn + 1;
  }
  
  static GET_XMATRIX_FOR(sheet) {
    Logger.log("Getting XMatrix for sheet " + sheet.getSheetName() + " id: " + sheet.getSheetId());
    var cached = CacheService.getDocumentCache().get("xmatrix" + sheet.getSheetId());
    if (cached != null) {
      Logger.log(`Found matrix in cache.`);
      return JSON.parse(cached);
    }
    var xmatrix = new XMatrix(sheet);
    CacheService.getDocumentCache().put("xmatrix" + sheet.getSheetId(), JSON.stringify(xmatrix),20);
    return xmatrix;
  }


  ttisForInitiative(initiativeNumber) {
    var ttis = [];

    for (var i = 1; i < 10; i++) {      
      var value = this.sbuMatrixSheet.getDataRange().getCell(this.ximprovements_start - initiativeNumber, this.xtti_start + i).getValue();
      if (value == '●' || value == '○') {
        ttis.push(i);
      }
    }
    return ttis;    
  }
  
  initiativeNumberForName(name) {
    
    var resources = this.sbuMatrixSheet.getRange(this.ximprovements_start-7, this.xcolumn, 8, 1).getValues();
    resources = resources.map(function (x) { return x[0]; })
    resources = resources.reverse();
//    return resources.toString()
    var index = resources.indexOf(name);
    if (index == -1) return -1;
    return index;    
  }
      
}
