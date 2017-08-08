
/*
 * GET home page.
 */

var map={};
var excelMaps={};
var Excel=require('exceljs');


/**
 * Home Page
 */
exports.index = function(req, res){
  res.render('index', { title: 'Excel Automations' });
};


/**
 * Update File Page 
 */
exports.updateExisting = function(req, res){
	  res.render('updateExistingFile', { title: 'Excel Automations' });
	};
	
/**
 * Create new File page 
 */	
exports.createNewFile = function(req, res){
		  res.render('createNewFile', { title: 'Excel Automations' });
};	


/**
 * View data page 
 */
exports.readDataFromExcel=function(req,res) {
	 res.render('viewData', { title: 'Excel Automations' });
};


/**
 * File creation main logic
 */
exports.filecreate=function(req,res) {
	var config=req.param("config");
	console.log(config)
	var fileName=req.param("fileName");
	var sheetName=req.param("sheetName");
	var key=config+"--"+sheetName;
	map[key]=config;
	
	var workbook = new Excel.Workbook();
	var sheet = workbook.addWorksheet(sheetName);
	var colNameString=config.split(",");
	var worksheet = workbook.getWorksheet(sheetName);
	var reColumns=[];
	for(i=0;i<colNameString.length;i++){
		var colInfo={};
		var c=colNameString[i].split('-');
		colInfo={header:c[0],key:c[0]};
		reColumns.push(colInfo);
	}
	excelMaps[key]=reColumns;
//	var reColumns=[
//	               {header:'FirstName',key:'firstname'},
//	               {header:'LastName',key:'lastname'},
//	               {header:'Other Name',key:'othername'}
//	           ];
	worksheet.columns = reColumns;
	
	workbook.xlsx.writeFile(fileName).then(function() {
	    console.log("xls file is written.");
	    
	});
	
	res.render('index',{title:'File Created Successfully'});
};



/**
 * Get Update page with entries on column name
 */
exports.updateSheetGet=function(req,res){
	var fileName=req.param('fileName');
	var sheetName=req.param('sheet');
	var key=fileName+"--"+sheetName;
	var reCols=excelMaps[key];
	if(reCols==null){
		addToMaps(fileName,sheetName,res);
	}
	
	
	
};


/**
 * Helper method
 * @param fileName Name of the File to update
 * @param sheetName SheetName to update
 * @param res
 */
function addToMaps(fileName,sheetName,res) {
	console.log("File name in maps");
	console.log(fileName)
	var key=fileName+"--"+sheetName;
	var workbook = new Excel.Workbook();
	workbook.xlsx.readFile(fileName).then(function(){
		var workSheet=workbook.getWorksheet(sheetName);
		console.log("First Column is ")
		console.log(JSON.stringify(workSheet.getRow(1).values));
		var cc=JSON.stringify(workSheet.getRow(1).values);
		console.log(typeof(cc));
		var jc=JSON.parse(cc);
		console.log(jc);
		var reColumns=[];
		for( i=0;i<jc.length;i++){
			console.log(jc[i]);
			if(jc[i]!=null){
				var d={};
				d={header:jc[i],key:jc[i]};
				reColumns.push(d);
			}
		}
		
		excelMaps[key]=reColumns;
		var reCols=excelMaps[key];
		res.render('enterSheetData',{colName:reCols,fileName:fileName,sheetName:sheetName});
	});
}


/**
 * Post processing of update
 */
exports.updateSheetPost=function(req,res){
	var fileName=req.param("fileName");
	var sheetName=req.param("sheeName");
	var key=fileName+"--"+sheetName;
	var reCols=excelMaps[key];
	var rowValues=[];
	for(i=0;i<reCols.length;i++){
		var c=reCols[i];
		rowValues.push(req.param(c['header']));
		
	}
	var workbook = new Excel.Workbook();
	workbook.xlsx.readFile(fileName).then(function(){
		var workSheet=workbook.getWorksheet(sheetName);
		workSheet.addRow(rowValues);
		workbook.xlsx.writeFile(fileName);
		res.render('updateExistingFile', { title: 'Excel Automations' });
	});
	
	
};


/**
 * Get File data
 */
exports.getData=function(req,res){
	var fileName=req.param("fileName");
	var sheetName=req.param("sheeName");
	var allRows=[];
	var workbook = new Excel.Workbook(); 
	workbook.xlsx.readFile(fileName)
	    .then(function() {
	        var worksheet = workbook.getWorksheet(sheetName);
	        worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
	    		var cc=JSON.stringify(row.values);
	    		var jc=JSON.parse(cc);
	    		var reColumns=[];
	    		for( i=0;i<jc.length;i++){
	    			if(jc[i]!=null){
	    				
	    				reColumns.push(jc[i]);
	    			}
	    		}
	    		allRows.push(reColumns);
	    		
	          });
	        
	        res.render('viewSheetDataPage',{data:allRows,fileName:fileName,sheetName:sheetName,title: 'Excel Automations'});
	    });
	
	
	
	
}

