var START = 1;
var END = 50;
var task = [];
var fs = require('fs');
var xml2js = require('xml2js');
var async = require('async');
var excelbuilder = require('msexcel-builder');


for(var pageNumber=START;pageNumber<=END;pageNumber++){


  (function(pageNumber){
    task.push(function(callback){
      fs.readFile('xml/page_'+pageNumber+'_layout.xml', function(err, result) {
        var miniResult=[];
        if(err){
          console.log(err);
          callback(null,miniResult);
        }else{
          var parser = new xml2js.Parser();
          parser.parseString(result, function (err, layout) {

            layout.layout.object.forEach(function(objLayout){
              objLayout['$'].pageNumber = pageNumber;
              miniResult.push(objLayout['$']);
            });
            callback(err, miniResult);
          });
        }
      });
    });
  })(pageNumber);

}

async.series(task,function(err, finalResult){
  var list = [];
  if(err){
    throw err;
  }else{
    finalResult.forEach(function(dta){
      dta.forEach(function(p){
        if(p.pageNumber && p.id && p.styleName && p.toolTip){
           list.push(p);
        }
      });
    });
  }

  // Create a new workbook file in current working-path
  var workbook = excelbuilder.createWorkbook('./', 'result.xlsx')

  // Create a new worksheet with 10 columns and 12 rows
  var sheet1 = workbook.createSheet('sheet1', 5, list.length + 1);

  // Fill some data
  for(var x=0;x<list.length;x++){
     sheet1.set(1, x + 1, list[x].pageNumber);
     sheet1.set(2, x + 1, list[x].id);
     sheet1.set(3, x + 1, list[x].styleName);
     sheet1.set(4, x + 1, list[x].toolTip);
  }

  // Save it
  workbook.save(function(err){
    if (err)
      throw err;
    else
      console.log('congratulations, your workbook created');
  });

});
