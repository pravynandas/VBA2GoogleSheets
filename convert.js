var fs = require('fs');
var es = require('event-stream');
var now = require('performance-now');

var totalLines = 0;
var t0 = now();
var t1;

console.time('line count');
// let _addACI = false;
// let _replaceDN1 = false;
// let _replaceDN2 = false;
// let _replaceOpenSSO = false;
// let _injectApplicationProfileId = false;
// let _currentDN = '';
// let _collectDN = false;

let _outFile = fs.createWriteStream('./SheetSamples/helloworld.vb.gs')
var s = fs
  .createReadStream('./VBASamples/helloworld.vb')
  .pipe(es.split())
  .pipe(
    es
      .mapSync(function(line) {
        let lineOut = line.trim();
        let _skipWrite = false;
        totalLines++;
        /**
         * EDIT 000: replace comments ' with //
         */
         let matches = lineOut.match(/^'/);
         if (matches) {
             console.log(matches[0])
             lineOut = '//' + lineOut.substr(1);
         }

        /**
         * EDIT 001: Replace Range(".*") with spreadsheet.getRange('.*')
         */
        matches = lineOut.match(/^Range\("(.*)"\)/);
        if (matches) {
            console.log(matches[0])
            console.log(matches[1])
            lineOut = lineOut.replace(matches[0], `    spreadsheet.getRange('${matches[1]}')`)
        }

        matches = lineOut.match(/^ActiveSheet\.Range\("(.*)"\)/);
        if (matches) {
            console.log(matches[0])
            console.log(matches[1])
            lineOut = lineOut.replace(matches[0], `spreadsheet.getRange('${matches[1]}')`)
        }   

        /**
         * EDIT 002: Replace ActiveCell. with spreadsheet.getCurrentCell().
         */
        matches = lineOut.match(/^ActiveCell./);
        if (matches) {
            console.log(matches[0])
            console.log(matches[1])
            lineOut = lineOut.replace(matches[0], `spreadsheet.getCurrentCell().`)
        } 

        /**
         * EDIT 003: Replace [FormulaR1C1 = ".*"] with spreadsheet.setFormula('.*').
         */
        matches = lineOut.match(/\.FormulaR1C1 = "(.*)"/);
        if (matches) {
            console.log(matches[0])
            console.log(matches[1])
            let _replaceFormula = matches[1];
            let _formula = matches[1].trim().toUpperCase();
            if (_formula.startsWith("=")) {
                if (_formula.startsWith("=CONCATENATE")) {
                    let _formulaMatches = _formula.match(/=CONCATENATE\((.*)\)/);
                    if (_formulaMatches) {
                        console.log(_formulaMatches[0]);
                        console.log(_formulaMatches[1]);
                        let _formulaParams = _formulaMatches[1];
                        let _a_formulaParams = _formulaParams.split(",");
                        console.log(_a_formulaParams)
                        let _gs_params = _a_formulaParams.reduce((arr, _param)=>{
                            if(_param.startsWith("RC")){
                                arr.push(_param.replace("RC", "R[0]C"))
                            } else if(_param.endsWith("C")){
                                arr.push(_param.replace("C", "C[0]"))
                            }
                            return arr;
                        },[]).join(',');
                        console.log('gs params:', _gs_params);
                        _replaceFormula = `=CONCATENATE(${_gs_params})`;
                    }
                }
                lineOut = lineOut.replace(matches[0], `.setFormula('${_replaceFormula}')`)
            } else {
                lineOut = lineOut.replace(matches[0], `.setValue('${_replaceFormula}')`)
            }
            
        } 

        /**
         * EDIT 004: Replace .Select with .activate();.
         */
        matches = lineOut.match(/\.Select$/);
        if (matches) {
            console.log(matches[0])
            console.log(matches[1])
            lineOut = lineOut.replace(matches[0], `.activate();`)
        } 

        /**
         * EDIT 005: Replace sub with function
         */
        matches = lineOut.match(/^Sub\s(.*)\((.*)\)/);
        if (matches) {
            console.log(matches[0])
            console.log(matches[1])
            console.log(matches[2])
            lineOut = lineOut.replace(matches[0], `function ${matches[1]}(${matches[2]}) {\n\n\tvar spreadsheet = SpreadsheetApp.getActive();`)
        } 
        matches = lineOut.match(/^End Sub$/);
        if (matches) {
            console.log(matches[0])
            lineOut = lineOut.replace(matches[0], `}`)
        } 



        // Write to out stream
        if (_skipWrite) {
            _skipWrite = false;
        } else {
            _outFile.write(lineOut + '\n');
        }
      })
      .on('error', function(err) {
        _outFile.close();
        console.log('Error while reading file.', err);
      })
      .on('end', function() {
        _outFile.close();
        console.log('Read entire file.');
        t1 = now();
        console.log(totalLines);
        console.timeEnd('line count');
        console.log(
          `Performance now line count timing: ` + (t1 - t0).toFixed(3) + `ms`,
        );

      }),
  );
