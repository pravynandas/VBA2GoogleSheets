const v2sRules = JSON.parse(JSON.stringify(require('./vba2sheets_macro_rules.json')))
//console.log(v2sRules);

const fs = require('fs');
fs.readFile("./VBASamples/helloworld.vb", function(err, vbBody){
    if (err) throw new Error("File ./VBASamples/helloworld.vb cannot be read");
    let vb = vbBody.toString();
    console.log(vb);

    // fs.readFile("./SheetSamples/helloworld.gs", function(err, gsBody){
    //     if (err) throw new Error("File ./SheetSamples/helloworld.gs cannot be read");
    //     let gs = gsBody.toString();
    //     console.log(gs);
    
        
    // });
    let match_range = vb.match(/.Range\("(.*)"\)/)[1]
    console.log(match_range)
});