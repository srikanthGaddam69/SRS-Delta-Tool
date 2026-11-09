const fs=require('fs');
let p=fs.readFileSync(__dirname+'/page.html','utf8');
const ex=fs.readFileSync(__dirname+'/../node_modules/exceljs/dist/exceljs.min.js','utf8').replace(/<\/script/gi,'<\\/script');
const core=fs.readFileSync(__dirname+'/core.js','utf8');
p=p.replace('/*EXCELJS*/',()=>ex).replace('/*CORE*/',()=>core);
fs.writeFileSync(__dirname+'/../index.html',p);console.log('size',p.length);
