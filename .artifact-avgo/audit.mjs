import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const dir='/Users/kevin/Documents/Codex/New project/outputs/01a0a636-1cf4-7961-ba49-c79cbb2a27ff';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(dir+'/AVGO_Valuation_Model.xlsx'));
w.recalculate();
// Independent arithmetic from sourced raw inputs and the Base scenario assumptions.
let ai=57600,other=48363-35900+(34800-21700-8752*1.03),soft=22726+8752*1.03,previousRevenue=105889;
const q4=(34800*(.66-.07)*.83+34800*.007-34800*.02-(105889-71089*4/3)*.15)*48/91;
let pv=q4/1.11**(48/365),cf=0,revenue=0;
for(let i=0;i<5;i++){ai*=1+[1,.55,.30,.20,.12][i];other*=1+[.08,.06,.05,.04,.03][i];soft*=1+[.20,.15,.12,.10,.08][i];revenue=ai+other+soft;cf=revenue*([.65,.64,.63,.62,.61][i]-[.06,.055,.05,.05,.05][i])*.83+revenue*.007-revenue*.02-(revenue-previousRevenue)*.15;pv+=cf/1.11**(48/365+i+1);previousRevenue=revenue;}
const terminalRevenue=revenue*1.03;
const terminal=terminalRevenue*(.61-.05)*.83+terminalRevenue*(.007-.02)-revenue*.15*.03;
const expected=(pv+terminal/(.11-.03)/1.11**(48/365+5)-59419+23975)/4937;
const actual=w.worksheets.getItem('Summary').getRange('E10').values[0][0];
if(Math.abs(actual-expected)>1e-8)throw Error('Independent DCF mismatch');
const s=w.worksheets.getItem('Assumptions');s.getRange('E4').values=[[2]];w.recalculate();const bear=w.worksheets.getItem('Summary').getRange('E10').values[0][0];if(!(bear<actual))throw Error('Exported selector failed');
s.getRange('E4').values=[[1]];w.recalculate();
console.log(JSON.stringify({independentDCF:expected,reopenedDCF:actual,reopenedBear:bear}));
const b=await w.render({sheetName:'Actuals',range:'H6:H38',scale:1.5,format:'png'});await fs.writeFile(dir+'/Sources_QA.png',new Uint8Array(await b.arrayBuffer()));
