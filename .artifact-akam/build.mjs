import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const out='/Users/kevin/Documents/Codex/New project/outputs/01a0b0af-17ce-75d3-87a0-681f422d5b6f';
const wb=Workbook.create();
const names=['Summary','Assumptions','Forecast','Valuation','Actuals','Checks'];
const ss=Object.fromEntries(names.map(n=>[n,wb.worksheets.add(n)]));
const navy='#17324D',blue='#0000FF',green='#008000',gray='#E9EEF3';
const num='#,##0.0;(#,##0.0);"-"',pct='0.0%;(0.0%);"-"',price='$0.00;($0.00);$0.00';
const v=(s,a,x)=>{ss[s].getRange(a).values=[[x]];if(typeof x==='number')ss[s].getRange(a).format.font.color=blue;};
const f=(s,a,x)=>{ss[s].getRange(a).formulas=[[x]];ss[s].getRange(a).format.font.color=s==='Summary'?navy:x.includes('!')?green:'#000000';};
const row=(s,r,l)=>v(s,'C'+r,l);
const band=(s,r,l,end='K')=>{row(s,r,l);ss[s].getRange(`C${r}:${end}${r}`).format.fill=gray;ss[s].getRange(`C${r}:${end}${r}`).format.font.bold=true;};
const note=(s,r,l)=>{row(s,r,l);ss[s].getRange('C'+r).format.font={name:'Arial',size:10,color:'#526171'};};
const total=(s,r,end='M')=>{ss[s].getRange(`C${r}:${end}${r}`).format.borders={top:{style:'thin',color:'#A6B3C0'}};ss[s].getRange(`C${r}:${end}${r}`).format.font.bold=true;};
const header=(s,a,vals)=>{ss[s].getRange(a).values=[vals];ss[s].getRange(a).format.fill=navy;ss[s].getRange(a).format.font={name:'Arial',size:10,bold:true,color:'#FFFFFF'};ss[s].getRange(a).format.horizontalAlignment='center';};
const per=(s,r,from='E',to='M')=>{ss[s].getRange(`${from}${r}:${to}${r}`).setNumberFormat(pct);ss[s].getRange(`C${r}:${to}${r}`).format.font.italic=true;};
for(const n of names){const s=ss[n];s.showGridLines=false;s.getRange('A1:O105').format.font={name:'Arial',size:10,color:'#000000'};s.getRange('A1:B105').format.columnWidth=2.5;s.getRange('C1:C105').format.columnWidth=43;s.getRange('D1:D105').format.columnWidth=3;s.getRange('E1:M105').format.columnWidth=15;s.getRange('C1:M105').format.rowHeight=20;s.getRange('E1:M105').setNumberFormat(num);s.getRange('C2').values=[[`AKAM | ${n==='Summary'?'Valuation':n}`]];s.getRange('C2').format.font={name:'Arial',size:16,bold:true,color:navy};s.getRange('C3:M3').format.borders={bottom:{style:'thin',color:navy}};row(n,4,'Case Selected:');if(n!=='Assumptions')f(n,'E4',"=CHOOSE('Assumptions'!$E$4,\"Base\",\"Bear\",\"Bull\")");s.getRange('E4').format.font.color=green;s.getRange('E4').format.horizontalAlignment='center';s.getRange('E4').format.borders={preset:'outside',style:'dashed',color:navy};if(n!=='Summary')s.freezePanes.freezeRows(6);}
ss.Summary.tabColor=navy;ss.Assumptions.tabColor='#50789D';ss.Actuals.tabColor='#D8C8AD';
const url={ir:'https://www.ir.akamai.com/news-releases/news-release-details/akamai-reports-second-quarter-2026-financial-results',sec:'https://www.sec.gov/Archives/edgar/data/1086222/000108622226000086/akam-20260630.htm',sup:'https://www.ir.akamai.com/static-files/6b47a3ae-7484-4682-b959-73a68a8842b5',market:'https://www.financecharts.com/stocks/AKAM/summary/price',market2:'https://finance.yahoo.co.jp/quote/AKAM/history',call:'https://www.fool.com/earnings/call-transcripts/2026/08/13/akamai-akam-q2-2026-earnings-call-transcript/'};
// Raw sources retain disclosed meaning. Amounts and shares are converted from thousands to millions.
note('Actuals',5,'USD millions and million shares, except prices. Source dates and locators are at right.');
header('Actuals','E6:G6',['2025 A','H1:26 A','Q2:26 A']);
const actual=[
 [7,'Security revenue',2243.404,1194.226,604.436],
 [8,'Delivery and other cloud applications',1650.862,785.135,395.927],
 [9,'Cloud infrastructure services',313.909,193.931,99.319],
 [10,'Total revenue',4208.175,2173.292,1099.682],
 [11,'GAAP operating income',null,194.778,80.284],
 [12,'Non-GAAP operating income',null,553.512,270.710],
 [13,'Expensed stock-based compensation',null,274.971,146.290],
 [14,'Total current stock compensation',null,348.113,178.157],
 [15,'Non-GAAP depreciation and amortization',null,289.223,145.356],
 [16,'Operating cash flow',null,638.774,326.266],
 [17,'Cash capex including internal software',null,417.600,225.753],
 [18,'Company capex, accrual basis',null,552.873,346.536],
 [19,'Acquired intangible amortization',null,50.276,25.089],
 [20,'Capitalized SBC / interest amortization',null,30.450,15.434],
 [21,'Restructuring charges',null,2.008,1.825],
 [22,'Acquisition costs / (benefit)',null,1.029,1.788],
 [23,'Interest and investment income',null,49.219,31.672],
 [24,'GAAP interest expense',null,17.335,9.078],
 [25,'Debt issuance cost amortization',null,5.180,3.032]
];
for(const [r,l,...xs] of actual){row('Actuals',r,l);xs.forEach((x,j)=>{if(x!==null)v('Actuals',String.fromCharCode(69+j)+r,x);});}
band('Actuals',27,'Capital structure, guidance and price','G');
const raws=[
 [28,'Cash and cash equivalents',1480.257],[29,'Current marketable securities',1875.130],[30,'Non-current marketable securities',1260.918],
 [31,'Debt principal',7640],[32,'Debt carrying value',7562.828],[33,'LayerX cash purchase, July 2026',205],
 [34,'2026 guided diluted shares, hedge adjusted',150],[35,'Reference closing price',104.48],
 [36,'Price / valuation date',new Date('2026-09-16T00:00:00Z')],[37,'Balance sheet date',new Date('2026-06-30T00:00:00Z')],[38,'2026 year end',new Date('2026-12-31T00:00:00Z')],
 [39,'2026 revenue guidance, low',4445],[40,'2026 revenue guidance, high',4530],[41,'2026 non-GAAP EBIT margin, low',.25],[42,'2026 non-GAAP EBIT margin, high',.26],[43,'2026 non-GAAP EPS, low',6.40],[44,'2026 non-GAAP EPS, high',7.05],[45,'2026 non-GAAP tax guidance',.19],
 [46,'2026 capex / revenue, call transcript',.40]
];for(const [r,l,x] of raws){row('Actuals',r,l);v('Actuals','E'+r,x);}
ss.Actuals.getRange('E35').setNumberFormat(price);ss.Actuals.getRange('E36:E38').setNumberFormat('mmm d, yyyy');for(const r of [41,42,45,46])per('Actuals',r,'E','E');ss.Actuals.getRange('E43:E44').setNumberFormat(price);
header('Actuals','E49:G49',['Note maturity','Warrants (M)','Strike / share']);
row('Actuals',49,'Hedges offset note conversion premium');
[[2033,18.546,155.02],[2027,9.898,178.74],[2029,10.015,180.44],[2030,8.689,247.35],[2032,9.171,282.68]].forEach((x,j)=>{ss.Actuals.getRange(`E${50+j}:G${50+j}`).values=[x];});ss.Actuals.getRange('E50:E54').setNumberFormat('0');ss.Actuals.getRange('G50:G54').setNumberFormat(price);
const sourceBlocks=[
 [7,'Rows 7–10: Q2 supplement, p. 3. 2025 categories are recast.',url.sup],
 [12,'Rows 11–13, 15–25: Q2 release, financial tables.',url.ir],
 [17,'Row 14: 10-Q, stockholders equity, pp. 9–12. Includes capitalized SBC.',url.sec],
 [22,'Rows 28–34, 50–54: 10-Q, balance sheet and Notes 6–7; release guidance.',url.sec],
 [28,'Rows 35–36: September 16 close, retrieved September 17, 2026.',url.market],
 [33,'Price cross-check: Yahoo Japan also reports $104.48. Another feed showed $104.45.',url.market2],
 [39,'Rows 39–45: Q2 release, full-year guidance, August 6, 2026.',url.ir],
 [44,'Row 46: CFO remarks, Aug. 6 call, third-party transcript. Accrual basis.',url.call]
];ss.Actuals.getRange('I1:I60').format.columnWidth=84;v('Actuals','I6','Source / locator');
for(const [r,l,u] of sourceBlocks){v('Actuals','I'+r,l);v('Actuals','I'+(r+1),u);ss.Actuals.getRange(`I${r}:I${r+1}`).format.wrapText=true;ss.Actuals.getRange(`I${r}:I${r+1}`).format.rowHeight=42;}
note('Actuals',57,'Retrieved September 17, 2026. Blank 2025 cells are not collected, not zero.');
note('Actuals',58,'Company has one reportable segment. Revenue categories are modeled separately.');
// One selector and active assumption row per driver.
v('Assumptions','E4',1);ss.Assumptions.getRange('E4').setNumberFormat('0');ss.Assumptions.getRange('E4').dataValidation={rule:{type:'list',values:['1','2','3']}};v('Assumptions','G4','1 Base / 2 Bear / 3 Bull');
row('Assumptions',7,'Reference price');f('Assumptions','E7',"='Actuals'!E35");ss.Assumptions.getRange('E7').setNumberFormat(price);
row('Assumptions',8,'Valuation date');f('Assumptions','E8',"='Actuals'!E36");ss.Assumptions.getRange('E8').setNumberFormat('mmm d, yyyy');
row('Assumptions',9,'Base diluted shares, hedge adjusted');f('Assumptions','E9',"='Actuals'!E34");
header('Assumptions','E11:J11',['2026 E','2027 E','2028 E','2029 E','2030 E','2031 E']);
const drivers=[
 [12,'Security revenue growth',[.095,.10,.095,.09,.085,.08],[.08,.06,.06,.055,.05,.05],[.10,.13,.13,.12,.11,.10]],
 [17,'Delivery / other cloud growth',[-.055,-.05,-.045,-.04,-.03,-.02],[-.07,-.09,-.08,-.07,-.06,-.05],[-.045,-.02,0,.01,.02,.02]],
 [22,'Cloud infrastructure growth',[.55,.85,.50,.35,.25,.20],[.45,.40,.30,.25,.20,.15],[.65,1.20,.75,.50,.35,.25]],
 [27,'Non-GAAP operating margin',[.255,.25,.27,.29,.305,.315],[.25,.225,.23,.245,.255,.26],[.26,.27,.30,.325,.34,.35]],
 [32,'Expensed SBC / revenue',[.125,.12,.115,.105,.10,.095],[.13,.135,.13,.125,.12,.115],[.12,.11,.10,.09,.08,.075]],
 [37,'Capitalized SBC / revenue',[.033,.032,.03,.028,.025,.023],[.035,.035,.033,.032,.03,.028],[.03,.028,.026,.024,.022,.02]],
 [42,'Non-GAAP D&A / revenue',[.135,.17,.18,.18,.175,.17],[.135,.17,.19,.19,.185,.18],[.135,.17,.18,.18,.175,.17]],
 [47,'Cash capex / revenue',[.37,.30,.25,.22,.20,.19],[.39,.35,.30,.27,.24,.22],[.36,.32,.27,.23,.205,.19]],
 [52,'Incremental NWC / revenue increase',[.10,.10,.10,.10,.10,.10],[.15,.15,.15,.15,.15,.15],[.08,.08,.08,.08,.08,.08]],
 [57,'Normalized cash tax rate',[.21,.21,.21,.21,.21,.21],[.23,.23,.23,.23,.23,.23],[.20,.20,.20,.20,.20,.20]]
];
for(const [r,l,...cases] of drivers){band('Assumptions',r,l+' — active','J');for(let k=0;k<3;k++){row('Assumptions',r+1+k,['Base','Bear','Bull'][k]);ss.Assumptions.getRange(`E${r+1+k}:J${r+1+k}`).values=[cases[k]];ss.Assumptions.getRange(`E${r+1+k}:J${r+1+k}`).format.font.color=blue;}for(let j=0;j<6;j++){const c=String.fromCharCode(69+j);f('Assumptions',c+r,`=IF(ISNUMBER(CHOOSE($E$4,${c}${r+1},${c}${r+2},${c}${r+3})),CHOOSE($E$4,${c}${r+1},${c}${r+2},${c}${r+3}),NA())`);}ss.Assumptions.getRange(`E${r}:J${r+3}`).setNumberFormat(pct);ss.Assumptions.getRange(`C${r}:J${r+3}`).format.font.italic=true;}
header('Assumptions','E63:H63',['Active','Base','Bear','Bull']);
const scalars=[[64,'WACC',.10,.115,.09],[65,'Terminal growth',.03,.025,.035],[66,'Forward non-GAAP P/E',17,13,22],[67,'2027+ net interest income / (expense)',50,0,75],[68,'2027+ non-GAAP earnings tax',.19,.21,.19]];
for(const [r,l,...xs] of scalars){row('Assumptions',r,l);ss.Assumptions.getRange(`F${r}:H${r}`).values=[xs];ss.Assumptions.getRange(`F${r}:H${r}`).format.font.color=blue;f('Assumptions','E'+r,`=IF(ISNUMBER(CHOOSE($E$4,F${r},G${r},H${r})),CHOOSE($E$4,F${r},G${r},H${r}),NA())`);}
for(const r of [64,65,68])per('Assumptions',r,'E','H');ss.Assumptions.getRange('E66:H66').setNumberFormat('0.0"x"');
const an=[
 'Blue cells are editable. Each case recalculates the same forecast and valuation.',
 'Cases are analyst assumptions, not consensus forecasts. 2026 revenue is built by category.',
 'Base 2026 cash capex of 37% plus capitalized SBC of 3.3% approximates 40% accrual guidance.',
 'Cash payment timing is uncertain. Forecast capex includes software; acquisitions are separate.',
 'SBC: expensed awards reduce operating profit. New capitalized awards are deducted in full.',
 'D&A uses non-GAAP D&A, avoiding a second addback of excluded intangible / SBC amortization.',
 'Terminal ratios stay at 2031 levels. Working capital growth resets to terminal revenue growth.',
 'WACC and P/E are judgmental valuation hurdles, not observed peer multiples or fitted estimates.',
 'Taxes: DCF uses normalized cash taxes. P/E uses a separate non-GAAP earnings tax assumption.',
 'Net interest is a simplifying forecast. Principal maturities are assumed refinanced; no buybacks.',
 'Interim cash bridge uses a straight-line share of H2 economic FCF; actual timing may differ.'
 ];an.forEach((x,i)=>note('Assumptions',71+i,x));
// Actuals + annual forecast, with H2 computed as the full-year estimate less reported H1.
note('Forecast',5,'USD millions, except EPS. A = actual. E = estimate. Full-year cash flows overlap H2; DCF excludes that overlap.');
header('Forecast','E6:M6',['2025 A','H1:26 A','H2:26 E','2026 E','2027 E','2028 E','2029 E','2030 E','2031 E']);
const labels={8:'Security revenue',9:'Delivery and other cloud applications',10:'Cloud infrastructure services',11:'Total revenue',12:'Revenue growth',14:'Non-GAAP operating margin',15:'Non-GAAP operating income',16:'Expensed SBC',17:'Operating income after expensed SBC',18:'Normalized cash tax rate',19:'Normalized operating cash taxes',20:'NOPAT after expensed SBC',22:'Non-GAAP depreciation and amortization',23:'Cash capex including software',24:'Current capitalized SBC cost',25:'Incremental working capital investment',27:'Unlevered cash flow after all SBC',29:'Annual net interest income / (expense)',30:'Non-GAAP earnings tax rate',31:'Non-GAAP net income (modeled)',32:'Base diluted shares, hedge adjusted',33:'Non-GAAP EPS (modeled)',35:'Reported operating cash flow',36:'Reported cash FCF before SBC',37:'Reported FCF less all current SBC',39:'Cash capex / revenue',40:'All current SBC / revenue'};
for(const [r,l] of Object.entries(labels))row('Forecast',r,l);
for(const [r,ar] of [[8,7],[9,8],[10,9],[11,10]])for(const c of ['E','F'])f('Forecast',c+r,`='Actuals'!${c}${ar}`);
for(const [r,ar] of [[15,12],[16,13],[22,15],[23,17],[35,16]])f('Forecast','F'+r,`='Actuals'!F${ar}`);
f('Forecast','F14','=F15/F11');f('Forecast','F17','=F15-F16');f('Forecast','F24',"='Actuals'!F14-'Actuals'!F13");f('Forecast','F36','=F35-F23');f('Forecast','F37','=F36-F16-F24');
for(let j=0;j<6;j++){const c=String.fromCharCode(72+j),p=j===0?'E':String.fromCharCode(71+j),a=String.fromCharCode(69+j);for(const [r,ar] of [[8,12],[9,17],[10,22]])f('Forecast',c+r,`=${p}${r}*(1+'Assumptions'!${a}${ar})`);f('Forecast',c+'11',`=SUM(${c}8:${c}10)`);f('Forecast',c+'12',`=${c}11/${p}11-1`);f('Forecast',c+'14',`='Assumptions'!${a}27`);f('Forecast',c+'15',`=${c}11*${c}14`);f('Forecast',c+'16',`=${c}11*'Assumptions'!${a}32`);f('Forecast',c+'17',`=${c}15-${c}16`);f('Forecast',c+'18',`='Assumptions'!${a}57`);f('Forecast',c+'19',`=MAX(0,${c}17)*${c}18`);f('Forecast',c+'20',`=${c}17-${c}19`);for(const [r,ar] of [[22,42],[23,47],[24,37]])f('Forecast',c+r,`=${c}11*'Assumptions'!${a}${ar}`);if(j>0){f('Forecast',c+'25',`=(${c}11-${p}11)*'Assumptions'!${a}52`);f('Forecast',c+'27',`=${c}20+${c}22-${c}23-${c}24-${c}25`);f('Forecast',c+'29',"='Assumptions'!E67");f('Forecast',c+'30',"='Assumptions'!E68");f('Forecast',c+'31',`=(${c}15+${c}29)*(1-${c}30)`);f('Forecast',c+'32',"='Assumptions'!E9");f('Forecast',c+'33',`=${c}31/${c}32`);}}
for(const r of [8,9,10,11,15,16,17,22,23,24])f('Forecast','G'+r,`=H${r}-F${r}`);
f('Forecast','G14','=G15/G11');f('Forecast','G18',"='Assumptions'!E57");f('Forecast','G19','=MAX(0,G17)*G18');f('Forecast','G20','=G17-G19');f('Forecast','G25',"=(H11-F11*2)*'Assumptions'!E52");f('Forecast','G27','=G20+G22-G23-G24-G25');
for(const c of ['F','G','H','I','J','K','L','M']){f('Forecast',c+'39',`=${c}23/${c}11`);f('Forecast',c+'40',`=(${c}16+${c}24)/${c}11`);}
for(const r of [12,14,18,30,39,40])per('Forecast',r);ss.Forecast.getRange('E33:M33').setNumberFormat(price);for(const r of [11,15,17,20,27,31,33,36,37])total('Forecast',r);
note('Forecast',43,'H2 = annual forecast less H1 actuals. H1 NOPAT / UFCF is not manufactured from reported cash flow.');
note('Forecast',44,'H2 NWC uses growth from annualized H1 revenue. Later years use the prior full-year revenue.');
note('Forecast',45,'All SBC is an economic cost; forecast shares stay flat before existing warrant dilution.');
note('Forecast',46,'Operating leases stay in operating costs. Lease liabilities are not also subtracted as debt.');
note('Forecast',47,'2026 P/E cross-check uses company EPS guidance on Valuation; 2027 EPS is modeled here.');
note('Forecast',48,'Terminal capex remains above D&A. No acquisition amortization tax shield is assumed.');
// DCF with an explicit, conservative economic cash bridge from the last reported balance sheet.
note('Valuation',5,'Current value at September 16, 2026. End-period discounting; straight-line H2 timing approximation.');
header('Valuation','E7:J7',['H2 remainder','2027 E','2028 E','2029 E','2030 E','2031 E']);
for(const [r,l] of [[8,'Years to period end'],[9,'Full-period UFCF after all SBC'],[10,'Share of period remaining'],[11,'Cash flow to discount'],[12,'End-period discount factor'],[13,'Present value of cash flow']])row('Valuation',r,l);
f('Valuation','E8',"=('Actuals'!E38-'Assumptions'!E8)/365");f('Valuation','E10',"=('Actuals'!E38-'Assumptions'!E8)/('Actuals'!E38-'Actuals'!E37)");
for(let j=0;j<6;j++){const c=String.fromCharCode(69+j),fc=j===0?'G':String.fromCharCode(72+j);if(j>0){f('Valuation',c+'8',`=$E$8+${j}`);v('Valuation',c+'10',1);}f('Valuation',c+'9',`='Forecast'!${fc}27`);f('Valuation',c+'11',`=${c}9*${c}10`);f('Valuation',c+'12',`=1/(1+'Assumptions'!$E$64)^${c}8`);f('Valuation',c+'13',`=${c}11*${c}12`);}
ss.Valuation.getRange('E8:J8').setNumberFormat('0.00');ss.Valuation.getRange('E10:J10').setNumberFormat(pct);ss.Valuation.getRange('E12:J12').setNumberFormat('0.000');
const vr={16:'PV of explicit cash flows',17:'Terminal revenue',18:'Terminal NOPAT after expensed SBC',19:'Terminal non-GAAP D&A',20:'Terminal cash capex',21:'Terminal capitalized SBC',22:'Terminal incremental NWC',23:'Terminal unlevered cash flow',24:'Terminal enterprise value',25:'PV of terminal value',26:'Enterprise value',28:'Reported cash and securities',29:'LayerX cash acquisition',30:'Elapsed H2 economic cash flow',31:'Elapsed H2 net interest after tax',32:'Adjusted economic cash balance',33:'Debt principal',34:'Equity value before warrants',35:'Base diluted shares',36:'Value before warrant dilution',37:'DCF value after warrant dilution',38:'Upside / (downside)',39:'PV terminal / enterprise value'};for(const [r,l] of Object.entries(vr))row('Valuation',r,l);
const vf={16:'=SUM(E13:J13)',17:"='Forecast'!M11*(1+'Assumptions'!E65)",18:"=E17*('Assumptions'!J27-'Assumptions'!J32)*(1-'Assumptions'!J57)",19:"=E17*'Assumptions'!J42",20:"=E17*'Assumptions'!J47",21:"=E17*'Assumptions'!J37",22:"='Forecast'!M11*'Assumptions'!E65*'Assumptions'!J52",23:'=E18+E19-E20-E21-E22',24:"=IF('Assumptions'!E64>'Assumptions'!E65,E23/('Assumptions'!E64-'Assumptions'!E65),NA())",25:'=E24*J12',26:'=E16+E25',28:"=SUM('Actuals'!E28:E30)",29:"='Actuals'!E33",30:'=E9*(1-E10)',31:"='Assumptions'!E67*(1-'Assumptions'!E68)*('Assumptions'!E8-'Actuals'!E37)/365",32:'=E28-E29+E30+E31',33:"='Actuals'!E31",34:'=E26+E32-E33',35:"='Assumptions'!E9",36:'=E34/E35',37:'=MIN(E36,K45:K49)',38:"=E37/'Assumptions'!E7-1",39:'=E25/E26'};for(const [r,x] of Object.entries(vf))f('Valuation','E'+r,x);
for(const r of [36,37])ss.Valuation.getRange('E'+r).setNumberFormat(price);for(const r of [38,39])per('Valuation',r,'E','E');for(const r of [23,26,32,34,37])total('Valuation',r,'E');
v('Valuation','G16','DCF sensitivity ($/share)');v('Valuation','G17','WACC / growth');
// Every cell recomputes discounting and terminal working capital, then applies the warrant dilution solution.
f('Valuation','E37','=MAX(0,MIN(E36,K45:K49))');
const dilute=eq=>`MAX(0,MIN((${eq})/$E$35,${[45,46,47,48,49].map(r=>`((${eq})+$I$${r})/($E$35+$H$${r})`).join(',')}))`;
for(let j=0;j<5;j++){const c=String.fromCharCode(72+j);f('Valuation',c+'18',`='Assumptions'!$E$65+${(j-2)*.005}`);}
for(let i=0;i<5;i++){const r=19+i;f('Valuation','G'+r,`='Assumptions'!$E$64+${(i-2)*.01}`);for(let j=0;j<5;j++){const c=String.fromCharCode(72+j),pv=['E','F','G','H','I','J'].map(x=>`$${x}$11/(1+$G${r})^$${x}$8`).join('+');const terminal=`(('Forecast'!$M$20+'Forecast'!$M$22-'Forecast'!$M$23-'Forecast'!$M$24)*(1+${c}$18)-'Forecast'!$M$11*${c}$18*'Assumptions'!$J$52)`;const eq=`${pv}+${terminal}/($G${r}-${c}$18)/(1+$G${r})^$J$8+$E$32-$E$33`;f('Valuation',c+r,`=IF($G${r}>${c}$18,${dilute(eq)},NA())`);}}
ss.Valuation.getRange('H18:L18').setNumberFormat(pct);ss.Valuation.getRange('G19:G23').setNumberFormat(pct);ss.Valuation.getRange('H19:L23').setNumberFormat(price);ss.Valuation.getRange('H19:L23').conditionalFormats.add('colorScale',{colors:['#FBE9E7','#FFFFFF','#DCE9F5']});ss.Valuation.getRange('J21').format.borders={preset:'outside',style:'medium',color:navy};
v('Valuation','G27','Earnings cross-check');
const peRows=[[28,'2026 guided EPS midpoint',"=AVERAGE('Actuals'!E43:E44)"],[29,'2027 modeled non-GAAP EPS',"='Forecast'!I33"],[30,'Assumed 2027 P/E',"='Assumptions'!E66"],[31,'P/E equity value before warrants','=J29*J30*$E$35'],[32,'P/E value after warrants',`=${dilute('$J$31')}`],[33,'P/E upside / (downside)',"=J32/'Assumptions'!E7-1"],[34,'Market P/E on 2026 guidance',"='Assumptions'!E7/J28"],[35,'Market P/E on modeled 2027',"='Assumptions'!E7/J29"]];
for(const [r,l,x] of peRows){v('Valuation','G'+r,l);f('Valuation','J'+r,x);}for(const r of [28,29,32])ss.Valuation.getRange('J'+r).setNumberFormat(price);ss.Valuation.getRange('J33').setNumberFormat(pct);for(const r of [30,34,35])ss.Valuation.getRange('J'+r).setNumberFormat('0.0"x"');
band('Valuation',42,'Existing warrants: equity value plus exercise proceeds, divided by diluted shares','K');
header('Valuation','E44:K44',['Maturity','Strike','Shares (M)','Cum. shares','Cum. proceeds','DCF before','DCF candidate']);
for(let i=0;i<5;i++){const r=45+i,ar=50+i;f('Valuation','E'+r,`='Actuals'!E${ar}`);f('Valuation','F'+r,`='Actuals'!G${ar}`);f('Valuation','G'+r,`='Actuals'!F${ar}`);f('Valuation','H'+r,`=SUM(G45:G${r})`);f('Valuation','I'+r,`=SUMPRODUCT(F45:F${r},G45:G${r})`);f('Valuation','J'+r,'=$E$36');f('Valuation','K'+r,`=($E$34+I${r})/($E$35+H${r})`);}
ss.Valuation.getRange('E45:E49').setNumberFormat('0');ss.Valuation.getRange('F45:F49').setNumberFormat(price);ss.Valuation.getRange('J45:K49').setNumberFormat(price);
note('Valuation',52,'Minimum candidate price solves dilution across sorted strike thresholds, assuming all warrants outstanding.');
note('Valuation',53,'Debt stays at face value. Note hedges offset conversion premium; warrants are valued separately.');
note('Valuation',54,'Cash bridge deducts SBC as an economic reserve; it is not a forecast of reported bank cash.');
note('Valuation',55,'H2 cash flows are allocated evenly around the valuation date. Capital-spending timing is a key limitation.');
note('Valuation',56,'No blended target. Equity is floored at zero; negative raw equity indicates a funding / enterprise-value shortfall.');
note('Valuation',57,'Warrant treatment uses intrinsic dilution, not option time value. Out-of-money warrants have no current deduction.');
// Summary uses finished outputs only.
note('Summary',5,'USD millions, except per-share values. Prepared September 17, 2026; reference price September 16.');
const sumRows=[[7,'Reference price',"='Assumptions'!E7"],[9,'DCF fair value / share',"='Valuation'!E37"],[10,'Upside / (downside)',"='Valuation'!E38"],[12,'P/E cross-check / share',"='Valuation'!J32"],[13,'2027 forward P/E assumption',"='Assumptions'!E66"],[15,'WACC',"='Assumptions'!E64"],[16,'Terminal growth',"='Assumptions'!E65"],[17,'PV terminal / enterprise value',"='Valuation'!E39"]];
for(const [r,l,x] of sumRows){row('Summary',r,l);f('Summary','E'+r,x);}for(const r of [7,9,12])ss.Summary.getRange('E'+r).setNumberFormat(price);for(const r of [10,15,16,17])per('Summary',r,'E','E');ss.Summary.getRange('E13').setNumberFormat('0.0"x"');total('Summary',9,'E');total('Summary',12,'E');
ss.Summary.getRange('D1:G40').format.columnWidth=19;
header('Summary','C21:H21',['Year','Revenue','Security','Cloud infra.','UFCF after SBC','Non-GAAP EPS']);
for(let j=0;j<5;j++){const r=22+j,c=String.fromCharCode(73+j);v('Summary','C'+r,`${2027+j} E`);for(const [sc,fr] of [['D',11],['E',8],['F',10],['G',27],['H',33]])f('Summary',sc+r,`='Forecast'!${c}${fr}`);}ss.Summary.getRange('D22:G26').setNumberFormat(num);ss.Summary.getRange('H22:H26').setNumberFormat(price);ss.Summary.getRange('G1:H40').format.columnWidth=19;
const chart=ss.Summary.charts.add('line',ss.Summary.getRange('C21:F26'));chart.title='Revenue forecast ($M)';chart.titleTextStyle.typeface='Arial';chart.titleTextStyle.fontSize=12;chart.setPosition('G7','M19');chart.legend={position:'bottom',textStyle:{typeface:'Arial',fontSize:10}};chart.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:10}};chart.yAxis={numberFormatCode:'#,##0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};chart.series.items.forEach((x,i)=>{x.line={fill:[navy,'#407BA7','#D78C28'][i],style:'solid',width:2};});
band('Summary',29,'What drives the valuation','M');
['Change the case selector on Assumptions: 1 Base, 2 Bear, 3 Bull. Blue cells are editable.',
 'DCF deducts all current SBC and cash capital spending. P/E uses non-GAAP earnings before SBC.',
 'The cloud build-out depresses near-term cash flow; the DCF depends on later growth and capex normalization.',
 'The net-debt bridge includes LayerX and estimated interim economic cash flow. Actual cash timing is uncertain.',
 'WACC, long-term growth, margins and multiples are analyst assumptions. No live market-data connection.'
].forEach((x,i)=>note('Summary',31+i,x));
// Independent reconciliation / reasonableness tests are terminal and do not feed outputs.
header('Checks','E6:F6',['Difference / test','Status']);
const checks=[
 ['2025 category revenue reconciliation',"=SUM('Actuals'!E7:E9)-'Actuals'!E10"],
 ['H1 category revenue reconciliation',"=SUM('Actuals'!F7:F9)-'Actuals'!F10"],
 ['Q2 category revenue reconciliation',"=SUM('Actuals'!G7:G9)-'Actuals'!G10"],
 ['H1 GAAP to non-GAAP EBIT bridge',"='Actuals'!F11+'Actuals'!F13+SUM('Actuals'!F19:F22)-'Actuals'!F12"],
 ['H1 cash FCF reconciliation',"='Forecast'!F36-('Actuals'!F16-'Actuals'!F17)"],
 ['H2 plus H1 revenue equals full year',"='Forecast'!G11+'Forecast'!F11-'Forecast'!H11"],
 ['H2 cash capex nonnegative',"=IF('Forecast'!G23>=0,0,1)"],
 ['H2 category revenue nonnegative',"=IF(MIN('Forecast'!G8:G10)>=0,0,1)"],
 ['WACC exceeds terminal growth',"=IF('Assumptions'!E64>'Assumptions'!E65,0,1)"],
 ['Sensitivity at active WACC / growth',"='Valuation'!J21-'Valuation'!E37"],
 ['Enterprise to equity reconciliation',"='Valuation'!E34-('Valuation'!E26+'Valuation'!E32-'Valuation'!E33)"],
 ['H2 elapsed / remaining UFCF split',"='Valuation'!E30+'Valuation'!E11-'Forecast'!G27"],
 ['Valuation date lies within H2:26',"=IF(AND('Assumptions'!E8>='Actuals'!E37,'Assumptions'!E8<'Actuals'!E38),0,1)"],
 ['Diluted share count positive',"=IF('Assumptions'!E9>0,0,1)"],
 ['Dilution / limited-liability bounds',"=IF(AND('Valuation'!E37<=MAX(0,'Valuation'!E36),'Valuation'!E37>=0),0,1)"],
 ['2026 revenue within company guide',"=IF(AND('Forecast'!H11>='Actuals'!E39,'Forecast'!H11<='Actuals'!E40),0,1)"],
 ['Terminal cash capex covers D&A',"=IF('Valuation'!E20>='Valuation'!E19,0,1)"],
 ['2028 incremental working capital',"='Forecast'!J25-('Forecast'!J11-'Forecast'!I11)*'Assumptions'!G52"]
];checks.forEach(([l,x],i)=>{const r=i+7;row('Checks',r,l);f('Checks','E'+r,x);f('Checks','F'+r,`=IF(ABS(E${r})<0.01,"OK","REVIEW")`);});ss.Checks.getRange('E7:E24').setNumberFormat('0.00;(0.00);0.00');ss.Checks.getRange('E7:F24').format.font.color='#000000';ss.Checks.getRange('E7:F24').conditionalFormats.addCustom('ABS($E7)>=0.01',{fill:'#FCE4D6',font:{color:'#9C0006',bold:true}});
note('Checks',27,'Reconciliation tolerance: 0.01 million / per share. Review may be expected for a scenario outside guidance.');
note('Checks',28,'Source values stay fixed when scenarios change. Sensitivity uses the same cash flow and dilution rules.');
ss.Checks.getRange('F7:F24').format.horizontalAlignment='center';
for(const n of names)v(n,'C2',`AKAM ${n==='Summary'?'Valuation Model':n}`);
// Validate the same live build, missing inputs, and sensitivity center for all scenarios.
const read=(s,a)=>ss[s].getRange(a).values[0][0];
const snapshots=[];for(const c of [1,2,3]){v('Assumptions','E4',c);wb.recalculate();snapshots.push({case:c,dcf:read('Valuation','E37'),pe:read('Valuation','J32'),rev2026:read('Forecast','H11'),ufcf2027:read('Forecast','I27')});if(Math.abs(read('Checks','E16'))>.0001)throw Error('Sensitivity center mismatch');}
v('Assumptions','E4',1);wb.recalculate();const baseline=read('Valuation','E37');v('Assumptions','J23',.30);wb.recalculate();if(read('Valuation','E37')===baseline)throw Error('Driver did not recalculate');v('Assumptions','J23',.20);
v('Assumptions','J24',null);wb.recalculate();if(typeof read('Valuation','E37')!=='number')throw Error('Inactive blank blocked Base');v('Assumptions','E4',2);wb.recalculate();if(typeof read('Valuation','E37')==='number')throw Error('Missing active input masked');v('Assumptions','J24',.15);v('Assumptions','E4',1);wb.recalculate();
for(let r=7;r<=24;r++)if(read('Checks','F'+r)!=='OK')throw Error('Base check failed: '+r+' '+read('Checks','E'+r));
// Independent numeric DCF reconstruction from the forecast cash flows.
const w=read('Assumptions','E64'),g=read('Assumptions','E65');let pv=0;for(const c of ['E','F','G','H','I','J'])pv+=read('Valuation',c+'11')/(1+w)**read('Valuation',c+'8');const rev=read('Forecast','M11');const termRev=rev*(1+g);const terminal=termRev*((read('Assumptions','J27')-read('Assumptions','J32'))*(1-read('Assumptions','J57'))+read('Assumptions','J42')-read('Assumptions','J47')-read('Assumptions','J37'))-rev*g*read('Assumptions','J52');const eq=pv+terminal/(w-g)/(1+w)**read('Valuation','J8')+read('Valuation','E32')-7640;let independent=eq/150;let nw=0,np=0;for(let r=50;r<=54;r++){nw+=read('Actuals','F'+r);np+=read('Actuals','F'+r)*read('Actuals','G'+r);independent=Math.min(independent,(eq+np)/(150+nw));}if(Math.abs(independent-baseline)>1e-7)throw Error('Independent DCF mismatch');
console.log(JSON.stringify({snapshots,independent,base:baseline,checks:checks.length}));
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},summary:'Final formula error scan'})).ndjson);
console.log((await wb.inspect({kind:'table',range:'Summary!C7:E17',include:'values,formulas',tableMaxRows:12,tableMaxCols:3,maxChars:2500})).ndjson);
await fs.mkdir(out,{recursive:true});
await (await SpreadsheetFile.exportXlsx(wb)).save(out+'/AKAM_Valuation_Model.xlsx');
for(const [n,range] of [['Summary','C2:M36'],['Assumptions','C2:J35'],['Assumptions','C36:J81'],['Forecast','C2:M48'],['Valuation','C2:L39'],['Valuation','C42:K56'],['Actuals','C2:I30'],['Actuals','C31:I58'],['Checks','C2:G28']]){const blob=await wb.render({sheetName:n,range,scale:1.3});await fs.writeFile(out+`/${n}_${range.replace(':','-')}.png`,new Uint8Array(await blob.arrayBuffer()));}
const reopened=await SpreadsheetFile.importXlsx(await FileBlob.load(out+'/AKAM_Valuation_Model.xlsx'));reopened.recalculate();const reopenedBase=reopened.worksheets.getItem('Valuation').getRange('E37').values[0][0];if(Math.abs(reopenedBase-baseline)>1e-7)throw Error('Reopened mismatch');reopened.worksheets.getItem('Assumptions').getRange('E4').values=[[2]];reopened.recalculate();const reopenedBear=reopened.worksheets.getItem('Valuation').getRange('E37').values[0][0];if(Math.abs(reopenedBear-snapshots[1].dcf)>1e-7)throw Error('Reopened Bear mismatch');
await fs.writeFile(out+'/audit.json',JSON.stringify({snapshots,independent,reopenedBase,reopenedBear,checks:checks.length},null,2));console.log(JSON.stringify({reopenedBase,reopenedBear,output:out+'/AKAM_Valuation_Model.xlsx'}));
