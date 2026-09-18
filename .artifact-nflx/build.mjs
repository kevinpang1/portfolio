import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const out='/Users/kevin/Documents/Codex/New project/outputs/01a0a62b-447e-7231-ad76-34e8211d28cd';
const w=Workbook.create();
const names=['Summary','Assumptions','Forecast','Valuation','Actuals','Checks'];
const sh=Object.fromEntries(names.map(n=>[n,w.worksheets.add(n)]));
const num='#,##0.0;(#,##0.0);"–"',pct='0.0%;(0.0%);"–"',price='$0.00;($0.00);"–"',mult='0.0"x"';
function v(s,a,x){sh[s].getRange(a).values=[[x]];if(typeof x==='number')sh[s].getRange(a).format.font.color='#0000FF';}
function f(s,a,x){sh[s].getRange(a).formulas=[[x]];sh[s].getRange(a).format.font.color=s==='Summary'?'#202020':x.includes('!')?'#008000':'#202020';}
function row(s,r,label,vals,start=4){v(s,'C'+r,label);if(vals)sh[s].getRangeByIndexes(r-1,start,1,vals.length).values=[vals];}
function band(s,r,label,end='L'){v(s,'C'+r,label);sh[s].getRange(`C${r}:${end}${r}`).format.fill='#ECEDEF';sh[s].getRange(`C${r}:${end}${r}`).format.font.bold=true;}
function total(s,r,end='L'){sh[s].getRange(`C${r}:${end}${r}`).format.borders={top:{style:'thin',color:'#A8A8A8'}};sh[s].getRange(`C${r}:${end}${r}`).format.font.bold=true;}
function note(s,r,t){v(s,'C'+r,t);sh[s].getRange(`C${r}:L${r}`).format.font.color='#666666';}
function fmt(s,a,code){sh[s].getRange(a).setNumberFormat(code);}
for(const n of names){let s=sh[n];s.showGridLines=false;s.getRange('A1:N95').format.font.name='Arial';s.getRange('A1:N95').format.font.size=10;s.getRange('A1:N95').format.rowHeight=21;s.getRange('A1:B95').format.columnWidth=2.5;s.getRange('C1:C95').format.columnWidth=43;s.getRange('D1:D95').format.columnWidth=3;s.getRange('E1:L95').format.columnWidth=13;s.getRange('M1:M95').format.columnWidth=3;s.getRange('N1:N95').format.columnWidth=80;s.getRange('C1:N95').format.verticalAlignment='center';s.getRange('E1:L95').setNumberFormat(num);v(n,'C2',n==='Summary'?'Netflix valuation':n);s.getRange('C2').format.font.size=16;s.getRange('C2:L2').format.borders={bottom:{style:'thin',color:'#E50914'}};v(n,'C4','Case Selected:');if(n!=='Assumptions')f(n,'E4','=Assumptions!E4');s.getRange('E4').format.borders={preset:'outside',style:'dashed',color:'#A0A0A0'};s.getRange('E4').format.horizontalAlignment='center';s.getRange('E4').format.font.color='#008000';if(n!=='Summary')s.freezePanes.freezeRows(7);}
sh.Summary.tabColor='#B20710';sh.Assumptions.tabColor='#E50914';sh.Actuals.tabColor='#E6D8C8';
const k='https://www.sec.gov/Archives/edgar/data/1065280/000106528026000034/nflx-20251231.htm';
const q='https://www.sec.gov/Archives/edgar/data/1065280/000106528026000212/nflx-20260630.htm';
const letter='https://www.sec.gov/Archives/edgar/data/1065280/000106528026000211/ex991_q226.htm';
note('Actuals',6,'USD millions; shares in millions. Cash outflows are positive except signed liability changes.');
row('Actuals',7,'Reported financial data',['2024A','2025A','H1 2026A','Q2 2026A']);
const actual=[
 [8,'Revenue',39000.966,45183.036,24809.695,12559.938],
 [9,'Operating income',10417.614,13326.603,8149.607,4192.610],
 [10,'Net income',8711.631,10981.201,8684.205,3401.414],
 [11,'Interest expense',718.733,776.510,437.762,175.685],
 [12,'Income tax provision',1254.026,1741.351,1931.467,667.172],
 [13,'Pretax income',9965.657,12722.552,10615.672,4068.586],
 [14,'Content amortization',15301.517,16422.166,8529.209,4311.309],
 [15,'Content asset additions (cash flow)',16223.617,17096.617,9774.440,4927.523],
 [16,'Content liability change (signed)',-779.135,-610.838,-136.578,-181.794],
 [17,'Other depreciation and amortization',328.914,333.389,199.105,100.530],
 [18,'Stock compensation',272.588,368.449,271.717,131.312],
 [19,'Operating cash flow',7361.364,10149.273,7034.017,1743.812],
 [20,'Property and equipment purchases',439.538,688.220,414.774,218.644],
 [21,'Weighted-average diluted shares',4392.608,4343.863,4279.776,4261.300],
 [22,'Weighted-average basic shares',4295.191,4249.512,4205.952,4189.303]];
for(const [r,l,...vs]of actual){row('Actuals',r,l,vs);sh.Actuals.getRange(`E${r}:H${r}`).format.font.color='#0000FF';}
v('Actuals','N8','2025 10-K, pp. 39–41 (2024 and 2025): '+k);v('Actuals','N9','Q2 2026 10-Q, pp. 3–6 (H1 and Q2): '+q);
band('Actuals',25,'Balance sheet and market inputs','H');
const inputs=[[26,'Cash and cash equivalents',9099.232,'June 30, 2026, 10-Q p. 6'],[27,'Short-term investments',28.678,'June 30, 2026, 10-Q p. 6'],[28,'Gross debt (rounded principal)',14400,'June 30, 2026, shareholder letter p. 5'],[29,'Period-end basic shares',4163.939676,'June 30, 2026, 10-Q p. 6; split-adjusted'],[30,'Market price',80.33,'September 14, 2026 close; ChartExchange'],[31,'Price / valuation date',46279,'September 14, 2026'],[32,'Balance sheet date',46203,'June 30, 2026'],[33,'FY2026 revenue guidance low',51000,'July 16, 2026 shareholder letter p. 2'],[34,'FY2026 revenue guidance high',51400,'July 16, 2026 shareholder letter p. 2'],[35,'FY2026 operating margin guidance',.315,'July 16, 2026 shareholder letter p. 2'],[36,'FY2026 advertising revenue outlook',3000,'July 16, 2026 shareholder letter p. 5'],[37,'FY2026 reported FCF guidance',12500,'July 16, 2026 shareholder letter p. 5; includes fee'],[38,'FY2026 content amortization growth',.1,'July 16, 2026 shareholder letter p. 2'],[39,'Cash content spend / amortization',1.1,'July 16, 2026 shareholder letter p. 5'],[40,'WBD termination fee (pretax)',2800,'Q1 2026; 10-Q Note 6'],[41,'Q3 2026 revenue guidance',12860,'July 16, 2026 shareholder letter p. 1']];
for(const [r,l,x,n]of inputs){row('Actuals',r,l,[x]);v('Actuals','G'+r,n);}
// Store dates as Excel serials derived from explicit UTC dates.
const serial=d=>(Date.parse(d)-Date.UTC(1899,11,30))/86400000;
v('Actuals','E31',serial('2026-09-14'));v('Actuals','E32',serial('2026-06-30'));fmt('Actuals','E31:E32','mmm d, yyyy');fmt('Actuals','E30',price);for(const r of [35,38])fmt('Actuals','E'+r,pct);fmt('Actuals','E39',mult);
v('Actuals','N26','Balance sheet: '+q);v('Actuals','N28','Guidance, gross debt and ads: '+letter);v('Actuals','N30','Market reference: https://chartexchange.com/symbol/nasdaq-nflx/historical/');v('Actuals','N31','Retrieved September 15, 2026; a second vendor showed $80.32 (one-cent difference).');v('Actuals','N40','Termination fee: '+q);
v('Assumptions','E4','Base');sh.Assumptions.getRange('E4').dataValidation={rule:{type:'list',values:['Base','Bear','Bull']}};sh.Assumptions.getRange('E4').format.fill='#FFF2CC';
note('Assumptions',6,'Blue numbers are editable. Active rows feed one forecast. 2027–2031 inputs are analyst assumptions.');
row('Assumptions',7,'Driver',['2024A','2025A','2026E','2027E','2028E','2029E','2030E','2031E']);
const groups=[
 [9,'Non-advertising revenue growth',[.105,.09,.08,.07,.06],[.06,.05,.04,.03,.03],[.14,.12,.11,.10,.08]],
 [14,'Advertising revenue growth',[.35,.30,.25,.20,.15],[.20,.15,.12,.10,.08],[.5,.4,.35,.3,.25]],
 [19,'Operating margin',[.33,.34,.345,.35,.355],[.30,.30,.30,.30,.30],[.345,.365,.38,.39,.40]],
 [24,'Content amortization growth',[.08,.07,.06,.05,.04],[.09,.08,.07,.06,.05],[.09,.08,.07,.06,.05]],
 [29,'Cash content / amortization',[1.1,1.09,1.08,1.06,1.05],[1.15,1.15,1.14,1.13,1.12],[1.08,1.07,1.05,1.04,1.03]],
 [34,'Tax rate',[.18,.18,.18,.18,.18],[.21,.21,.21,.21,.21],[.16,.16,.16,.16,.16]],
 [39,'Other D&A / revenue',[.008,.008,.008,.008,.008],[.008,.008,.008,.008,.008],[.008,.008,.008,.008,.008]],
 [44,'Non-content capex / revenue',[.016,.016,.016,.016,.016],[.019,.019,.019,.019,.019],[.015,.015,.015,.015,.015]],
 [49,'Other operating WC cash use / revenue',[.02,.02,.02,.02,.02],[.025,.025,.025,.025,.025],[.015,.015,.015,.015,.015]],
 [54,'Stock compensation / revenue',[.011,.011,.011,.011,.011],[.012,.012,.012,.012,.012],[.01,.01,.01,.01,.01]]];
const cols=['G','H','I','J','K','L'];
for(const [r,l,b,bear,bull]of groups){band('Assumptions',r,l);for(let j=0;j<3;j++){v('Assumptions','C'+(r+j+1),['Base','Bear','Bull'][j]);let arr=[b,bear,bull][j];for(let i=1;i<6;i++)v('Assumptions',cols[i]+(r+j+1),arr[i-1]);}for(let i=0;i<6;i++)f('Assumptions',cols[i]+r,`=IF($E$4="Base",IF(ISNUMBER(${cols[i]}${r+1}),${cols[i]}${r+1},NA()),IF($E$4="Bear",IF(ISNUMBER(${cols[i]}${r+2}),${cols[i]}${r+2},NA()),IF($E$4="Bull",IF(ISNUMBER(${cols[i]}${r+3}),${cols[i]}${r+3},NA()),NA())))`);fmt('Assumptions',`E${r}:L${r+3}`,r===29?mult:pct);sh.Assumptions.getRange(`C${r}:L${r+3}`).format.font.italic=true;}
for(const r of [10,11,12,15,16,17])v('Assumptions','G'+r,0);
for(const [r,b,be,bu]of [[20,.315,.305,.32],[25,.1,.1,.1],[30,1.1,1.13,1.08],[35,.18,.21,.16],[40,.008,.008,.008],[45,.016,.019,.015],[50,.02,.025,.015],[55,.011,.012,.01]]){[b,be,bu].forEach((x,j)=>v('Assumptions','G'+(r+j),x));}
f('Assumptions','G20','=Actuals!E35');f('Assumptions','G25','=Actuals!E38');f('Assumptions','G30','=Actuals!E39');
for(const [r,fr]of [[19,11],[24,18],[29,22],[34,15],[39,25],[44,27],[54,31]])for(const c of ['E','F'])f('Assumptions',c+r,`=Forecast!${c}${fr}`);
v('Assumptions','N9','2026 non-ad revenue = total revenue less ads; 2026 growth rows are unused.');v('Assumptions','N24','Content spending is explicit. Cash ratio includes payment timing and content liabilities.');v('Assumptions','N49','Positive = cash use. Includes non-content operating assets/liabilities; excludes content.');
band('Assumptions',60,'Valuation assumptions');row('Assumptions',61,'Metric',['Active','Base','Bear','Bull']);
const scalar=[[62,'2026 revenue',51200,50500,51700],[63,'2026 advertising revenue',3000,2700,3300],[64,'WACC',.10,.115,.09],[65,'Terminal growth',.03,.025,.035],[66,'2027 forward P/E',28,22,34],[67,'2027 forward EV/EBIT',22,17,27],[68,'Annual recurring interest expense',702.74,702.74,702.74]];
for(const [r,l,...vs]of scalar){row('Assumptions',r,l,null);vs.forEach((x,j)=>v('Assumptions',String.fromCharCode(70+j)+r,x));f('Assumptions','E'+r,`=IF($E$4="Base",F${r},IF($E$4="Bear",G${r},IF($E$4="Bull",H${r},NA())))`);}f('Assumptions','F62','=AVERAGE(Actuals!E33:E34)');f('Assumptions','F63','=Actuals!E36');for(const c of ['F','G','H'])f('Assumptions',c+'68','=Actuals!H11*4');fmt('Assumptions','E64:H65',pct);fmt('Assumptions','E66:H67',mult);
band('Assumptions',71,'Interim capital bridge assumptions');row('Assumptions',72,'Buybacks since June 30 ($m)',[0]);row('Assumptions',73,'Average buyback price',null);f('Assumptions','E73','=Actuals!E30');fmt('Assumptions','E73',price);row('Assumptions',74,'Other net cash inflow since June 30',[0]);
v('Assumptions','N64','WACC and terminal growth are selected assumptions, not observed market estimates.');v('Assumptions','N66','Selected multiples are analyst judgments; no peer-median claim.');v('Assumptions','N72','Unreported activity defaults to zero. Update buybacks and other cash flows when known.');v('Assumptions','N74','Exclude borrowings unless gross debt is also updated. No future M&A is forecast.');
note('Forecast',6,'USD millions, except per share. E = estimates. Stock compensation remains an expense in FCFF.');row('Forecast',7,'Operating forecast',['2024A','2025A','2026E','2027E','2028E','2029E','2030E','2031E']);
const labels={8:'Non-advertising revenue',9:'Advertising revenue',10:'Total revenue',11:'Operating margin',12:'Operating income',13:'Recurring interest expense',14:'Pretax income, recurring',15:'Tax rate',16:'Net income, recurring',17:'Content amortization',18:'Content amortization growth',19:'Content asset additions',20:'Content liability change',21:'Cash content spending',22:'Cash content / amortization',23:'Content cash spend above amortization',24:'Other depreciation and amortization',25:'Other D&A / revenue',26:'Non-content capex',27:'Non-content capex / revenue',28:'Other operating WC cash use',29:'WC cash use / revenue',30:'Stock compensation',31:'Stock compensation / revenue',32:'NOPAT',33:'Economic FCFF',34:'Economic cash flow to equity',35:'Reported-style recurring FCF',36:'Reported operating cash flow',37:'Reported FCF',38:'Diluted shares',39:'Recurring EPS',40:'Revenue growth'};
for(const [r,l]of Object.entries(labels))row('Forecast',+r,l);
for(const c of ['E','F']){for(const [r,sr]of [[10,8],[12,9],[13,11],[14,13],[16,10],[17,14],[19,15],[20,16],[24,17],[26,20],[30,18],[36,19],[38,21]])f('Forecast',c+r,`=Actuals!${c}${sr}`);for(const [r,eq]of [[11,'12/10'],[15,'TAX'],[21,'19-20'],[22,'21/17'],[23,'21-17'],[25,'24/10'],[27,'26/10'],[31,'30/10'],[37,'36-26'],[39,'16/38']])f('Forecast',c+r,eq==='TAX'?`=Actuals!${c}12/${c}14`:'='+eq.replace(/\d+/g,x=>c+x));}
f('Forecast','F18','=F17/E17-1');f('Forecast','F40','=F10/E10-1');
for(let i=0;i<6;i++){let c=cols[i],p=String.fromCharCode(c.charCodeAt(0)-1);f('Forecast',c+'8',i===0?'=Assumptions!E62-Assumptions!E63':`=${p}8*(1+Assumptions!${c}9)`);f('Forecast',c+'9',i===0?'=Assumptions!E63':`=${p}9*(1+Assumptions!${c}14)`);f('Forecast',c+'10',`=SUM(${c}8:${c}9)`);for(const[r,ar]of [[11,19],[15,34],[18,24],[22,29],[25,39],[27,44],[29,49],[31,54]])f('Forecast',c+r,`=Assumptions!${c}${ar}`);f('Forecast',c+'13','=Assumptions!$E$68');const forms={12:`${c}10*${c}11`,14:`${c}12-${c}13`,16:`${c}14*(1-${c}15)`,17:`${p}17*(1+${c}18)`,21:`${c}17*${c}22`,23:`${c}21-${c}17`,24:`${c}10*${c}25`,26:`${c}10*${c}27`,28:`${c}10*${c}29`,30:`${c}10*${c}31`,32:`${c}12*(1-${c}15)`,33:`${c}32-${c}23+${c}24-${c}26-${c}28`,34:`${c}33-${c}13*(1-${c}15)`,35:`${c}34+${c}30`,39:`${c}16/${c}38`,40:`${c}10/${p}10-1`};for(const[r,x]of Object.entries(forms))f('Forecast',c+r,'='+x);f('Forecast',c+'38','=Valuation!$E$31');}
for(const r of [11,15,18,25,27,29,31,40]){fmt('Forecast',`E${r}:L${r}`,pct);sh.Forecast.getRange(`C${r}:L${r}`).format.font.italic=true;}fmt('Forecast','E22:L22',mult);fmt('Forecast','E39:L39',price);for(const r of [10,12,16,21,33,35,37,39])total('Forecast',r);
note('Forecast',43,'FCFF = NOPAT + content amortization − cash content spending + other D&A − capex − operating WC cash use.');
note('Forecast',44,'Historical net income and EPS are reported; forecast earnings exclude the termination fee and interest income.');
note('Forecast',45,'Future shares are constant after the interim bridge. SBC is expensed, with no FCFF add-back or future dilution.');
note('Forecast',46,'Reported-style FCF adds SBC back for comparison only. Other noncash and cash-tax timing differences are not modeled.');
band('Forecast',49,'2026 guidance and second-half requirements');
for(const[r,l]of [[50,'H2 revenue required'],[51,'Q3 revenue guidance'],[52,'Implied Q4 revenue'],[53,'H2 operating income required'],[54,'H2 operating margin required'],[56,'Reported FCF guidance'],[57,'Fee, estimated after-tax contribution'],[58,'Underlying FCF guidance proxy'],[59,'Model reported-style recurring FCF'],[60,'Model less underlying guidance proxy']])row('Forecast',r,l);
f('Forecast','G50','=G10-Actuals!G8');f('Forecast','G51','=Actuals!E41');f('Forecast','G52','=G50-G51');f('Forecast','G53','=G12-Actuals!G9');f('Forecast','G54','=G53/G50');fmt('Forecast','G54',pct);f('Forecast','G56','=Actuals!E37');f('Forecast','G57','=Actuals!E40*(1-G15)');f('Forecast','G58','=G56-G57');f('Forecast','G59','=G35');f('Forecast','G60','=G59-G58');note('Forecast',62,'Fee tax effect uses the selected tax rate as an estimate; it is not a company-disclosed cash tax allocation.');
note('Valuation',6,'Spot value on September 14, 2026. Calendar-year end discounting; current-year cash flows prorated by days.');
row('Valuation',7,'DCF cash flows',['2026 stub','2027E','2028E','2029E','2030E','2031E']);
for(const[r,l]of [[8,'Year-end cash flow date'],[9,'Years to cash flow'],[10,'Cash flow fraction'],[11,'Annual economic FCFF'],[12,'FCFF in valuation period'],[13,'Discount factor'],[14,'Present value of FCFF']])row('Valuation',r,l);
for(let i=0;i<6;i++){let c=String.fromCharCode(69+i),fc=cols[i];f('Valuation',c+'8',`=DATE(${2026+i},12,31)`);f('Valuation',c+'9',`=(${c}8-Actuals!$E$31)/365`);f('Valuation',c+'10',i===0?'=E9':'=1');f('Valuation',c+'11',`=Forecast!${fc}33`);f('Valuation',c+'12',`=${c}11*${c}10`);f('Valuation',c+'13',`=1/(1+Assumptions!$E$64)^${c}9`);f('Valuation',c+'14',`=${c}12*${c}13`);}fmt('Valuation','E8:J8','mmm d, yyyy');fmt('Valuation','E10:J10',pct);fmt('Valuation','E13:J13','0.000x');total('Valuation',14,'J');
const valLabels={17:'Terminal growth',18:'WACC',19:'2032 terminal FCFF',20:'Terminal enterprise value',21:'PV of terminal value',22:'PV of explicit cash flows',23:'Enterprise value',26:'June 30 cash + short-term investments',27:'Estimated interim economic equity FCF',28:'Interim buybacks',29:'Other interim cash inflow',30:'Estimated cash at valuation date',31:'Estimated diluted shares',32:'Gross debt',33:'Net debt',34:'Equity value',35:'DCF value / share',36:'DCF upside / downside',37:'Terminal value / enterprise value',40:'2027 selected P/E',41:'2027 recurring EPS',42:'P/E implied value / share',43:'2027 selected EV/EBIT',44:'2027 operating income',45:'EV/EBIT implied enterprise value',46:'EV/EBIT implied value / share',49:'Market price',50:'Implied 2027 P/E',51:'Implied 2027 EV/EBIT',52:'2027 economic FCF yield'};
for(const[r,l]of Object.entries(valLabels))row('Valuation',+r,l);
const vf={17:'Assumptions!E65',18:'Assumptions!E64',19:'J11*(1+E17)',20:'IF(E18>E17,E19/(E18-E17),NA())',21:'E20*J13',22:'SUM(E14:J14)',23:'SUM(E21:E22)',26:'SUM(Actuals!E26:E27)',27:'Forecast!G34*(Actuals!E31-Actuals!E32)/365',28:'Assumptions!E72',29:'Assumptions!E74',30:'E26+E27-E28+E29',31:'Actuals!E29+Actuals!H21-Actuals!H22-E28/Assumptions!E73',32:'Actuals!E28',33:'E32-E30',34:'E23-E33',35:'E34/E31',36:'E35/E49-1',37:'E21/E23',40:'Assumptions!E66',41:'Forecast!H39',42:'E40*E41',43:'Assumptions!E67',44:'Forecast!H12',45:'E43*E44',46:'(E45-E33)/E31',49:'Actuals!E30',50:'E49/E41',51:'(E49*E31+E33)/E44',52:'Forecast!H34/(E49*E31)'};for(const[r,x]of Object.entries(vf))f('Valuation','E'+r,'='+x);for(const r of [17,18,36,37,52])fmt('Valuation','E'+r,pct);for(const r of [35,41,42,46,49])fmt('Valuation','E'+r,price);for(const r of [40,43,50,51])fmt('Valuation','E'+r,mult);for(const r of [23,30,33,34,35,42,46])total('Valuation',r,'J');
v('Valuation','G26','June 30 cash includes the fee receipt. No second fee add-back.');v('Valuation','G27','Uniform recurring cash generation since June 30 is an approximation.');v('Valuation','G31','Period-end basic shares + Q2 incremental dilution − interim repurchases.');v('Valuation','G32','Rounded gross principal; leases and content liabilities stay in operating cash flows.');v('Valuation','G40','Spot forward multiples applied to calendar 2027 earnings.');
band('Valuation',55,'DCF sensitivity: WACC and terminal growth','J');
v('Valuation','E56','WACC / g');for(let j=0;j<5;j++){let c=String.fromCharCode(70+j);f('Valuation',c+'56',`=$E$17+${(j-2)*.005}`);fmt('Valuation',c+'56',pct);}for(let i=0;i<5;i++){let r=57+i;f('Valuation','E'+r,`=$E$18+${(i-2)*.01}`);fmt('Valuation','E'+r,pct);for(let j=0;j<5;j++){let c=String.fromCharCode(70+j);let pv=Array.from({length:6},(_,n)=>{let cc=String.fromCharCode(69+n);return `$${cc}$12/(1+$E${r})^$${cc}$9`;}).join('+');f('Valuation',c+r,`=IF($E${r}>${c}$56,(${pv}+$J$11*(1+${c}$56)/($E${r}-${c}$56)/(1+$E${r})^$J$9-$E$33)/$E$31,NA())`);fmt('Valuation',c+r,price);}}sh.Valuation.getRange('F57:J61').conditionalFormats.add('colorScale',{colors:['#FADBD8','#FFFFFF','#C9DCD0'],thresholds:['min','50%','max']});
note('Valuation',64,'Terminal FCFF grows at g with final-year margin and cash conversion retained. Sensitivity re-discounts every cash flow.');
note('Summary',6,'USD millions, except per-share values. Market reference: September 14, 2026 close.');
row('Summary',8,'Valuation method',['Value / share','Vs. market']);
for(const[r,l,vr]of [[9,'DCF',35],[10,'2027 P/E cross-check',42],[11,'2027 EV/EBIT cross-check',46]]){row('Summary',r,l);f('Summary','E'+r,`=Valuation!E${vr}`);f('Summary','F'+r,`=E${r}/Valuation!E49-1`);}fmt('Summary','E9:E11',price);fmt('Summary','F9:F11',pct);total('Summary',9,'F');row('Summary',13,'Market price');f('Summary','E13','=Valuation!E49');fmt('Summary','E13',price);row('Summary',14,'Valuation date');f('Summary','E14','=Actuals!E31');fmt('Summary','E14','mmm d, yyyy');
band('Summary',17,'Operating outlook');row('Summary',18,'Metric',['2025A','2026E','2027E','2028E','2029E','2030E','2031E']);for(const[r,l,fr]of [[19,'Revenue',10],[20,'Operating margin',11],[21,'Operating income',12],[22,'Economic FCFF',33],[23,'Recurring EPS',39]]){row('Summary',r,l);for(let i=0;i<7;i++)f('Summary',String.fromCharCode(69+i)+r,`=Forecast!${String.fromCharCode(70+i)}${fr}`);}fmt('Summary','E20:K20',pct);fmt('Summary','E23:K23',price);
band('Summary',26,'Key valuation drivers','K');for(const[r,l,x,fo]of [[27,'WACC','Assumptions!E64',pct],[28,'Terminal growth','Assumptions!E65',pct],[29,'Terminal value share of enterprise value','Valuation!E37',pct],[30,'Implied market 2027 P/E','Valuation!E50',mult],[31,'Implied market 2027 EV/EBIT','Valuation!E51',mult]]){row('Summary',r,l);f('Summary','E'+r,'='+x);fmt('Summary','E'+r,fo);}
note('Summary',34,'Change the case and blue inputs on Assumptions. All forecast and valuation outputs recalculate.');note('Summary',35,'DCF is the primary estimate. Selected multiples are cross-check assumptions, not observed peer medians.');note('Summary',36,'Interim cash generation is estimated. Unreported buybacks default to zero; update the capital bridge when known.');note('Summary',37,'Future FCFF excludes the WBD fee and retains SBC expense. Content spending and payment timing drive cash conversion.');note('Summary',38,'Advertising is separate from non-ad revenue. No undisclosed subscriber count or ARPU is assumed.');
const ch=sh.Summary.charts.add('bar',[sh.Summary.getRange('C8:C11'),sh.Summary.getRange('E8:E11')]);ch.title='Value per share';ch.setPosition('H8','L16');ch.series.items[0].fill='#B20710';ch.titleTextStyle.fontSize=12;ch.titleTextStyle.typeface='Arial';ch.yAxis={numberFormatCode:'$0',numberFormatSourceLinked:false};
note('Checks',6,'Independent diagnostics. Differences use a $0.01m tolerance unless specified.');row('Checks',7,'Check',['Difference / value','Result']);
const checks=[
 ['2025 income reconciliation','=Actuals!F13-Actuals!F12-Actuals!F10','ABS(E8)<0.01'],
 ['2025 reported FCF reconciliation','=Forecast!F37-(Actuals!F19-Actuals!F20)','ABS(E9)<0.01'],
 ['2026 revenue components','=Forecast!G10-Assumptions!E62','ABS(E10)<0.01'],
 ['2026 H2 revenue is nonnegative','=Forecast!G50','E11>=0'],
 ['2026 implied Q4 revenue is nonnegative','=Forecast!G52','E12>=0'],
 ['2026 H2 EBIT is nonnegative','=Forecast!G53','E13>=0'],
 ['WACC exceeds terminal growth','=Assumptions!E64-Assumptions!E65','E14>0'],
 ['Sensitivity center ties to DCF','=Valuation!H59-Valuation!E35','ABS(E15)<0.000001'],
 ['Estimated shares are positive','=Valuation!E31','E16>0'],
 ['Estimated cash is nonnegative','=Valuation!E30','E17>=0'],
 ['2026 cash bridge conservation','=Valuation!E30-(SUM(Actuals!E26:E27)+Valuation!E27-Assumptions!E72+Assumptions!E74)','ABS(E18)<0.01'],
 ['FCFF to reported-style FCF bridge','=Forecast!G35-(Forecast!G33-Forecast!G13*(1-Forecast!G15)+Forecast!G30)','ABS(E19)<0.01'],
 ['Date falls within 2026 after June 30','=Actuals!E31','AND(E20>=Actuals!E32,E20<DATE(2027,1,1))']];
checks.forEach(([l,x,cond],i)=>{let r=i+8;row('Checks',r,l);f('Checks','E'+r,x);f('Checks','F'+r,`=IF(${cond},"OK","REVIEW")`);});fmt('Checks','E8:E20','0.00;(0.00);0.00');sh.Checks.getRange('E8:E20').format.font.color='#202020';sh.Checks.getRange('F8:F20').conditionalFormats.add('containsText',{text:'REVIEW',format:{fill:'#FCE4D6',font:{bold:true,color:'#B20710'}}});
note('Checks',23,'The guidance gap is diagnostic, not an accounting error. The forecast is independently built.');note('Checks',24,'Native Excel recalculation has not been tested; formula behavior is tested in the workbook calculation engine.');
for(const n of names){for(const r of (n==='Assumptions'?[7,61]:n==='Summary'?[8,18]:[7])){sh[n].getRange(`C${r}:L${r}`).format.fill='#272727';sh[n].getRange(`C${r}:L${r}`).format.font.color='#FFFFFF';sh[n].getRange(`E${r}:L${r}`).format.horizontalAlignment='center';}sh[n].getRange('D1:D95').format.fill='#FFFFFF';}
// Presentation fixes after visual inspection.
v('Summary','E22','n.a.');v('Assumptions','E24','n.a.');
v('Summary','C23','EPS (historical reported)');v('Forecast','C39','EPS (historical reported)');
v('Forecast','C13','Interest expense');v('Forecast','C14','Pretax income');v('Forecast','C16','Net income');
fmt('Assumptions','E29:L32','0.00"x"');fmt('Forecast','E22:L22','0.00"x"');
sh.Checks.getRange('F8:F20').format.horizontalAlignment='center';fmt('Checks','E20','mmm d, yyyy');
sh.Summary.getRange('G8').format.fill='#FFFFFF';
for(const n of names)sh[n].getRange('E4').format.font.color='#008000';
v('Valuation','G32','Gross principal. Content and leases are operating costs.');
sh.Actuals.getRange('E26:E41').format.font.color='#0000FF';
sh.Assumptions.getRange('E72').format.font.color='#0000FF';sh.Assumptions.getRange('E74').format.font.color='#0000FF';
// Behavior tests run against the same build; restore the original case and inputs.
let results={};for(const cs of ['Bear','Base','Bull']){v('Assumptions','E4',cs);w.recalculate();results[cs]=sh.Valuation.getRange('E35:E36').values;}v('Assumptions','E4','Base');w.recalculate();const base=sh.Valuation.getRange('E35').values[0][0];v('Assumptions','L20',.365);w.recalculate();const changed=sh.Valuation.getRange('E35').values[0][0];if(!(changed>base))throw Error('Late-period margin did not affect DCF');v('Assumptions','L20',.355);
v('Assumptions','H11',null);w.recalculate();if(sh.Valuation.getRange('E35').values[0][0]!==base)throw Error('Unselected blank affects Base');v('Assumptions','E4','Bear');w.recalculate();const missing=sh.Valuation.getRange('E35').values[0][0];if(typeof missing==='number')throw Error('Selected blank silently becomes zero');v('Assumptions','H11',.06);v('Assumptions','E4','Base');
w.recalculate();console.log('CASE TESTS',JSON.stringify(results),'margin change',changed,'missing',missing);
console.log((await w.inspect({kind:'table',range:'Summary!C8:F14',include:'values,formulas',tableMaxRows:7,tableMaxCols:4})).ndjson);
console.log((await w.inspect({kind:'table',range:'Checks!C8:F20',include:'values',tableMaxRows:13,tableMaxCols:4})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},summary:'Final formula error scan'})).ndjson);
await fs.mkdir(out,{recursive:true});const file=await SpreadsheetFile.exportXlsx(w);await file.save(out+'/NFLX_Valuation.xlsx');
for(const [n,range,suffix]of [['Summary','C2:L38',''],['Assumptions','C24:L35',''],['Forecast','C2:L40',''],['Valuation','C2:L37',''],['Actuals','C25:L41',''],['Checks','C2:H24',''],['Valuation','C55:J64','_sensitivity'],['Assumptions','C60:L74','_capital'],['Forecast','C49:L62','_guidance']]){let im=await w.render({sheetName:n,range,scale:1.5,format:'png'});await fs.writeFile(out+'/'+n+suffix+'.png',new Uint8Array(await im.arrayBuffer()));}
await fs.writeFile(out+'/verification.json',JSON.stringify({results,base,changed},null,2));
