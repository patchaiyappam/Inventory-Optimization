import * as XLSX from "xlsx";

export type InventoryItem = {
  sku:string; name:string; category:string; supplier:string; supplierType:string; warehouse:string;
  purchaseDate:string; lastReceiptDate:string|null; lastIssueDate:string|null; lastMovementDate:string|null;
  quantity:number; unitCost:number; inventoryValue:number; avgMonthlyConsumption:number;
  annualConsumptionValue:number; leadTimeWeeks:number; moq:number; openPurchaseOrders:number|null;
  confirmedRequirementDate:string|null; abc:string; fsn:string; xyz:string; ageMonths:number;
  agingBucket:string; cv:number; carryingRate:number; annualCarryingCost:number;
  cumulativeCarryingCost:number; inventoryDays:number; turnoverRatio:number; capitalBlocked:number;
  moqExcessQuantity:number; moqExcessValue:number; risk:string; riskReasons:string[];
  recommendedAction:string; recoverableValueEstimate:number; recoveryRateAssumption:number; recoveryStatus:string;
  unavailableFields?:string[];
};

export type Transaction = {id:string;sku:string;date:string;type:string;quantity:number;unitCost:number|null;warehouse:string;reference:string|null;remarks:string};
export type RecoveryAction = {sku:string;decision:string;owner:string;deadline:string;status:string;note:string;actualProceeds?:number|null;actualCosts?:number|null;evidence?:string};
export type AppData = {
  meta:{datasetName:string;mode:string;sourceLabel:string;currency:string;generatedAt:string;limitations:string[];transactionSourceComplete?:boolean;transactionCoverageStart?:string;transactionCoverageEnd?:string};
  assumptions:Record<string,number|string>;items:InventoryItem[];transactions:Transaction[];recoveryActions:RecoveryAction[];recoveryResults:unknown[]
};

export type Filters = {search:string;category:string;supplier:string;warehouse:string;abc:string;fsn:string;xyz:string;aging:string;risk:string;valueMin:string;valueMax:string;costMin:string;costMax:string};
export const EMPTY_FILTERS:Filters={search:"",category:"All",supplier:"All",warehouse:"All",abc:"All",fsn:"All",xyz:"All",aging:"All",risk:"All",valueMin:"",valueMax:"",costMin:"",costMax:""};

export type ImportIssue={row:number;field:string;message:string;severity:"error"|"warning"};
export type ImportPreview={fileName:string;items:InventoryItem[];issues:ImportIssue[];rowCount:number;sheetName:string};
export type TransactionCoverage={sourceComplete:boolean;startMonth?:string;endMonth?:string};
export type MonthlyObservation={month:string;value:number|null;status:"observed"|"zero"|"missing"};
export type MonthlySeries={points:MonthlyObservation[];observedCount:number;missingMonths:string[];zeroFilledMonths:string[];sourceComplete:boolean};
export type ForecastPoint={month:string;actual:number|null;predicted:number;phase:"train"|"test"|"gap"|"future"};
export type ForecastResult={sku:string;observations:number;trainCount:number;testCount:number;horizon:number;mae:number;rmse:number;mape:number|null;points:ForecastPoint[];missingMonths:string[];zeroFilledMonths:string[];limitations:string[]};

export const SAFETY_TIME_RATIO=0.5;
export const ECONOMIC_REVIEW_THRESHOLD_RATIO=1;

const HEADER_ALIASES:Record<string,string[]>={
  sku:["sku","materialcode","itemcode","productcode","stockcode"],name:["name","materialname","itemname","description","productname"],
  category:["category","materialcategory","productcategory"],supplier:["supplier","vendor","suppliername","vendorname"],
  warehouse:["warehouse","location","storagelocation","plant"],quantity:["quantity","qty","stockquantity","onhand","onhandquantity","currentstock"],
  unitCost:["unitcost","cost","price","unitprice","standardcost"],purchaseDate:["purchasedate","procurementdate"],
  lastReceiptDate:["lastreceiptdate","receiptdate","latestreceiptdate"],lastIssueDate:["lastissuedate","issuedate","latestissuedate"],
  lastMovementDate:["lastmovementdate","movementdate"],avgMonthlyConsumption:["avgmonthlyconsumption","averagemonthlyconsumption","monthlydemand","averageusage"],
  annualConsumptionValue:["annualconsumptionvalue","annualusagevalue"],leadTimeWeeks:["leadtimeweeks","leadtime","leadtimeweek"],
  moq:["moq","minimumorderquantity"],openPurchaseOrders:["openpurchaseorders","openpo","openpoqty"],
  confirmedRequirementDate:["confirmedrequirementdate","requirementdate","needbydate"],abc:["abc","abcclass"],fsn:["fsn","fsnclass"],
  xyz:["xyz","xyzclass"],ageMonths:["agemonths","inventoryagemonths","age"],cv:["cv","coefficientofvariation"]
};

const key=(value:unknown)=>String(value??"").trim().toLowerCase().replace(/[^a-z0-9]/g,"");
const num=(value:unknown)=>{const cleaned=String(value??"").replace(/[$,%\s,]/g,"");if(!cleaned)return null;const n=Number(cleaned);return Number.isFinite(n)?n:null};
const iso=(value:unknown)=>{if(value==null||value==="")return null;if(typeof value==="number"){const d=XLSX.SSF.parse_date_code(value);return d?new Date(Date.UTC(d.y,d.m-1,d.d)).toISOString():null}const d=new Date(String(value));return Number.isNaN(d.getTime())?null:d.toISOString()};
const monthsBetween=(date:string|null,now:Date)=>date?Math.max(0,(now.getTime()-new Date(date).getTime())/(30.4375*86400000)):0;
const aging=(m:number)=>m<3?"Fresh":m<6?"Normal":m<12?"Aging":m<24?"Old":"Dead";
const pick=(row:Record<string,unknown>,field:string)=>{const aliases=HEADER_ALIASES[field]||[];const hit=Object.keys(row).find(k=>aliases.includes(key(k)));return hit?row[hit]:undefined};
const text=(value:unknown)=>String(value??"").trim();
const missing=(item:InventoryItem,label:string)=>(item.unavailableFields||[]).some(x=>x.toLowerCase()===label.toLowerCase());
const validMonth=(value?:string)=>!!value&&/^\d{4}-(0[1-9]|1[0-2])$/.test(value);

export function assessRisk(item:InventoryItem){
  const missingFields:string[]=[];
  if(!Number.isFinite(item.ageMonths)||missing(item,"inventory age"))missingFields.push("material age");
  if(!item.lastMovementDate&&!item.lastIssueDate&&!item.lastReceiptDate)missingFields.push("movement history");
  if(!["F","S","N"].includes(item.fsn)||missing(item,"FSN class"))missingFields.push("FSN classification");
  if(missingFields.length)return {risk:"Insufficient data",reasons:missingFields.map(x=>`Missing ${x}`),missingFields};
  const reasons:string[]=[];
  if(item.ageMonths>=24)reasons.push("Inventory age is at least 24 months");else if(item.ageMonths>=12)reasons.push("Inventory age is at least 12 months");
  if(item.fsn==="N")reasons.push("Classified as non-moving");
  if(item.turnoverRatio>0&&item.turnoverRatio<1)reasons.push("Turnover is below 1.0x");
  const risk=reasons.length>=3||item.ageMonths>=24&&item.fsn==="N"?"Critical":reasons.length>=2||item.ageMonths>=24?"High":reasons.length?"Medium":"Low";
  return {risk,reasons:reasons.length?reasons:["No configured review trigger"],missingFields};
}

export function stockCoverage(item:InventoryItem,weeksPerMonth=4.33){
  if(missing(item,"lead time")||missing(item,"average monthly consumption")||!Number.isFinite(item.leadTimeWeeks)||item.leadTimeWeeks<0||!Number.isFinite(item.avgMonthlyConsumption)||item.avgMonthlyConsumption<0)return null;
  const safetyTimeWeeks=item.leadTimeWeeks*SAFETY_TIME_RATIO,coverageWeeks=item.leadTimeWeeks+safetyTimeWeeks,targetQuantity=item.avgMonthlyConsumption*(coverageWeeks/weeksPerMonth);
  return {leadTimeWeeks:item.leadTimeWeeks,safetyTimeWeeks,coverageWeeks,targetQuantity};
}

export function potentiallyTrappedCapital(item:InventoryItem,weeksPerMonth=4.33){
  if(!item.confirmedRequirementDate||missing(item,"confirmed requirement date"))return null;const coverage=stockCoverage(item,weeksPerMonth);if(!coverage)return null;
  const excessQuantity=Math.max(0,item.quantity-coverage.targetQuantity);return {...coverage,excessQuantity,value:excessQuantity*item.unitCost};
}

export function projectedCarryingCost(item:InventoryItem,analysisDate:string,rate:number,thresholdRatio=ECONOMIC_REVIEW_THRESHOLD_RATIO){
  if(!item.confirmedRequirementDate||missing(item,"confirmed requirement date")||missing(item,"lead time")||!Number.isFinite(item.leadTimeWeeks)||item.leadTimeWeeks<0)return null;
  const requirement=new Date(item.confirmedRequirementDate),analysis=new Date(analysisDate);if(Number.isNaN(requirement.getTime())||Number.isNaN(analysis.getTime())||!Number.isFinite(rate)||rate<0)return null;
  const safetyTimeWeeks=item.leadTimeWeeks*SAFETY_TIME_RATIO,bufferDays=(item.leadTimeWeeks+safetyTimeWeeks)*7,holdingDays=Math.max(0,Math.floor((requirement.getTime()-analysis.getTime())/86400000-bufferDays)),cost=item.inventoryValue*rate*holdingDays/365,ratioToInventoryValue=item.inventoryValue>0?cost/item.inventoryValue:0;
  const reviewStatus=ratioToInventoryValue>=thresholdRatio?"Threshold crossed — investigate return, exchange, or transfer":"Below economic review threshold";
  return {holdingDays,cost,safetyTimeWeeks,bufferDays,ratioToInventoryValue,reviewStatus};
}

export function normalizeInventoryItems(items:InventoryItem[],weeksPerMonth=4.33,recoveryRate=0.4){
  return items.map(source=>{const item={...source,unavailableFields:[...(source.unavailableFields||[])]};if(!item.confirmedRequirementDate&&!missing(item,"confirmed requirement date"))item.unavailableFields!.push("confirmed requirement date");if(!item.lastMovementDate&&!item.lastIssueDate&&!item.lastReceiptDate&&!missing(item,"movement history"))item.unavailableFields!.push("movement history");const risk=assessRisk(item),trapped=potentiallyTrappedCapital(item,weeksPerMonth),trappedValue=trapped?.value??Number.NaN;return {...item,risk:risk.risk,riskReasons:risk.reasons,recommendedAction:risk.risk==="Low"?"Monitor":"Review evidence, demand, redeployment and supplier options",capitalBlocked:trappedValue,moqExcessQuantity:trapped?.excessQuantity??Number.NaN,moqExcessValue:trappedValue,recoverableValueEstimate:Number.isFinite(trappedValue)?trappedValue*recoveryRate:Number.NaN}});
}

export async function parseInventoryFile(file:File,carryingRate:number,recoveryRate:number):Promise<ImportPreview>{
  if(file.size>15*1024*1024)throw new Error("File exceeds the 15 MB local import limit.");const book=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true}),sheetName=book.SheetNames[0];if(!sheetName)throw new Error("No worksheet was found.");const rows=XLSX.utils.sheet_to_json<Record<string,unknown>>(book.Sheets[sheetName],{defval:"",raw:true});if(!rows.length)throw new Error("The first worksheet has no data rows.");if(rows.length>10000)throw new Error("Import is limited to 10,000 rows for this browser demonstration.");
  const issues:ImportIssue[]=[];const seen=new Set<string>();const items:InventoryItem[]=[];const now=new Date();
  rows.forEach((row,index)=>{const rowNo=index+2,sku=text(pick(row,"sku")),quantity=num(pick(row,"quantity")),unitCost=num(pick(row,"unitCost"));if(!sku)issues.push({row:rowNo,field:"SKU",message:"SKU is required.",severity:"error"});if(sku&&seen.has(sku.toLowerCase()))issues.push({row:rowNo,field:"SKU",message:"Duplicate SKU.",severity:"error"});if(quantity==null||quantity<0)issues.push({row:rowNo,field:"Quantity",message:"Quantity must be a non-negative number.",severity:"error"});if(unitCost==null||unitCost<0)issues.push({row:rowNo,field:"Unit cost",message:"Unit cost must be a non-negative number.",severity:"error"});if(issues.some(i=>i.row===rowNo&&i.severity==="error"))return;
    seen.add(sku.toLowerCase());const unavailable:string[]=[];const missingText=(field:string,label:string)=>{const v=text(pick(row,field));if(!v)unavailable.push(label);return v||"Unavailable"};const purchaseDate=iso(pick(row,"purchaseDate")),lastReceiptDate=iso(pick(row,"lastReceiptDate")),lastIssueDate=iso(pick(row,"lastIssueDate")),lastMovementDate=iso(pick(row,"lastMovementDate"));if(!purchaseDate&&!lastReceiptDate)unavailable.push("receipt/purchase date");if(!lastMovementDate&&!lastIssueDate&&!lastReceiptDate)unavailable.push("movement history");
    const suppliedAge=num(pick(row,"ageMonths")),ageValue=suppliedAge??monthsBetween(lastReceiptDate||purchaseDate,now);if(suppliedAge==null&&!purchaseDate&&!lastReceiptDate)unavailable.push("inventory age");const monthlyInput=num(pick(row,"avgMonthlyConsumption")),monthly=monthlyInput??0;if(monthlyInput==null)unavailable.push("average monthly consumption");const leadInput=num(pick(row,"leadTimeWeeks")),lead=leadInput??0;if(leadInput==null)unavailable.push("lead time");const requirement=iso(pick(row,"confirmedRequirementDate"));if(!requirement)unavailable.push("confirmed requirement date");
    const value=quantity!*unitCost!,annualConsumption=num(pick(row,"annualConsumptionValue"))??monthly*12*unitCost!,turnover=value>0?annualConsumption/value:0,fsn=text(pick(row,"fsn")).toUpperCase()||"Unavailable";if(fsn==="Unavailable")unavailable.push("FSN class");const item:InventoryItem={sku,name:missingText("name","material name"),category:missingText("category","category"),supplier:missingText("supplier","supplier"),supplierType:"Unavailable",warehouse:missingText("warehouse","warehouse"),purchaseDate:purchaseDate||"",lastReceiptDate,lastIssueDate,lastMovementDate,quantity:quantity!,unitCost:unitCost!,inventoryValue:value,avgMonthlyConsumption:monthly,annualConsumptionValue:annualConsumption,leadTimeWeeks:lead,moq:num(pick(row,"moq"))??0,openPurchaseOrders:num(pick(row,"openPurchaseOrders")),confirmedRequirementDate:requirement,abc:text(pick(row,"abc")).toUpperCase()||"Unavailable",fsn,xyz:text(pick(row,"xyz")).toUpperCase()||"Unavailable",ageMonths:ageValue,agingBucket:aging(ageValue),cv:num(pick(row,"cv"))??0,carryingRate,annualCarryingCost:value*carryingRate,cumulativeCarryingCost:ageValue>0?value*carryingRate*ageValue/12:0,inventoryDays:monthlyInput!=null&&monthly>0?quantity!/monthly*30.4375:0,turnoverRatio:turnover,capitalBlocked:0,moqExcessQuantity:0,moqExcessValue:0,risk:"Insufficient data",riskReasons:[],recommendedAction:"Review missing data",recoverableValueEstimate:0,recoveryRateAssumption:recoveryRate,recoveryStatus:"Not reviewed",unavailableFields:unavailable};["abc","xyz"].forEach(field=>{if((item as unknown as Record<string,string>)[field]==="Unavailable")unavailable.push(`${field.toUpperCase()} class`)});const normalized=normalizeInventoryItems([item],4.33,recoveryRate)[0];if(unavailable.length)issues.push({row:rowNo,field:"Optional fields",message:`Unavailable: ${[...new Set(unavailable)].join(", ")}.`,severity:"warning"});items.push(normalized);
  });return {fileName:file.name,items,issues,rowCount:rows.length,sheetName};
}

export function applyFilters(items:InventoryItem[],f:Filters){const minV=num(f.valueMin),maxV=num(f.valueMax),minC=num(f.costMin),maxC=num(f.costMax),query=f.search.trim().toLowerCase();return items.filter(i=>(!query||`${i.sku} ${i.name} ${i.supplier} ${i.category} ${i.warehouse}`.toLowerCase().includes(query))&&(f.category==="All"||i.category===f.category)&&(f.supplier==="All"||i.supplier===f.supplier)&&(f.warehouse==="All"||i.warehouse===f.warehouse)&&(f.abc==="All"||i.abc===f.abc)&&(f.fsn==="All"||i.fsn===f.fsn)&&(f.xyz==="All"||i.xyz===f.xyz)&&(f.aging==="All"||i.agingBucket===f.aging)&&(f.risk==="All"||i.risk===f.risk)&&(minV==null||i.inventoryValue>=minV)&&(maxV==null||i.inventoryValue<=maxV)&&(minC==null||i.annualCarryingCost>=minC)&&(maxC==null||i.annualCarryingCost<=maxC))}

export function pairedPearson(items:InventoryItem[],a:keyof InventoryItem,b:keyof InventoryItem){const labels:Partial<Record<keyof InventoryItem,string[]>>={ageMonths:["inventory age"],avgMonthlyConsumption:["average monthly consumption"],leadTimeWeeks:["lead time"],abc:["abc class"],fsn:["fsn class"],xyz:["xyz class"]};const pairs=items.flatMap(i=>{const av=Number(i[a]),bv=Number(i[b]),unavailable=(i.unavailableFields||[]).map(x=>x.toLowerCase()),blocked=(field:keyof InventoryItem)=>(labels[field]||[]).some(label=>unavailable.includes(label));return !Number.isFinite(av)||!Number.isFinite(bv)||blocked(a)||blocked(b)?[]:[[av,bv] as [number,number]]});if(pairs.length<3)return {value:null,n:pairs.length,reason:"Fewer than three paired observations"};const mx=pairs.reduce((s,p)=>s+p[0],0)/pairs.length,my=pairs.reduce((s,p)=>s+p[1],0)/pairs.length;let top=0,dx=0,dy=0;pairs.forEach(([x,y])=>{const xd=x-mx,yd=y-my;top+=xd*yd;dx+=xd*xd;dy+=yd*yd});if(dx===0||dy===0)return {value:null,n:pairs.length,reason:"Zero variance"};return {value:top/Math.sqrt(dx*dy),n:pairs.length,reason:null}}

const addMonth=(month:string,n:number)=>{const [y,m]=month.split("-").map(Number),d=new Date(Date.UTC(y,m-1+n,1));return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}`};
const monthDistance=(start:string,end:string)=>{const [sy,sm]=start.split("-").map(Number),[ey,em]=end.split("-").map(Number);return (ey-sy)*12+em-sm};
const monthRange=(start:string,end:string)=>Array.from({length:Math.max(0,monthDistance(start,end))+1},(_,i)=>addMonth(start,i));

function issueSeries(transactions:Transaction[],skus:Set<string>,coverage:TransactionCoverage={sourceComplete:false}):MonthlySeries{const map=new Map<string,number>();transactions.filter(t=>skus.has(t.sku)&&t.type.toLowerCase()==="issue"&&validMonth(t.date?.slice(0,7))&&Number.isFinite(Number(t.quantity))).forEach(t=>{const month=t.date.slice(0,7);map.set(month,(map.get(month)||0)+Math.abs(Number(t.quantity)))});const observed=[...map.keys()].sort(),start=validMonth(coverage.startMonth)?coverage.startMonth!:observed[0],end=validMonth(coverage.endMonth)?coverage.endMonth!:observed[observed.length-1];if(!start||!end||start>end)return {points:[],observedCount:0,missingMonths:[],zeroFilledMonths:[],sourceComplete:coverage.sourceComplete};const points=monthRange(start,end).map(month=>map.has(month)?{month,value:map.get(month)!,status:"observed" as const}:coverage.sourceComplete?{month,value:0,status:"zero" as const}:{month,value:null,status:"missing" as const});return {points,observedCount:points.filter(p=>p.value!=null).length,missingMonths:points.filter(p=>p.status==="missing").map(p=>p.month),zeroFilledMonths:points.filter(p=>p.status==="zero").map(p=>p.month),sourceComplete:coverage.sourceComplete}}

export function monthlyIssueSeries(transactions:Transaction[],sku:string,coverage:TransactionCoverage={sourceComplete:false}){return issueSeries(transactions,new Set([sku]),coverage)}
export function categoryIssueSeries(transactions:Transaction[],items:InventoryItem[],category:string,coverage:TransactionCoverage={sourceComplete:false}){return issueSeries(transactions,new Set(items.filter(i=>i.category===category).map(i=>i.sku)),coverage)}
function linearFit(pairs:{x:number;y:number}[]){const n=pairs.length,mx=pairs.reduce((s,p)=>s+p.x,0)/n,my=pairs.reduce((s,p)=>s+p.y,0)/n;let top=0,bottom=0;pairs.forEach(({x,y})=>{top+=(x-mx)*(y-my);bottom+=(x-mx)**2});const slope=bottom?top/bottom:0;return {predict:(x:number)=>Math.max(0,my-slope*mx+slope*x)}}

function forecastMonthlySeries(series:MonthlySeries,id:string,horizon:number):ForecastResult|null{const actuals=series.points.flatMap((p,index)=>p.value==null?[]:[{...p,index,value:p.value}]);if(actuals.length<6)return null;const testCount=Math.min(3,Math.max(2,Math.floor(actuals.length*.25))),trainCount=actuals.length-testCount;if(trainCount<4)return null;const train=actuals.slice(0,trainCount),test=actuals.slice(trainCount),fit=linearFit(train.map(p=>({x:p.index,y:p.value}))),testIndexes=new Set(test.map(p=>p.index));const points:ForecastPoint[]=series.points.map((p,index)=>({month:p.month,actual:p.value,predicted:fit.predict(index),phase:p.value==null?"gap":testIndexes.has(index)?"test":"train"}));const heldOut=points.filter(p=>p.phase==="test"&&p.actual!=null),mae=heldOut.reduce((s,p)=>s+Math.abs(p.actual!-p.predicted),0)/heldOut.length,rmse=Math.sqrt(heldOut.reduce((s,p)=>s+(p.actual!-p.predicted)**2,0)/heldOut.length),nonzero=heldOut.filter(p=>p.actual!>0),mape=nonzero.length?nonzero.reduce((s,p)=>s+Math.abs((p.actual!-p.predicted)/p.actual!),0)/nonzero.length:null,last=series.points[series.points.length-1].month;for(let h=1;h<=horizon;h++)points.push({month:addMonth(last,h),actual:null,predicted:fit.predict(series.points.length+h-1),phase:"future"});const limitations=["Illustrative numerical demand regression only; it does not estimate slow-moving, non-moving, or excess-inventory probabilities.","MAE and RMSE use only the final chronologically held-out observed months; no future month is shuffled into training.","No production forecast-accuracy claim is made."];if(series.missingMonths.length)limitations.push(`${series.missingMonths.length} calendar gap(s) remain missing because source completeness is not confirmed; they were not converted to zero.`);if(series.zeroFilledMonths.length)limitations.push(`${series.zeroFilledMonths.length} calendar gap(s) were set to zero because the source period was explicitly marked complete.`);return {sku:id,observations:actuals.length,trainCount,testCount,horizon,mae,rmse,mape,points,missingMonths:series.missingMonths,zeroFilledMonths:series.zeroFilledMonths,limitations}}

export function forecastSku(transactions:Transaction[],sku:string,horizon=3,coverage:TransactionCoverage={sourceComplete:false}){return forecastMonthlySeries(monthlyIssueSeries(transactions,sku,coverage),sku,horizon)}
export function forecastCategory(transactions:Transaction[],items:InventoryItem[],category:string,horizon=3,coverage:TransactionCoverage={sourceComplete:false}){return forecastMonthlySeries(categoryIssueSeries(transactions,items,category,coverage),category,horizon)}

export function formatCurrencyValue(value:number,sourceCurrency:string,displayCurrency=sourceCurrency,exchangeRate?:number|null){const canConvert=displayCurrency!==sourceCurrency&&Number.isFinite(exchangeRate)&&Number(exchangeRate)>0,currency=canConvert?displayCurrency:sourceCurrency,amount=canConvert?value*Number(exchangeRate):value;if(!Number.isFinite(amount))return {text:"Unavailable",currency,amount,converted:canConvert,reason:"Numeric value is unavailable."};let textValue:string;try{textValue=new Intl.NumberFormat("en-US",{style:"currency",currency,maximumFractionDigits:0}).format(amount)}catch{textValue=`${currency} ${Math.round(amount).toLocaleString("en-US")}`}return {text:textValue,currency,amount,converted:canConvert,reason:displayCurrency!==sourceCurrency&&!canConvert?"A positive exchange rate is required; source-currency values remain unchanged.":null}}
export const unique=(items:InventoryItem[],field:keyof InventoryItem)=>[...new Set(items.map(i=>String(i[field]||"Unavailable")))].sort();
