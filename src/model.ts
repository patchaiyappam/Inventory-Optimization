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
export type AppData = {meta:{datasetName:string;mode:string;sourceLabel:string;currency:string;generatedAt:string;limitations:string[]};assumptions:Record<string,number|string>;items:InventoryItem[];transactions:Transaction[];recoveryActions:RecoveryAction[];recoveryResults:unknown[]};

export type Filters = {search:string;category:string;supplier:string;warehouse:string;abc:string;fsn:string;xyz:string;aging:string;risk:string;valueMin:string;valueMax:string;costMin:string;costMax:string};
export const EMPTY_FILTERS:Filters={search:"",category:"All",supplier:"All",warehouse:"All",abc:"All",fsn:"All",xyz:"All",aging:"All",risk:"All",valueMin:"",valueMax:"",costMin:"",costMax:""};

export type ImportIssue={row:number;field:string;message:string;severity:"error"|"warning"};
export type ImportPreview={fileName:string;items:InventoryItem[];issues:ImportIssue[];rowCount:number;sheetName:string};
export type ForecastPoint={month:string;actual:number|null;predicted:number};
export type ForecastResult={sku:string;observations:number;trainCount:number;testCount:number;horizon:number;mae:number;rmse:number;mape:number|null;points:ForecastPoint[];limitations:string[]};

const HEADER_ALIASES:Record<string,string[]>={
  sku:["sku","materialcode","itemcode","productcode","stockcode"], name:["name","materialname","itemname","description","productname"],
  category:["category","materialcategory","productcategory"], supplier:["supplier","vendor","suppliername","vendorname"],
  warehouse:["warehouse","location","storagelocation","plant"], quantity:["quantity","qty","stockquantity","onhand","onhandquantity","currentstock"],
  unitCost:["unitcost","cost","price","unitprice","standardcost"], purchaseDate:["purchasedate","procurementdate"],
  lastReceiptDate:["lastreceiptdate","receiptdate","latestreceiptdate"], lastIssueDate:["lastissuedate","issuedate","latestissuedate"],
  lastMovementDate:["lastmovementdate","movementdate"], avgMonthlyConsumption:["avgmonthlyconsumption","averagemonthlyconsumption","monthlydemand","averageusage"],
  annualConsumptionValue:["annualconsumptionvalue","annualusagevalue"], leadTimeWeeks:["leadtimeweeks","leadtime","leadtimeweek"],
  moq:["moq","minimumorderquantity"], openPurchaseOrders:["openpurchaseorders","openpo","openpoqty"],
  confirmedRequirementDate:["confirmedrequirementdate","requirementdate","needbydate"], abc:["abc","abcclass"], fsn:["fsn","fsnclass"],
  xyz:["xyz","xyzclass"], ageMonths:["agemonths","inventoryagemonths","age"], cv:["cv","coefficientofvariation"]
};

const key=(value:unknown)=>String(value??"").trim().toLowerCase().replace(/[^a-z0-9]/g,"");
const num=(value:unknown)=>{const cleaned=String(value??"").replace(/[$,%\s,]/g,"");if(!cleaned)return null;const n=Number(cleaned);return Number.isFinite(n)?n:null};
const iso=(value:unknown)=>{if(value==null||value==="")return null;if(typeof value==="number"){const d=XLSX.SSF.parse_date_code(value);return d?new Date(Date.UTC(d.y,d.m-1,d.d)).toISOString():null}const d=new Date(String(value));return Number.isNaN(d.getTime())?null:d.toISOString()};
const monthsBetween=(date:string|null,now:Date)=>date?Math.max(0,(now.getTime()-new Date(date).getTime())/(30.4375*86400000)):0;
const aging=(m:number)=>m<3?"Fresh":m<6?"Normal":m<12?"Aging":m<24?"Old":"Dead";
const pick=(row:Record<string,unknown>,field:string)=>{const aliases=HEADER_ALIASES[field]||[];const hit=Object.keys(row).find(k=>aliases.includes(key(k)));return hit?row[hit]:undefined};
const text=(value:unknown)=>String(value??"").trim();
const classifyRisk=(age:number,fsn:string,turnover:number)=>{const reasons:string[]=[];if(age>=24)reasons.push("Inventory age is at least 24 months");else if(age>=12)reasons.push("Inventory age is at least 12 months");if(fsn==="N")reasons.push("Classified as non-moving");if(turnover>0&&turnover<1)reasons.push("Turnover is below 1.0x");const risk=reasons.length>=3||age>=24&&fsn==="N"?"Critical":reasons.length>=2||age>=24?"High":reasons.length?"Medium":"Low";return {risk,reasons:reasons.length?reasons:["No configured review trigger"]}};

export async function parseInventoryFile(file:File,carryingRate:number,recoveryRate:number):Promise<ImportPreview>{
  if(file.size>15*1024*1024)throw new Error("File exceeds the 15 MB local import limit.");
  const book=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
  const sheetName=book.SheetNames[0];
  if(!sheetName)throw new Error("No worksheet was found.");
  const rows=XLSX.utils.sheet_to_json<Record<string,unknown>>(book.Sheets[sheetName],{defval:"",raw:true});
  if(!rows.length)throw new Error("The first worksheet has no data rows.");
  if(rows.length>10000)throw new Error("Import is limited to 10,000 rows for this browser demonstration.");
  const issues:ImportIssue[]=[];const seen=new Set<string>();const items:InventoryItem[]=[];const now=new Date();
  rows.forEach((row,index)=>{
    const rowNo=index+2,sku=text(pick(row,"sku")),quantity=num(pick(row,"quantity")),unitCost=num(pick(row,"unitCost"));
    if(!sku)issues.push({row:rowNo,field:"SKU",message:"SKU is required.",severity:"error"});
    if(sku&&seen.has(sku.toLowerCase()))issues.push({row:rowNo,field:"SKU",message:"Duplicate SKU.",severity:"error"});
    if(quantity==null||quantity<0)issues.push({row:rowNo,field:"Quantity",message:"Quantity must be a non-negative number.",severity:"error"});
    if(unitCost==null||unitCost<0)issues.push({row:rowNo,field:"Unit cost",message:"Unit cost must be a non-negative number.",severity:"error"});
    if(issues.some(i=>i.row===rowNo&&i.severity==="error"))return;
    seen.add(sku.toLowerCase());
    const unavailable:string[]=[];const missingText=(field:string,label:string)=>{const v=text(pick(row,field));if(!v)unavailable.push(label);return v||"Unavailable"};
    const purchaseDate=iso(pick(row,"purchaseDate")),lastReceiptDate=iso(pick(row,"lastReceiptDate")),lastIssueDate=iso(pick(row,"lastIssueDate")),lastMovementDate=iso(pick(row,"lastMovementDate"));
    if(!purchaseDate&&!lastReceiptDate)unavailable.push("receipt/purchase date");
    const ageValue=num(pick(row,"ageMonths"))??monthsBetween(lastReceiptDate||purchaseDate,now);if(!pick(row,"ageMonths")&&!purchaseDate&&!lastReceiptDate)unavailable.push("inventory age");
    const monthly=num(pick(row,"avgMonthlyConsumption"))??0;if(pick(row,"avgMonthlyConsumption")===""||pick(row,"avgMonthlyConsumption")==null)unavailable.push("average monthly consumption");
    const lead=num(pick(row,"leadTimeWeeks"))??0;if(pick(row,"leadTimeWeeks")===""||pick(row,"leadTimeWeeks")==null)unavailable.push("lead time");
    const requirement=iso(pick(row,"confirmedRequirementDate"));if(!requirement)unavailable.push("confirmed requirement date");
    const value=quantity!*unitCost!,annualConsumption=num(pick(row,"annualConsumptionValue"))??monthly*12*unitCost!;
    const turnover=value>0?annualConsumption/value:0;const fsn=text(pick(row,"fsn")).toUpperCase()||"Unavailable";if(fsn==="Unavailable")unavailable.push("FSN class");
    const risk=classifyRisk(ageValue,fsn,turnover);const targetQty=monthly>0?monthly*((lead+2)/4.33):quantity!;const excessQty=monthly>0?Math.max(0,quantity!-targetQty):0;
    const trapped=excessQty*unitCost!;const item:InventoryItem={
      sku,name:missingText("name","material name"),category:missingText("category","category"),supplier:missingText("supplier","supplier"),supplierType:"Unavailable",warehouse:missingText("warehouse","warehouse"),
      purchaseDate:purchaseDate||"",lastReceiptDate,lastIssueDate,lastMovementDate,quantity:quantity!,unitCost:unitCost!,inventoryValue:value,avgMonthlyConsumption:monthly,
      annualConsumptionValue:annualConsumption,leadTimeWeeks:lead,moq:num(pick(row,"moq"))??0,openPurchaseOrders:num(pick(row,"openPurchaseOrders")),confirmedRequirementDate:requirement,
      abc:text(pick(row,"abc")).toUpperCase()||"Unavailable",fsn,xyz:text(pick(row,"xyz")).toUpperCase()||"Unavailable",ageMonths:ageValue,agingBucket:aging(ageValue),cv:num(pick(row,"cv"))??0,
      carryingRate,annualCarryingCost:value*carryingRate,cumulativeCarryingCost:ageValue>0?value*carryingRate*ageValue/12:0,inventoryDays:monthly>0?quantity!/monthly*30.4375:0,
      turnoverRatio:turnover,capitalBlocked:trapped,moqExcessQuantity:excessQty,moqExcessValue:trapped,risk:risk.risk,riskReasons:risk.reasons,
      recommendedAction:risk.risk==="Low"?"Monitor":"Review demand, redeployment and supplier-return options",recoverableValueEstimate:trapped*recoveryRate,recoveryRateAssumption:recoveryRate,recoveryStatus:"Not reviewed",unavailableFields:unavailable
    };
    ["abc","xyz"].forEach(field=>{if((item as unknown as Record<string,string>)[field]==="Unavailable")unavailable.push(`${field.toUpperCase()} class`)});
    if(unavailable.length)issues.push({row:rowNo,field:"Optional fields",message:`Unavailable: ${[...new Set(unavailable)].join(", ")}.`,severity:"warning"});
    items.push(item);
  });
  return {fileName:file.name,items,issues,rowCount:rows.length,sheetName};
}

export function applyFilters(items:InventoryItem[],f:Filters){
  const minV=num(f.valueMin),maxV=num(f.valueMax),minC=num(f.costMin),maxC=num(f.costMax),query=f.search.trim().toLowerCase();
  return items.filter(i=>(!query||`${i.sku} ${i.name} ${i.supplier} ${i.category} ${i.warehouse}`.toLowerCase().includes(query))
    &&(f.category==="All"||i.category===f.category)&&(f.supplier==="All"||i.supplier===f.supplier)&&(f.warehouse==="All"||i.warehouse===f.warehouse)
    &&(f.abc==="All"||i.abc===f.abc)&&(f.fsn==="All"||i.fsn===f.fsn)&&(f.xyz==="All"||i.xyz===f.xyz)&&(f.aging==="All"||i.agingBucket===f.aging)&&(f.risk==="All"||i.risk===f.risk)
    &&(minV==null||i.inventoryValue>=minV)&&(maxV==null||i.inventoryValue<=maxV)&&(minC==null||i.annualCarryingCost>=minC)&&(maxC==null||i.annualCarryingCost<=maxC));
}

export function pairedPearson(items:InventoryItem[],a:keyof InventoryItem,b:keyof InventoryItem){
  const missingLabels:Partial<Record<keyof InventoryItem,string[]>>={
    ageMonths:["inventory age","receipt/purchase date"],avgMonthlyConsumption:["average monthly consumption"],
    leadTimeWeeks:["lead time"],abc:["abc class"],fsn:["fsn class"],xyz:["xyz class"]
  };
  const pairs=items.flatMap(i=>{
    const av=Number(i[a]),bv=Number(i[b]),missing=(i.unavailableFields||[]).map(x=>x.toLowerCase());
    const unavailable=(field:keyof InventoryItem)=>(missingLabels[field]||[]).some(label=>missing.includes(label));
    if(!Number.isFinite(av)||!Number.isFinite(bv)||unavailable(a)||unavailable(b))return[];
    return[[av,bv] as [number,number]];
  });
  if(pairs.length<3)return {value:null,n:pairs.length,reason:"Fewer than three paired observations"};
  const mx=pairs.reduce((s,p)=>s+p[0],0)/pairs.length,my=pairs.reduce((s,p)=>s+p[1],0)/pairs.length;let top=0,dx=0,dy=0;
  pairs.forEach(([x,y])=>{const xd=x-mx,yd=y-my;top+=xd*yd;dx+=xd*xd;dy+=yd*yd});
  if(dx===0||dy===0)return {value:null,n:pairs.length,reason:"Zero variance"};return {value:top/Math.sqrt(dx*dy),n:pairs.length,reason:null};
}

export function monthlyIssueSeries(transactions:Transaction[],sku:string){
  const map=new Map<string,number>();transactions.filter(t=>t.sku===sku&&t.type.toLowerCase()==="issue"&&t.date&&Number.isFinite(Number(t.quantity))).forEach(t=>{const month=t.date.slice(0,7);map.set(month,(map.get(month)||0)+Math.abs(Number(t.quantity)))});
  return [...map].sort((a,b)=>a[0].localeCompare(b[0])).map(([month,value])=>({month,value}));
}

function linearFit(values:number[]){const n=values.length,mx=(n-1)/2,my=values.reduce((s,v)=>s+v,0)/n;let top=0,bottom=0;values.forEach((v,i)=>{top+=(i-mx)*(v-my);bottom+=(i-mx)**2});const slope=bottom?top/bottom:0;return {slope,intercept:my-slope*mx,predict:(x:number)=>Math.max(0,my-slope*mx+slope*x)};}
const addMonth=(month:string,n:number)=>{const [y,m]=month.split("-").map(Number);const d=new Date(Date.UTC(y,m-1+n,1));return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}`};

export function forecastSku(transactions:Transaction[],sku:string,horizon=3):ForecastResult|null{
  const series=monthlyIssueSeries(transactions,sku);if(series.length<6)return null;const testCount=Math.min(3,Math.max(2,Math.floor(series.length*.25))),trainCount=series.length-testCount;if(trainCount<4)return null;
  const fit=linearFit(series.slice(0,trainCount).map(x=>x.value));const predictions:ForecastPoint[]=series.map((x,i)=>({month:x.month,actual:x.value,predicted:fit.predict(i)}));const test=predictions.slice(trainCount);
  const mae=test.reduce((s,p)=>s+Math.abs(p.actual!-p.predicted),0)/test.length;const rmse=Math.sqrt(test.reduce((s,p)=>s+(p.actual!-p.predicted)**2,0)/test.length);const nonzero=test.filter(p=>p.actual!>0);const mape=nonzero.length?nonzero.reduce((s,p)=>s+Math.abs((p.actual!-p.predicted)/p.actual!),0)/nonzero.length:null;
  const last=series[series.length-1].month;for(let h=1;h<=horizon;h++)predictions.push({month:addMonth(last,h),actual:null,predicted:fit.predict(series.length+h-1)});
  return {sku,observations:series.length,trainCount,testCount,horizon,mae,rmse,mape,points:predictions,limitations:["Illustrative linear time-trend only; seasonality and causal drivers are not modeled.","Validation uses the final dated observations and never shuffles future months into training.","Predictions are advisory and must not trigger purchasing or disposal decisions."]};
}

export function projectedCarryingCost(item:InventoryItem,analysisDate:string,rate:number){
  if(!item.confirmedRequirementDate||!item.leadTimeWeeks)return null;const requirement=new Date(item.confirmedRequirementDate),analysis=new Date(analysisDate);if(Number.isNaN(requirement.getTime())||Number.isNaN(analysis.getTime()))return null;
  const holdingDays=Math.floor((requirement.getTime()-analysis.getTime())/86400000-item.leadTimeWeeks*7);if(holdingDays<=0)return {holdingDays:0,cost:0};return {holdingDays,cost:item.inventoryValue*rate*holdingDays/365};
}

export const unique=(items:InventoryItem[],field:keyof InventoryItem)=>[...new Set(items.map(i=>String(i[field]||"Unavailable")))].sort();

