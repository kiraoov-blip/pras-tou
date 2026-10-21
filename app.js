console.info("PRAS - TOU model v0.6.0 loaded (목표 효과 역산)");

const META_URL="./data/tou_data.xlsx";
const LAND_URL="./data/usage_land.xlsx";
const JEJU_URL="./data/usage_jeju.xlsx";
const PERIODS=["경부하","중간부하","최대부하"];
const PERIOD_CLASS={"경부하":"off","중간부하":"mid","최대부하":"peak"};
const SEASONS=["하계","춘추계","동계"];
const DAY_TYPES=["평일","토요일","일·공휴일"];
const FULL_YEARS=[2022,2023,2024,2025];
const CATEGORIES=[
  {id:"ALL_TOU",name:"전체종별",sheet:null,group:null},
  {id:"IND_GAP2",name:"산업용(갑)Ⅱ",sheet:"산업용(갑)II",group:"IND_GAP2"},
  {id:"IND_EUL",name:"산업용(을)",sheet:"산업용(을)",group:"IND_EUL"},
  {id:"GEN_GAP2",name:"일반용(갑)Ⅱ",sheet:"일반용(갑)II",group:"GEN_GAP2"},
  {id:"GEN_EUL",name:"일반용(을)",sheet:"일반용(을)",group:"GEN_EUL"},
  {id:"EDU_EUL",name:"교육용(을)",sheet:"교육용(을)",group:"EDU_EUL"},
  {id:"EV",name:"EV충전전력",sheet:"EV충전전력",group:"EV"},
];
const SINGLE_CATEGORY_IDS=CATEGORIES.filter(x=>x.id!=="ALL_TOU").map(x=>x.id);

const state={
  metaWorkbook:null,
  weights:[],metaTariffs:[],
  usage:{LAND:{},JEJU:{}},
  selectedCategory:"ALL_TOU",
  activeSeason:"하계",graphDayType:"전체",
  scenarioSchedule:null,scenarioRates:null,
  // 전체종별은 종별마다 요금표가 달라 단가표를 직접 고칠 수 없으므로, 종별
  // 요금표(allRateVersion)에 계절·시간대별 배율(rateMult)을 곱해 단가를 조정한다.
  allRateVersion:"POST",rateMult:null,
  // 시나리오 주말 할인율(토·일·공휴일 11~14시, 산업용(을)·EV). 0.5 = 50% 할인.
  discountRates:null,discountForce:null,
  lastResult:null,solverProposal:null,
};

const $=id=>document.getElementById(id);
const text=v=>String(v??"").trim();
const number=v=>{const n=Number(String(v??"").replace(/,/g,""));return Number.isFinite(n)?n:0};
const clone=o=>JSON.parse(JSON.stringify(o));
const catMeta=id=>CATEGORIES.find(x=>x.id===id);
const selectedRegion=()=>$("region").value;
const selectedCategoryId=()=>$("category").value;
const ymd=d=>d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`:"-";
const dateKey=d=>`${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}`;
const fmtWon=n=>{const a=Math.abs(n),s=n<0?"-":"";if(a>=1e12)return s+(a/1e12).toLocaleString("ko-KR",{maximumFractionDigits:3})+"조원";if(a>=1e8)return s+(a/1e8).toLocaleString("ko-KR",{maximumFractionDigits:1})+"억원";if(a>=1e4)return s+(a/1e4).toLocaleString("ko-KR",{maximumFractionDigits:1})+"만원";return s+Math.round(a).toLocaleString("ko-KR")+"원"};
const fmtEnergy=n=>{const a=Math.abs(n),s=n<0?"-":"";if(a>=1e9)return s+(a/1e9).toLocaleString("ko-KR",{maximumFractionDigits:2})+"TWh";if(a>=1e6)return s+(a/1e6).toLocaleString("ko-KR",{maximumFractionDigits:2})+"GWh";if(a>=1e3)return s+(a/1e3).toLocaleString("ko-KR",{maximumFractionDigits:2})+"MWh";return s+a.toLocaleString("ko-KR",{maximumFractionDigits:1})+"kWh"};
const signedWon=n=>(n>=0?"+":"")+fmtWon(n);
const cssSign=n=>n>=0?"pos":"neg";

// ── KPI 카드 전용 표시 — PRAS-DER·CARE-Jeju의 metric-card와 동일하게 "숫자는
// 소수 첫째자리까지 + 단위는 숫자보다 작은 글씨"로 통일한다. 위 fmtWon/fmtEnergy는
// 단위별로 소수자리가 제각각(조원 3자리, 억원/만원 1자리 등)이고 단위가 숫자
// 문자열 안에 그대로 붙어 있어(예: "3.63GWh") 그대로는 이 카드 스타일에 맞지
// 않는다 — 값과 단위를 분리해 반환하는 전용 버전을 따로 둔다(표·다른 곳의
// 기존 fmtWon/fmtEnergy 호출은 그대로 둔다).
const splitWon=n=>{const a=Math.abs(n),s=n<0?"-":"";if(a>=1e12)return{value:s+(a/1e12).toLocaleString("ko-KR",{minimumFractionDigits:1,maximumFractionDigits:1}),unit:"조원"};if(a>=1e8)return{value:s+(a/1e8).toLocaleString("ko-KR",{minimumFractionDigits:1,maximumFractionDigits:1}),unit:"억원"};if(a>=1e4)return{value:s+(a/1e4).toLocaleString("ko-KR",{minimumFractionDigits:1,maximumFractionDigits:1}),unit:"만원"};return{value:s+a.toLocaleString("ko-KR",{minimumFractionDigits:1,maximumFractionDigits:1}),unit:"원"}};
const signedSplitWon=n=>{const r=splitWon(n);return{value:(n>=0?"+":"")+r.value,unit:r.unit}};
// 분석 판매량 KPI는 항상 TWh 단위로 표시한다(요청). 전국·전체종별은 수백 TWh,
// 특정 요금제만 고르면 1TWh 미만도 나오므로 값 크기에 따라 소수자리만 늘린다.
const splitEnergyTWh=n=>{const a=Math.abs(n),s=n<0?"-":"",t=a/1e9,d=t>=100?1:t>=1?2:3;return{value:s+t.toLocaleString("ko-KR",{minimumFractionDigits:d,maximumFractionDigits:d}),unit:"TWh"}};
const eok=n=>(n/1e8).toLocaleString("ko-KR",{minimumFractionDigits:1,maximumFractionDigits:1});
const signedEok=n=>(n>=0?"+":"")+eok(n);
const seasonLabel=s=>s==="춘추계"?"봄·가을철":s==="동계"?"겨울철":s;
const pct=(v,d=1)=>(v*100).toLocaleString("ko-KR",{minimumFractionDigits:d,maximumFractionDigits:d})+"%";
/** strong 엘리먼트 안을 "값 + 더 작은 글씨의 단위(<small>)"로 채운다(PRAS-DER
 * .metric-card strong small 패턴과 동일한 마크업 구조). */
function setKpiValue(el,split){el.textContent="";el.append(document.createTextNode(split.value),Object.assign(document.createElement("small"),{textContent:split.unit}))}

function sheetRows(wb,name){const ws=wb.Sheets[name];return ws?XLSX.utils.sheet_to_json(ws,{defval:"",raw:true}):[]}
function normalizeSeason(v){const s=text(v).replace(/\s/g,"");if(["하계","여름","여름철"].includes(s))return"하계";if(["춘추계","봄가을","봄·가을","봄가을철","봄·가을철"].includes(s))return"춘추계";if(["동계","겨울","겨울철"].includes(s))return"동계";return s}
function normalizeDayType(v){const s=text(v).replace(/\s/g,"");if(["평일","주중"].includes(s))return"평일";if(["토요일","토"].includes(s))return"토요일";if(["일·공휴일","일/공휴일","일공휴일","일요일·공휴일","일요일/공휴일","일요일","공휴일"].includes(s))return"일·공휴일";return s}
function seasonFromMonth(m){return [11,12,1,2].includes(m)?"동계":[6,7,8].includes(m)?"하계":"춘추계"}
function dateFromRaw(v){
  if(v instanceof Date&&!Number.isNaN(v.getTime()))return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v==="number"){
    const digits=String(Math.trunc(v));
    if(/^\d{8}$/.test(digits))return new Date(+digits.slice(0,4),+digits.slice(4,6)-1,+digits.slice(6,8));
    if(v>20000&&v<80000&&typeof XLSX!=="undefined"&&XLSX.SSF){const p=XLSX.SSF.parse_date_code(v);if(p)return new Date(p.y,p.m-1,p.d)}
  }
  const digits=text(v).replace(/[^0-9]/g,"");
  if(digits.length>=8)return new Date(+digits.slice(0,4),+digits.slice(4,6)-1,+digits.slice(6,8));
  return null;
}

// 요금 적용용 휴일구분. 전기요금표의 토요일·공휴일 계산기준은 임시공휴일을 제외하므로
// 2025-01-27, 2025-06-03은 달력상 임시공휴일이지만 요금요일구분은 평일로 둔다.
const TARIFF_HOLIDAYS={
  "20250101":"일·공휴일","20250127":"평일","20250128":"일·공휴일","20250129":"일·공휴일","20250130":"일·공휴일",
  "20250301":"일·공휴일","20250303":"일·공휴일","20250505":"일·공휴일","20250506":"일·공휴일","20250603":"평일","20250606":"일·공휴일",
  "20250815":"일·공휴일","20251003":"일·공휴일","20251005":"일·공휴일","20251006":"일·공휴일","20251007":"일·공휴일","20251008":"일·공휴일","20251009":"일·공휴일","20251225":"일·공휴일",
  "20260101":"일·공휴일","20260216":"일·공휴일","20260217":"일·공휴일","20260218":"일·공휴일","20260301":"일·공휴일","20260302":"일·공휴일",
  "20260505":"일·공휴일","20260524":"일·공휴일","20260525":"일·공휴일","20260603":"일·공휴일","20260606":"일·공휴일"
};
function generatedDayType(d){const key=dateKey(d);if(TARIFF_HOLIDAYS[key])return TARIFF_HOLIDAYS[key];const day=d.getDay();return day===0?"일·공휴일":day===6?"토요일":"평일"}

function parseUsageSheet(ws,sheetName){
  if(!ws)throw new Error(`사용량 시트 '${sheetName}' 없음`);
  const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
  if(matrix.length<2)return[];
  const headers=matrix[0].map(v=>text(v).replace(/\s/g,""));
  const findIndex=aliases=>headers.findIndex(h=>aliases.includes(h));
  const dateIdx=findIndex(["날짜","일자","일시"]),yearIdx=findIndex(["연도","연"]),monthIdx=findIndex(["월"]),seasonIdx=findIndex(["계절"]),dayIdx=findIndex(["요일","요일구분"]);
  const hourIdx=Array.from({length:24},(_,i)=>findIndex([`H${String(i+1).padStart(2,"0")}`]));
  if(dateIdx<0||hourIdx.some(x=>x<0))throw new Error(`'${sheetName}' 시트에 날짜 또는 H01~H24 필수열이 없음`);
  const out=[];
  for(let r=1;r<matrix.length;r++){
    const row=matrix[r],date=dateFromRaw(row[dateIdx]);
    if(!date||Number.isNaN(date.getTime()))continue;
    const year=yearIdx>=0&&number(row[yearIdx])?number(row[yearIdx]):date.getFullYear();
    const month=monthIdx>=0&&number(row[monthIdx])?number(row[monthIdx]):date.getMonth()+1;
    if(year<2022||year>2026)continue;
    if(year===2026&&(month>6||date>new Date(2026,5,30)))continue;
    const hasHourly=hourIdx.some(idx=>row[idx]!==""&&row[idx]!==null&&row[idx]!==undefined);
    if(!hasHourly)continue;
    const hours=hourIdx.map(idx=>number(row[idx]));
    const suppliedSeason=seasonIdx>=0?normalizeSeason(row[seasonIdx]):"";
    const suppliedDay=dayIdx>=0?normalizeDayType(row[dayIdx]):"";
    const season=(year>=2025||!SEASONS.includes(suppliedSeason))?seasonFromMonth(month):suppliedSeason;
    const dayType=(year>=2025||!DAY_TYPES.includes(suppliedDay))?generatedDayType(date):suppliedDay;
    out.push({date,year,month,season,dayType,hours});
  }
  return out;
}

// 신규 종별 요금표. 산업용(갑)Ⅱ·산업용(을)은 기존 tou_data.xlsx의 검증된 요금표를 사용한다.
const STATIC_TARIFF_SPECS={
  GEN_GAP2:{contract:"300kW 미만",items:[
    ["고압A","선택Ⅰ",7170,{"경부하":[89.4,89.4,98.1],"중간부하":[140.6,96.8,128.5],"최대부하":[163.1,108.1,143.3]}],
    ["고압A","선택Ⅱ",8230,{"경부하":[84.1,84.1,92.8],"중간부하":[135.3,91.5,123.2],"최대부하":[157.8,102.8,138.0]}],
    ["고압B","선택Ⅰ",7170,{"경부하":[88.8,88.8,97.8],"중간부하":[137.4,94.7,125.1],"최대부하":[153.8,100.1,139.3]}],
    ["고압B","선택Ⅱ",8230,{"경부하":[83.5,83.5,92.5],"중간부하":[132.1,89.4,119.8],"최대부하":[148.5,94.8,134.0]}],
  ]},
  GEN_EUL:{contract:"300kW 이상",items:[
    ["고압A","선택Ⅰ",7220,{"경부하":[92.8,92.8,99.8],"중간부하":[145.7,115.3,145.9],"최대부하":[227.8,146.0,203.4]}],
    ["고압A","선택Ⅱ",8320,{"경부하":[87.3,87.3,94.3],"중간부하":[140.2,109.8,140.4],"최대부하":[222.3,140.5,197.9]}],
    ["고압A","선택Ⅲ",9810,{"경부하":[86.4,86.4,93.7],"중간부하":[139.6,108.5,139.8],"최대부하":[209.9,132.2,186.7]}],
    ["고압B","선택Ⅰ",6630,{"경부하":[95.9,95.9,102.9],"중간부하":[148.2,118.2,148.2],"최대부하":[229.4,148.5,204.4]}],
    ["고압B","선택Ⅱ",7380,{"경부하":[92.1,92.1,99.1],"중간부하":[144.4,114.4,144.4],"최대부하":[225.6,144.7,200.6]}],
    ["고압B","선택Ⅲ",8190,{"경부하":[90.4,90.4,97.5],"중간부하":[142.7,112.8,142.7],"최대부하":[224.0,143.1,198.9]}],
  ]},
  EDU_EUL:{contract:"1,000kW 이상",items:[
    ["고압A","선택Ⅰ",6090,{"경부하":[76.5,76.5,80.5],"중간부하":[121.2,90.9,119.7],"최대부하":[187.1,111.4,158.4]}],
    ["고압A","선택Ⅱ",6980,{"경부하":[72.0,72.0,76.0],"중간부하":[116.7,86.4,115.2],"최대부하":[182.6,106.9,153.9]}],
    ["고압B","선택Ⅰ",6090,{"경부하":[75.0,75.0,78.8],"중간부하":[118.5,89.2,116.8],"최대부하":[181.4,109.0,154.1]}],
    ["고압B","선택Ⅱ",6980,{"경부하":[70.5,70.5,74.3],"중간부하":[114.0,84.7,112.3],"최대부하":[176.9,104.5,149.6]}],
  ]},
  EV:{items:[
    ["자가소비용","저압","기본",2390,{"경부하":[84.3,85.4,107.4],"중간부하":[172.0,97.2,154.9],"최대부하":[259.2,102.1,217.5]}],
    ["자가소비용","고압","기본",2580,{"경부하":[79.2,80.2,96.6],"중간부하":[137.4,91.0,127.7],"최대부하":[190.4,94.9,165.5]}],
    ["충전서비스 제공사업자용","저압","선택Ⅰ",2390,{"경부하":[95.9,85.4,110.6],"중간부하":[162.2,97.2,143.1],"최대부하":[203.5,102.1,172.0]}],
    ["충전서비스 제공사업자용","저압","선택Ⅱ",2390,{"경부하":[83.1,85.4,105.8],"중간부하":[140.0,97.2,126.7],"최대부하":[270.8,102.1,227.0]}],
    ["충전서비스 제공사업자용","저압","선택Ⅲ",2390,{"경부하":[90.1,85.4,115.5],"중간부하":[138.6,97.2,125.4],"최대부하":[236.0,102.1,198.4]}],
    ["충전서비스 제공사업자용","고압","선택Ⅰ",2580,{"경부하":[89.8,80.2,99.4],"중간부하":[129.9,91.0,118.4],"최대부하":[151.2,94.9,132.4]}],
    ["충전서비스 제공사업자용","고압","선택Ⅱ",2580,{"경부하":[78.2,80.2,95.2],"중간부하":[113.0,91.0,105.5],"최대부하":[198.6,94.9,172.4]}],
    ["충전서비스 제공사업자용","고압","선택Ⅲ",2580,{"경부하":[84.5,80.2,103.6],"중간부하":[111.9,91.0,104.5],"최대부하":[174.0,94.9,151.6]}],
  ]}
};
function generateStaticTariffs(){
  const rows=[];
  for(const [cat,spec] of Object.entries(STATIC_TARIFF_SPECS)){
    const postStart=(cat==="EV")?20260416:20260601;
    for(const item of spec.items){
      let contract=spec.contract,voltage,choice,base,rates;
      if(cat==="EV"){[contract,voltage,choice,base,rates]=item}else{[voltage,choice,base,rates]=item}
      for(const ver of ["PRE","POST"]){
        for(let si=0;si<3;si++)for(const p of PERIODS)rows.push({
          "요금그룹ID":cat,"버전ID":ver,"버전명":ver==="PRE"?"개편 전":"개편 후",
          "적용시작일":ver==="PRE"?20250401:postStart,"적용종료일":ver==="PRE"?postStart-1:99991231,
          "계약전력구간":contract,"전압구분":voltage,"선택요금":choice,"계절":SEASONS[si],"부하시간대":p,
          "전력량요금":rates[p][si],"기본요금":base,"비고":""
        });
      }
    }
  }
  return rows;
}
const STATIC_TARIFFS=generateStaticTariffs();

async function fetchBuffer(url){const r=await fetch(`${url}?t=${Date.now()}`,{cache:"no-store"});if(!r.ok)throw new Error(`${url}: HTTP ${r.status}`);return r.arrayBuffer()}
async function loadGithubData(){
  try{
    if(typeof XLSX==="undefined")throw new Error("SheetJS 라이브러리를 불러오지 못함");
    const [metaBuf,landBuf,jejuBuf]=await Promise.all([fetchBuffer(META_URL),fetchBuffer(LAND_URL),fetchBuffer(JEJU_URL)]);
    state.metaWorkbook=XLSX.read(metaBuf,{type:"array",cellDates:true});
    const land=XLSX.read(landBuf,{type:"array",cellDates:true}),jeju=XLSX.read(jejuBuf,{type:"array",cellDates:true});
    state.weights=sheetRows(state.metaWorkbook,"요금가중치");
    state.metaTariffs=sheetRows(state.metaWorkbook,"요금표").filter(r=>["IND_GAP2","IND_EUL"].includes(text(r["요금그룹ID"])));
    for(const c of CATEGORIES.filter(x=>x.sheet)){
      state.usage.LAND[c.id]=parseUsageSheet(land.Sheets[c.sheet],`육지/${c.sheet}`);
      state.usage.JEJU[c.id]=parseUsageSheet(jeju.Sheets[c.sheet],`제주/${c.sheet}`);
      if(!state.usage.LAND[c.id].length||!state.usage.JEJU[c.id].length)throw new Error(`${c.name} 사용량자료가 비어 있음`);
    }
    populateCategories();populateYears();populateVersions();selectCategory();
    console.info("GitHub 데이터 연결 완료",{weights:state.weights.length});
  }catch(e){console.error(e);$("warning").innerHTML=`<b>데이터 오류:</b> ${e.message}`}
}

function populateCategories(){
  $("category").innerHTML=CATEGORIES.map(c=>`<option value="${c.id}">${c.name}</option>`).join("");
  $("category").value=state.selectedCategory;
}
function allUsageYears(){const rows=state.usage.LAND.IND_EUL||[];return[...new Set(rows.map(r=>r.year))].sort((a,b)=>b-a)}
function populateYears(){
  const ys=allUsageYears();
  $("year").innerHTML=ys.map(y=>`<option value="${y}">${y===2026?"2026년(1~6월)":`${y}년`}</option>`).join("")+`<option value="AVG">2022~2025 연평균</option>`;
  $("year").value=ys.includes(2025)?"2025":String(ys[0]||"");
}
function populateVersions(){
  const opts='<option value="PRE">개편 전</option><option value="POST">개편 후</option>';
  $("baseVersion").innerHTML=opts;$("scenarioVersion").innerHTML=opts;
  $("baseVersion").value="PRE";$("scenarioVersion").value="POST";
}
function selectedCats(){return selectedCategoryId()==="ALL_TOU"?SINGLE_CATEGORY_IDS:[selectedCategoryId()]}
function selectedRegions(){return selectedRegion()==="NATIONAL"?["LAND","JEJU"]:[selectedRegion()]}
function editingRegion(){return selectedRegion()==="JEJU"?"JEJU":"LAND"}
function categoryTariffs(cat){return cat==="IND_GAP2"||cat==="IND_EUL"?state.metaTariffs.filter(r=>text(r["요금그룹ID"])===cat):STATIC_TARIFFS.filter(r=>text(r["요금그룹ID"])===cat)}
function unique(rows,key){return[...new Set(rows.map(r=>text(r[key])).filter(Boolean))]}

function selectCategory(){
  state.selectedCategory=selectedCategoryId();
  updateCustomerScopeControls();
  updateWeekendDiscountControl();
  loadScenarioPreset();
  updateWarning();
}
function updateCustomerScopeControls(){
  const allCat=selectedCategoryId()==="ALL_TOU";
  if(allCat){$("customerScope").value="all";$("customerScope").disabled=true}else $("customerScope").disabled=false;
  const all=$("customerScope").value==="all";
  ["contract","voltage","choice"].forEach(id=>$(id).disabled=all||allCat);
  if(all||allCat){
    $("contract").innerHTML='<option value="ALL">전체</option>';$("voltage").innerHTML='<option value="ALL">전체</option>';$("choice").innerHTML='<option value="ALL">전체</option>';
  }else{
    const rows=categoryTariffs(selectedCategoryId());
    $("contract").innerHTML=unique(rows,"계약전력구간").map(v=>`<option>${v}</option>`).join("");
    populateVoltage();
  }
  document.querySelectorAll(".rate-adjust").forEach(b=>b.disabled=allCat);
}
function populateVoltage(){
  if($("customerScope").value==="all"||selectedCategoryId()==="ALL_TOU")return;
  const rows=categoryTariffs(selectedCategoryId()).filter(r=>text(r["계약전력구간"])===$("contract").value);
  $("voltage").innerHTML=unique(rows,"전압구분").map(v=>`<option>${v}</option>`).join("");populateChoice();
}
function populateChoice(){
  if($("customerScope").value==="all"||selectedCategoryId()==="ALL_TOU")return;
  const rows=categoryTariffs(selectedCategoryId()).filter(r=>text(r["계약전력구간"])===$("contract").value&&text(r["전압구분"])===$("voltage").value);
  $("choice").innerHTML=unique(rows,"선택요금").map(v=>`<option>${v}</option>`).join("");
}
function updateWeekendDiscountControl(){
  const cat=selectedCategoryId(),eligible=["IND_EUL","EV","ALL_TOU"].includes(cat),box=$("useScenarioDiscount"),label=$("weekendDiscountLabel");
  box.disabled=!eligible;box.checked=eligible;label.classList.toggle("disabled",!eligible);
  label.title=eligible?"산업용(을)·EV의 주말(토·일·공휴일) 11~14시 전력량요금 할인":"해당 종별에는 주말할인이 적용되지 않음";
}
function updateWeekendDiscountText(){
  const r=state.discountRates;if(!r)return;
  const on=SEASONS.filter(s=>r[s]>0).map(s=>`${seasonLabel(s)} ${pct(r[s],r[s]*1000%10?1:0)}`);
  $("weekendDiscountText").textContent=on.length?`주말 할인 적용(${on.join(", ")})`:"주말 할인 적용(할인율 0%)";
  document.querySelectorAll(".disc-rate").forEach(i=>{if(document.activeElement!==i)i.value=Math.round(r[i.dataset.season]*1000)/10});
}

function weightYears(){
  const y=$("year").value;
  if(y==="AVG")return FULL_YEARS;
  const n=number(y),available=[...new Set(state.weights.map(r=>number(r["연도"])).filter(Boolean))].sort((a,b)=>a-b);
  if(available.includes(n))return[n];
  const previous=available.filter(x=>x<=n).at(-1);return previous?[previous]:[];
}
function weightRows(cat){const years=weightYears();return state.weights.filter(r=>text(r["종별ID"])===cat&&years.includes(number(r["연도"])))}
function officialSales(cat,year){return state.weights.filter(r=>text(r["종별ID"])===cat&&number(r["연도"])===year).reduce((s,r)=>s+number(r["판매량(kWh)"]),0)}
function rawNationalAnnual(cat,year){return ["LAND","JEJU"].reduce((sum,rg)=>sum+(state.usage[rg][cat]||[]).filter(r=>r.year===year).reduce((s,r)=>s+r.hours.reduce((a,b)=>a+b,0),0),0)}
function annualCalibrationFactor(cat,year){
  if(!$("useAnnualCalibration").checked)return 1;
  const off=officialSales(cat,year),raw=rawNationalAnnual(cat,year);
  return off>0&&raw>0?off/raw:1;
}
function specificShare(cat,year){
  const yearsAvailable=[...new Set(state.weights.filter(r=>text(r["종별ID"])===cat).map(r=>number(r["연도"])))].sort((a,b)=>a-b);
  const wy=yearsAvailable.includes(year)?year:yearsAvailable.filter(y=>y<=year).at(-1);
  if(!wy)return 1;
  const rows=state.weights.filter(r=>text(r["종별ID"])===cat&&number(r["연도"])===wy),total=rows.reduce((s,r)=>s+number(r["판매량(kWh)"]),0);
  if(!total)return 1;
  let match=rows.filter(r=>text(r["계약전력구간"])===$("contract").value&&text(r["전압구분"])===$("voltage").value);
  // 교육용은 선택요금별 판매실적이 없으므로 전압 비중만 적용
  if(cat!=="EDU_EUL"&&match.some(r=>text(r["선택요금"])!=="미구분"))match=match.filter(r=>text(r["선택요금"])===$("choice").value);
  const share=match.reduce((s,r)=>s+number(r["판매량(kWh)"]),0)/total;
  return share>0?share:1;
}
function usageScale(cat,year){return annualCalibrationFactor(cat,year)*(($("customerScope").value==="specific"&&selectedCategoryId()!=="ALL_TOU")?specificShare(cat,year):1)}

function buildRates(cat,version){
  let rows=categoryTariffs(cat).filter(r=>text(r["버전ID"])===version);
  const specific=$("customerScope").value==="specific"&&selectedCategoryId()!=="ALL_TOU";
  if(specific)rows=rows.filter(r=>text(r["계약전력구간"])===$("contract").value&&text(r["전압구분"])===$("voltage").value&&text(r["선택요금"])===$("choice").value);
  const out={};SEASONS.forEach(s=>out[s]={});
  if(specific){
    for(const s of SEASONS)for(const p of PERIODS){const list=rows.filter(r=>text(r["계절"])===s&&text(r["부하시간대"])===p).map(r=>number(r["전력량요금"]));out[s][p]=list.length?list.reduce((a,b)=>a+b,0)/list.length:0}
    return out;
  }
  const wr=weightRows(cat);
  for(const s of SEASONS)for(const p of PERIODS){
    const candidates=rows.filter(r=>text(r["계절"])===s&&text(r["부하시간대"])===p);
    let weighted=0,totalW=0;
    for(const w of wr){
      let m=candidates.filter(t=>text(t["계약전력구간"])===text(w["계약전력구간"])&&text(t["전압구분"])===text(w["전압구분"]));
      const ch=text(w["선택요금"]);
      if(ch&&ch!=="미구분"){
        const exact=m.filter(t=>text(t["선택요금"])===ch);
        if(exact.length)m=exact; // 일반용(갑)Ⅱ 선택Ⅲ처럼 현 분석기간 요금표와 미매칭이면 전압 내 평균
      }
      if(!m.length)continue;
      const rate=m.reduce((a,t)=>a+number(t["전력량요금"]),0)/m.length,wgt=number(w["판매량(kWh)"]);
      weighted+=rate*wgt;totalW+=wgt;
    }
    if(totalW)out[s][p]=Math.round(weighted/totalW*10)/10;
    else{const list=candidates.map(r=>number(r["전력량요금"]));out[s][p]=list.length?Math.round(list.reduce((a,b)=>a+b,0)/list.length*10)/10:0}
  }
  return out;
}
function aggregateRateTable(version){
  const out={};SEASONS.forEach(s=>out[s]={});
  const catWeights={};
  for(const cat of SINGLE_CATEGORY_IDS){
    const ys=weightYears(),official=ys.reduce((sum,y)=>sum+officialSales(cat,y),0);
    const raw=selectedRegions().reduce((sum,rg)=>sum+(state.usage[rg][cat]||[]).filter(r=>analysisYears().includes(r.year)).reduce((s,r)=>s+r.hours.reduce((a,b)=>a+b,0),0),0);
    catWeights[cat]=official||raw||1;
  }
  for(const s of SEASONS)for(const p of PERIODS){let num=0,den=0;for(const cat of SINGLE_CATEGORY_IDS){const r=buildRates(cat,version)[s][p],w=catWeights[cat];if(r>0){num+=r*w;den+=w}}out[s][p]=den?Math.round(num/den*10)/10:0}
  return out;
}

function landWeekdayPeriod(version,season,h){
  if(h>=22||h<8)return"경부하";
  if(season==="동계")return(9<=h&&h<12)||(16<=h&&h<19)?"최대부하":"중간부하";
  if(version==="PRE")return(11<=h&&h<12)||(13<=h&&h<18)?"최대부하":"중간부하";
  return 15<=h&&h<21?"최대부하":"중간부하";
}
function jejuWeekdayPeriod(h){if(h>=22||h<8)return"경부하";return h<16?"중간부하":"최대부하"}
function buildDefaultSchedule(version,region){
  const schedule={};
  for(const s of SEASONS){schedule[s]={};for(const d of DAY_TYPES){schedule[s][d]=[];for(let h=0;h<24;h++){
    let p=region==="JEJU"?jejuWeekdayPeriod(h):landWeekdayPeriod(version,s,h);
    if(d==="토요일"&&p==="최대부하")p="중간부하";if(d==="일·공휴일")p="경부하";schedule[s][d].push(p);
  }}}
  return schedule;
}
const DISCOUNT_CATS=["IND_EUL","EV"];
const DISCOUNT_HOURS=[11,12,13]; // 11~14시
function discountMatrix(cat,rates){
  const d={};for(const s of SEASONS){d[s]={};for(const day of DAY_TYPES)d[s][day]=Array(24).fill(1)}
  if(DISCOUNT_CATS.includes(cat))for(const s of SEASONS)if(rates[s]>0)for(const day of ["토요일","일·공휴일"])for(const h of DISCOUNT_HOURS)d[s][day][h]=1-rates[s];
  return d;
}
function defaultDiscountRates(version){return{"하계":0,"춘추계":version==="POST"?0.5:0,"동계":0}}
// 기준안 할인은 요금표 규정 그대로(개편 후: 봄·가을철 50%), 시나리오 할인은
// 사용자가 정한(또는 9번 역산으로 찾은) 계절별 할인율을 쓴다.
function buildDiscount(cat,version){return discountMatrix(cat,defaultDiscountRates(version))}
function scenarioDiscountFor(cat){return discountMatrix(cat,state.discountRates||defaultDiscountRates($("scenarioVersion").value))}
function discountEnabled(){return state.discountForce??$("useScenarioDiscount").checked}
function emptyRateMult(){const m={};for(const s of SEASONS){m[s]={};for(const p of PERIODS)m[s][p]=1}return m}
function applyMult(r,mult){const out={};for(const s of SEASONS){out[s]={};for(const p of PERIODS)out[s][p]=r[s][p]*(mult?mult[s][p]:1)}return out}
function scenarioRatesFor(cat){return selectedCategoryId()==="ALL_TOU"?applyMult(buildRates(cat,state.allRateVersion),state.rateMult):state.scenarioRates}
function allTouDisplayRates(){const r=applyMult(aggregateRateTable(state.allRateVersion),state.rateMult);for(const s of SEASONS)for(const p of PERIODS)r[s][p]=Math.round(r[s][p]*10)/10;return r}
function scenarioScheduleForRegion(region){
  if(selectedRegion()==="NATIONAL")return region==="LAND"?state.scenarioSchedule:buildDefaultSchedule($("scenarioVersion").value,"JEJU");
  return state.scenarioSchedule;
}

function analysisYears(){const y=$("year").value;return y==="AVG"?FULL_YEARS:[number(y)]}
function yearFactor(){return $("year").value==="AVG"?1/FULL_YEARS.length:1}
function rowIncluded(r){return analysisYears().includes(r.year)&&($("scope").value==="연간"||r.season===$("scope").value)}

function loadScenarioPreset(){
  if(!state.usage.LAND.IND_EUL)return;
  setScenarioFromVersion($("scenarioVersion").value);
}
function copyBase(){setScenarioFromVersion($("baseVersion").value)}
// "기준안과 동일하게"를 눌렀을 때 전체종별은 화면 단가표만 기준안으로 바뀌고
// 실제 계산은 여전히 시나리오 버전 요금표를 쓰던 문제가 있어, 계산에 쓰는
// 요금표 버전(allRateVersion)과 주말 할인율도 함께 바꾼다.
function setScenarioFromVersion(version){
  state.scenarioSchedule=buildDefaultSchedule(version,editingRegion());
  state.allRateVersion=version;state.rateMult=emptyRateMult();
  state.scenarioRates=selectedCategoryId()==="ALL_TOU"?allTouDisplayRates():buildRates(selectedCategoryId(),version);
  state.discountRates=defaultDiscountRates(version);
  clearSolverResult();
  renderAll();
}
function renderAll(){renderHours();renderScheduleCompare();renderRates();updateWeekendDiscountText();calculate();updateWarning()}
function renderHours(){
  const day=$("dayType").value,a=state.scenarioSchedule?.[state.activeSeason]?.[day]||[];
  $("hourGrid").innerHTML=a.map((p,h)=>`<div class="hour ${PERIOD_CLASS[p]}" data-hour="${h}"><b>${String(h).padStart(2,"0")}~${String((h+1)%24).padStart(2,"0")}</b>${p}</div>`).join("");
  document.querySelectorAll(".hour").forEach(el=>el.onclick=()=>{const h=number(el.dataset.hour),c=state.scenarioSchedule[state.activeSeason][day][h];state.scenarioSchedule[state.activeSeason][day][h]=PERIODS[(PERIODS.indexOf(c)+1)%3];clearSolverResult();renderAll()});
}
function renderScheduleCompare(){
  const day=$("dayType").value;
  const base=buildDefaultSchedule($("baseVersion").value,editingRegion())[state.activeSeason][day];
  const scenario=state.scenarioSchedule[state.activeSeason][day];
  const makeRow=(label,values)=>`<div class="schedule-row"><div class="schedule-label">${label}</div><div class="schedule-hours">${values.map((p,i)=>`<div class="schedule-cell"><span>${String(i).padStart(2,"0")}</span><div class="mini ${PERIOD_CLASS[p]}">${p[0]}</div></div>`).join("")}</div></div>`;
  $("scheduleCompare").innerHTML=makeRow("기준안",base)+makeRow("시나리오",scenario);
}
function rateTable(r,editable){let h=`<thead><tr><th>계절</th>${PERIODS.map(p=>`<th>${p}</th>`).join("")}</tr></thead><tbody>`;for(const s of SEASONS){h+=`<tr><td>${s==="춘추계"?"봄·가을철":s}</td>`;for(const p of PERIODS)h+=editable?`<td><input type="number" step="0.1" data-season="${s}" data-period="${p}" value="${r[s][p].toFixed(1)}"></td>`:`<td>${r[s][p].toFixed(1)}</td>`;h+="</tr>"}return h+"</tbody>"}
function renderRates(){
  const allCat=selectedCategoryId()==="ALL_TOU",base=allCat?aggregateRateTable($("baseVersion").value):buildRates(selectedCategoryId(),$("baseVersion").value);
  if(allCat)state.scenarioRates=allTouDisplayRates();
  $("baseRates").innerHTML=rateTable(base,false);$("scenarioRates").innerHTML=rateTable(state.scenarioRates,!allCat);
  if(!allCat)document.querySelectorAll("#scenarioRates input").forEach(i=>i.oninput=()=>{state.scenarioRates[i.dataset.season][i.dataset.period]=number(i.value);clearSolverResult();calculate()});
  document.querySelectorAll(".rate-adjust").forEach(b=>b.disabled=allCat);
}

function emptyResult(){const bySeason={},byPeriod={},bySP={},win={};SEASONS.forEach(s=>{bySeason[s]={usage:0,rev:0};win[s]=0;bySP[s]={};PERIODS.forEach(p=>bySP[s][p]=0)});PERIODS.forEach(p=>byPeriod[p]={usage:0,rev:0});return{totalUsage:0,totalRev:0,bySeason,byPeriod,bySP,win,minDate:null,maxDate:null,breakdown:{}}}
function mergeResult(dst,src,cat){dst.totalUsage+=src.totalUsage;dst.totalRev+=src.totalRev;if(!dst.minDate||src.minDate<dst.minDate)dst.minDate=src.minDate;if(!dst.maxDate||src.maxDate>dst.maxDate)dst.maxDate=src.maxDate;for(const s of SEASONS){dst.bySeason[s].usage+=src.bySeason[s].usage;dst.bySeason[s].rev+=src.bySeason[s].rev;dst.win[s]+=src.win[s];for(const p of PERIODS)dst.bySP[s][p]+=src.bySP[s][p]}for(const p of PERIODS){dst.byPeriod[p].usage+=src.byPeriod[p].usage;dst.byPeriod[p].rev+=src.byPeriod[p].rev}dst.breakdown[cat]=(dst.breakdown[cat]||0)+src.totalRev}
function calcPart(cat,region,kind){
  const rows=(state.usage[region][cat]||[]).filter(rowIncluded),factor=yearFactor(),res=emptyResult();
  const baseVer=$("baseVersion").value;
  const schedule=kind==="base"?buildDefaultSchedule(baseVer,region):scenarioScheduleForRegion(region);
  const rates=kind==="base"||kind==="schedule"?buildRates(cat,baseVer):scenarioRatesFor(cat);
  const discount=kind==="final"?scenarioDiscountFor(cat):buildDiscount(cat,baseVer);
  const useDiscount=kind!=="final"||discountEnabled();
  const eligible=DISCOUNT_CATS.includes(cat),scaleCache={};
  for(const r of rows){
    if(!res.minDate||r.date<res.minDate)res.minDate=r.date;if(!res.maxDate||r.date>res.maxDate)res.maxDate=r.date;
    const scale=scaleCache[r.year]??=factor*usageScale(cat,r.year),weekend=r.dayType!=="평일",sp=res.bySP[r.season];
    for(let h=0;h<24;h++){
      const q=r.hours[h]*scale,p=schedule[r.season][r.dayType][h],dm=useDiscount?discount[r.season][r.dayType][h]:1,gross=q*rates[r.season][p],rev=gross*dm;
      res.totalUsage+=q;res.totalRev+=rev;res.bySeason[r.season].usage+=q;res.bySeason[r.season].rev+=rev;res.byPeriod[p].usage+=q;res.byPeriod[p].rev+=rev;sp[p]+=rev;
      // 주말 할인 구간(할인 전 금액) — 9번 역산에서 할인율 1%p당 매출 변화를 구하는 데 쓴다
      if(eligible&&weekend&&h>=11&&h<14)res.win[r.season]+=gross;
    }
  }
  return res;
}
function calcVariant(kind){const res=emptyResult();for(const cat of selectedCats())for(const rg of selectedRegions())mergeResult(res,calcPart(cat,rg,kind),cat);return res}
function computeEffects(){const base=calcVariant("base"),sch=calcVariant("schedule"),rat=calcVariant("rate"),fin=calcVariant("final");return{base,sch,rat,fin,ds:sch.totalRev-base.totalRev,dr:rat.totalRev-sch.totalRev,dd:fin.totalRev-rat.totalRev}}
function calculate(){
  if(!state.scenarioSchedule||!state.scenarioRates)return;
  const e=computeEffects(),{base,fin,ds,dr,dd}=e,delta=fin.totalRev-base.totalRev;
  setKpiValue($("kUsage"),splitEnergyTWh(fin.totalUsage));$("kPeriod").textContent=`${ymd(fin.minDate)}~${ymd(fin.maxDate)}`;setKpiValue($("kBase"),splitWon(base.totalRev));setKpiValue($("kScenario"),splitWon(fin.totalRev));setKpiValue($("kDelta"),signedSplitWon(delta));$("kDelta").className=cssSign(delta);$("kDeltaPct").textContent=base.totalRev?`${delta>=0?"+":""}${(delta/base.totalRev*100).toFixed(1)}%`:"-";
  const bAvg=base.totalUsage?base.totalRev/base.totalUsage:0,sAvg=fin.totalUsage?fin.totalRev/fin.totalUsage:0,dAvg=sAvg-bAvg;
  $("kBaseAvg").textContent=base.totalUsage?bAvg.toFixed(1):"-";$("kScenarioAvg").textContent=fin.totalUsage?sAvg.toFixed(1):"-";
  setKpiValue($("kAvgDiff"),{value:(dAvg>=0?"+":"")+dAvg.toFixed(1),unit:"원/kWh"});$("kAvgDiff").className=cssSign(dAvg);$("kAvgDiffPct").textContent=bAvg?`${dAvg>=0?"+":""}${(dAvg/bAvg*100).toFixed(1)}% (시나리오 − 기준)`:"-";
  const bPeak=base.totalUsage?base.byPeriod["최대부하"].usage/base.totalUsage:0,sPeak=fin.totalUsage?fin.byPeriod["최대부하"].usage/fin.totalUsage:0,dPeak=(sPeak-bPeak)*100;
  setKpiValue($("kPeakShare"),{value:(sPeak*100).toFixed(1),unit:"%"});$("kPeakShareNote").textContent=`기준 ${(bPeak*100).toFixed(1)}% → ${dPeak>=0?"+":""}${dPeak.toFixed(1)}%p`;
  for(const [id,v] of [["dSchedule",ds],["dRate",dr],["dDiscount",dd]]){$(id).textContent=signedWon(v);$(id).className=cssSign(v)}
  renderSeasonTable(base,fin);renderPeriodTable(base,fin);renderChart();state.lastResult={base,fin,ds,dr,dd};
  updateSolverCurrent(e);
}
function renderSeasonTable(b,f){let h="<thead><tr><th>계절</th><th>기준안</th><th>시나리오</th><th>증감</th></tr></thead><tbody>";for(const s of SEASONS){const d=f.bySeason[s].rev-b.bySeason[s].rev;h+=`<tr><td>${s==="춘추계"?"봄·가을철":s}</td><td>${fmtWon(b.bySeason[s].rev)}</td><td>${fmtWon(f.bySeason[s].rev)}</td><td class="${cssSign(d)}">${signedWon(d)}</td></tr>`}$("seasonTable").innerHTML=h+"</tbody>"}
function renderPeriodTable(b,f){let h="<thead><tr><th>시간대</th><th>기준 사용량</th><th>시나리오 사용량</th><th>변화</th></tr></thead><tbody>";for(const p of PERIODS){const d=f.byPeriod[p].usage-b.byPeriod[p].usage;h+=`<tr><td>${p}</td><td>${fmtEnergy(b.byPeriod[p].usage)}</td><td>${fmtEnergy(f.byPeriod[p].usage)}</td><td class="${cssSign(d)}">${d>=0?"+":""}${fmtEnergy(d)}</td></tr>`}$("periodTable").innerHTML=h+"</tbody>"}

function aggregatedDailyRows(){
  const map=new Map(),factor=yearFactor();
  for(const cat of selectedCats())for(const rg of selectedRegions())for(const r of (state.usage[rg][cat]||[]).filter(rowIncluded)){
    const key=dateKey(r.date);if(!map.has(key))map.set(key,{date:r.date,year:r.year,month:r.month,season:r.season,dayType:r.dayType,hours:Array(24).fill(0)});
    const dst=map.get(key),scale=factor*usageScale(cat,r.year);for(let h=0;h<24;h++)dst.hours[h]+=r.hours[h]*scale;
  }
  return[...map.values()].sort((a,b)=>a.date-b.date);
}
function getGraphDaySeries(dayType){const rows=aggregatedDailyRows(),selected=dayType==="전체"?rows:rows.filter(r=>r.dayType===dayType),hourUsage=Array(24).fill(0);for(const r of selected)for(let h=0;h<24;h++)hourUsage[h]+=r.hours[h];return{vals:hourUsage.map(v=>selected.length?v/selected.length:0),rowCount:selected.length}}
function getYAxisBounds(vals){const dataMin=Math.min(...vals),dataMax=Math.max(...vals),mode=$("yAxisMode").value;let yMin=0,yMax=Math.max(dataMax*1.08,1);if(mode==="zoom"){const span=Math.max(dataMax-dataMin,dataMax*.02,1);yMin=Math.max(0,dataMin-span*.18);yMax=dataMax+span*.18}else if(mode==="custom"){const a=number($("yAxisMin").value)*1e6,b=number($("yAxisMax").value)*1e6;if(b>a){yMin=a;yMax=b}else{const span=Math.max(dataMax-dataMin,dataMax*.02,1);yMin=Math.max(0,dataMin-span*.18);yMax=dataMax+span*.18}}if(yMax<=yMin)yMax=yMin+1;if(mode!=="custom"){$("yAxisMin").value=(yMin/1e6).toFixed(1);$("yAxisMax").value=(yMax/1e6).toFixed(1)}$("yAxisInfo").textContent=`현재 세로축 ${(yMin/1e6).toFixed(1)}~${(yMax/1e6).toFixed(1)}GWh/일`;return{yMin,yMax}}
function updateYAxisControlState(){const custom=$("yAxisMode").value==="custom";$("yAxisMin").disabled=!custom;$("yAxisMax").disabled=!custom}
function renderChart(){
  // 화면이 넓어질수록(특히 브라우저 전체화면) svg가 실제로 렌더링되는 폭은
  // CSS(.chart{width:100%})를 따라 계속 넓어지는데, viewBox 폭은 900으로
  // 고정돼 있고 preserveAspectRatio="none"이 가로세로를 각각 다른 배율로
  // 늘려 채웠다 — 그 결과 안의 숫자 글자(<text>)도 가로로만 길게 늘어나
  // 보기 흉해지는 문제가 있었다. 고정폭(900) 대신 svg의 실제 렌더링 폭을
  // 읽어 viewBox 폭으로 그대로 쓰면 1좌표=1픽셀이 유지되어(preserveAspectRatio
  // 기본값도 그대로 두면 되지만, x/y 배율이 이제 같아지므로 사실상 의미가
  // 없어진다) 글자·선이 가로세로 어느 쪽으로도 왜곡되지 않는다. 폭을 못
  // 구하는 경우(예: 렌더링 전)에는 기존 900을 그대로 쓴다.
  const svg=$("loadChart"),w=svg.getBoundingClientRect().width||900,h=285,l=64,rr=18,t=36,b=35,day=state.graphDayType,graph=getGraphDaySeries(day),vals=graph.vals;
  svg.setAttribute("viewBox",`0 0 ${w} ${h}`);
  const regionLabel=$("region").selectedOptions[0]?.textContent||"",catLabel=$("category").selectedOptions[0]?.textContent||"",yearLabel=$("year").selectedOptions[0]?.textContent||"",scopeLabel=$("scope").selectedOptions[0]?.textContent||"";
  $("graphDayInfo").textContent=`${regionLabel} · ${catLabel} · ${yearLabel} · ${scopeLabel} · ${day==="전체"?"전체 요일":day} ${graph.rowCount.toLocaleString()}일 평균`;
  if(!graph.rowCount){svg.innerHTML='<text x="450" y="145" text-anchor="middle" font-size="14" fill="#5f7267">선택한 조건의 부하자료가 없음</text>';return}
  const {yMin,yMax}=getYAxisBounds(vals),x=i=>l+i*(w-l-rr)/23,y=v=>h-b-(v-yMin)/(yMax-yMin)*(h-t-b);let out="";
  for(let i=0;i<24;i++){const xx=l+i*(w-l-rr)/24,ww=(w-l-rr)/24;if(day==="전체")out+=`<rect x="${xx}" y="3" width="${ww-1}" height="19" fill="#eaf1ec"/><text x="${xx+ww/2}" y="16" text-anchor="middle" font-size="8" fill="#5f7267">전체</text>`;else{const p=state.scenarioSchedule[state.activeSeason][day][i];out+=`<rect x="${xx}" y="3" width="${ww-1}" height="19" fill="var(--${PERIOD_CLASS[p]})"/><text x="${xx+ww/2}" y="16" text-anchor="middle" font-size="8">${p[0]}</text>`}}
  for(let k=0;k<=4;k++){const ratio=k/4,yy=t+(h-t-b)*ratio,tick=yMax-(yMax-yMin)*ratio;out+=`<line x1="${l}" y1="${yy}" x2="${w-rr}" y2="${yy}" stroke="#dde6e0"/><text x="${l-7}" y="${yy+4}" text-anchor="end" font-size="9" fill="#5f7267">${(tick/1e6).toFixed(1)}</text>`}
  out+=`<text x="8" y="${t-7}" font-size="9" fill="#5f7267">GWh/일</text><polyline points="${vals.map((v,i)=>`${x(i)},${y(v)}`).join(" ")}" fill="none" stroke="#1f8a5f" stroke-width="3"/>`;
  // 요청에 따라 선 위의 원형 마크(점)는 제거했다. 시간 눈금은 그래프가 화면
  // 전체 폭으로 넓어져(4번 패널 전체 폭 배치) 24칸이 다 들어갈 공간이 있으므로,
  // 짝수 시간만 걸러내던 것을 없애고 0~23시 전체를 표시한다(글자가 24개라
  // 폰트 크기만 살짝 줄여 서로 겹치지 않게 한다).
  vals.forEach((v,i)=>{out+=`<text x="${x(i)}" y="${h-13}" text-anchor="middle" font-size="8" fill="#5f7267">${i}</text>`});svg.innerHTML=out;
}

function updateWarning(){
  if(!state.usage.LAND.IND_EUL)return;
  const msgs=[],rg=selectedRegion(),cat=selectedCategoryId(),y=$("year").value;
  msgs.push(`<b>지역:</b> ${$("region").selectedOptions[0].textContent}`);
  if(rg!=="NATIONAL")msgs.push(`세부 전압·선택요금 구성은 지역별 자료가 없어 <b>전국 판매량 비중</b>을 대리 적용함`);
  if(rg==="NATIONAL")msgs.push(`전국 사용량은 동일 일자·시간의 <b>육지+제주</b>를 합산함`);
  if(cat==="ALL_TOU")msgs.push(`<b>전체종별:</b> 6개 종별을 각각 해당 요금으로 계산한 후 매출을 합산함 · 단가 직접수정은 비활성화`);
  if(cat==="EV")msgs.push(`EV 세부요금별 판매량 비중 자료가 없어 전체고객은 등록 EV요금의 단순평균을 적용함`);
  if(cat==="EDU_EUL")msgs.push(`교육용(을)은 선택요금별 판매량이 없어 전압별 판매량 가중 후 선택요금은 평균 처리함`);
  if(cat==="GEN_GAP2"&&$("customerScope").value==="all")msgs.push(`일반용(갑)Ⅱ의 판매실적 중 현 분석기간 요금표와 직접 매칭되지 않는 선택Ⅲ은 동일 전압 선택Ⅰ·Ⅱ 평균으로 처리함`);
  if(y==="2026")msgs.push(`<b>2026년:</b> 1~6월만 분석 · 세부요금 구성은 2025년 전국 판매량 비중 사용 · 공식 연간 판매량 보정은 미적용`);
  if(y==="AVG")msgs.push(`연평균은 완전연도인 <b>2022~2025년</b>만 사용함`);
  if(rg==="NATIONAL")msgs.push(`전국 분석에서 시간대 셀 직접수정은 <b>육지 시간대</b>에 적용하며 제주 고유 시간대는 유지함`);
  $("warning").innerHTML=msgs.join(" · ");
}

function exportCsv(){
  const r=state.lastResult;if(!r)return;
  const rows=[["항목","값"],["분석지역",$("region").selectedOptions[0].textContent],["분석종별",$("category").selectedOptions[0].textContent],["연도",$("year").selectedOptions[0].textContent],["범위",$("scope").selectedOptions[0].textContent],["고객범위",$("customerScope").selectedOptions[0].textContent],["계약전력구간",$("contract").value],["전압구분",$("voltage").value],["선택요금",$("choice").value],["기준안 매출(원)",r.base.totalRev],["시나리오 매출(원)",r.fin.totalRev],["증감액(원)",r.fin.totalRev-r.base.totalRev],["시간대 효과(원)",r.ds],["단가 효과(원)",r.dr],["주말 할인 효과(원)",r.dd],["기준 평균단가(원/kWh)",r.base.totalUsage?r.base.totalRev/r.base.totalUsage:""],["시나리오 평균단가(원/kWh)",r.fin.totalUsage?r.fin.totalRev/r.fin.totalUsage:""],["시나리오 주말 할인율(하계/봄·가을/겨울)",SEASONS.map(x=>pct(state.discountRates[x])).join(" / ")],[],["계절","기준안(원)","시나리오(원)","증감(원)"]];
  for(const s of SEASONS)rows.push([s,r.base.bySeason[s].rev,r.fin.bySeason[s].rev,r.fin.bySeason[s].rev-r.base.bySeason[s].rev]);
  if(selectedCategoryId()==="ALL_TOU"){rows.push([], ["종별","기준안(원)","시나리오(원)","증감(원)"]);for(const cat of SINGLE_CATEGORY_IDS)rows.push([catMeta(cat).name,r.base.breakdown[cat]||0,r.fin.breakdown[cat]||0,(r.fin.breakdown[cat]||0)-(r.base.breakdown[cat]||0)])}
  const csv="\ufeff"+rows.map(row=>row.map(v=>`"${String(v??"").replace(/"/g,'""')}"`).join(",")).join("\r\n"),blob=new Blob([csv],{type:"text/csv;charset=utf-8"}),a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="PRAS_TOU_매출_분석결과.csv";a.click();URL.revokeObjectURL(a.href);
}

// ───────────────────────── 9. 목표 효과 역산 ─────────────────────────
// 5번 요인분해는 "시간대 변경 → 단가 변경 → 주말 할인" 순서의 순차효과이므로,
// 역산도 같은 순서로 하나씩 맞춘다. 앞 단계(시간대)를 바꾸면 뒤 단계(단가·할인)
// 효과도 함께 움직이기 때문에, 각 단계는 앞 단계 결과를 반영한 상태에서 푼다.
// 모든 조정은 선택한 계절 안에서만 일어난다.
const SOLVER_KEYS=["scenarioSchedule","scenarioRates","rateMult","discountRates","discountForce"];
function withOverride(o,fn){const saved={};for(const k of SOLVER_KEYS){saved[k]=state[k];if(k in o)state[k]=o[k]}try{return fn()}finally{Object.assign(state,saved)}}
function solverSeasons(){const v=$("solveSeason").value;return v==="ALL"?SEASONS:[v]}
function solverBands(){return[...document.querySelectorAll(".solve-band")].filter(x=>x.checked).map(x=>x.value)}
function readTarget(id){const v=$(id).value.trim();return v===""?null:number(v)*1e8}
function scheduleRegions(){return selectedRegion()==="NATIONAL"?["LAND"]:[selectedRegion()]}
const satBand=b=>b==="최대부하"?"중간부하":b;

// 시간대 효과 = 시나리오 시간대 + 기준 단가 + 기준 할인. 시간 h를 어느 시간대로
// 두느냐에 따른 매출을 시간·시간대별로 미리 합산해 두면 시간대를 바꿀 때의
// 효과 변화를 즉시 더하고 뺄 수 있다(부하이동이 없으므로 정확히 더해짐).
function scheduleContrib(seasons){
  const A={},baseVer=$("baseVersion").value,factor=yearFactor();
  for(const s of seasons){A[s]={};for(const d of ["평일","토요일"])A[s][d]=Array.from({length:24},()=>({"경부하":0,"중간부하":0,"최대부하":0}))}
  for(const cat of selectedCats()){
    const rates=buildRates(cat,baseVer),disc=buildDiscount(cat,baseVer),scaleCache={};
    for(const rg of scheduleRegions())for(const r of (state.usage[rg][cat]||[])){
      if(r.dayType==="일·공휴일"||!seasons.includes(r.season)||!rowIncluded(r))continue;
      const scale=scaleCache[r.year]??=factor*usageScale(cat,r.year),cell=A[r.season][r.dayType],rt=rates[r.season];
      for(let h=0;h<24;h++){const q=r.hours[h]*scale*disc[r.season][r.dayType][h];if(!q)continue;const c=cell[h];c["경부하"]+=q*rt["경부하"];c["중간부하"]+=q*rt["중간부하"];c["최대부하"]+=q*rt["최대부하"]}
    }
  }
  return A;
}
// 평일 시간대를 시간 단위로 바꾸고 토요일은 요금표 규칙(평일 최대부하 → 토요일
// 중간부하)을 따른다. 일·공휴일은 종일 경부하 그대로. 목표에 가장 가깝게
// 맞추되, 오차가 허용범위 안이면 바꾸는 시간 수가 가장 적은 안을 고른다(동적계획법).
function solveSchedule(targetDelta,seasons,keepNight){
  const A=scheduleContrib(seasons),sched=state.scenarioSchedule,vars=[];
  for(const s of seasons)for(let h=0;h<24;h++){
    if(keepNight&&(h>=22||h<8))continue;
    const curW=sched[s]["평일"][h],curS=sched[s]["토요일"][h],cur=A[s]["평일"][h][curW]+A[s]["토요일"][h][curS],opts=[];
    for(const b of PERIODS){if(b===curW||(keepNight&&b==="경부하"))continue;const val=A[s]["평일"][h][b]+A[s]["토요일"][h][satBand(b)]-cur;opts.push({b,sat:satBand(b),val})}
    if(opts.some(o=>Math.abs(o.val)>1))vars.push({s,h,curW,curS,opts});
  }
  let lo=0,hi=0;for(const v of vars){lo+=Math.min(0,...v.opts.map(o=>o.val));hi+=Math.max(0,...v.opts.map(o=>o.val))}
  const bin=Math.max((hi-lo)/4000,1),layers=[];let cur=new Map([[0,{c:0,sum:0}]]);
  for(const v of vars){
    const next=new Map(),relax=(k,c,sum,prev,choice)=>{const e=next.get(k);if(!e||c<e.c||(c===e.c&&Math.abs(sum-targetDelta)<Math.abs(e.sum-targetDelta)))next.set(k,{c,sum,prev,choice})};
    for(const [k,st] of cur){relax(k,st.c,st.sum,k,-1);v.opts.forEach((o,i)=>{const sum=st.sum+o.val;relax(Math.round(sum/bin),st.c+1,sum,k,i)})}
    layers.push(next);cur=next;
  }
  const tol=Math.max(bin*1.5,Math.abs(targetDelta)*0.003);let bestK=null,best=null;
  for(const [k,st] of cur){const err=Math.abs(st.sum-targetDelta),ok=err<=tol;const better=!best||(ok&&!best.ok)||(ok===best.ok&&(ok?(st.c<best.st.c||(st.c===best.st.c&&err<best.err)):err<best.err));if(better){best={st,err,ok};bestK=k}}
  const newSchedule=clone(sched),changes=[];let k=bestK;
  for(let i=layers.length-1;i>=0;i--){const st=layers[i].get(k);if(st.choice>=0){const v=vars[i],o=v.opts[st.choice];newSchedule[v.s]["평일"][v.h]=o.b;newSchedule[v.s]["토요일"][v.h]=o.sat;changes.push({s:v.s,h:v.h,from:v.curW,to:o.b,satFrom:v.curS,satTo:o.sat})}k=st.prev}
  changes.sort((a,b)=>SEASONS.indexOf(a.s)-SEASONS.indexOf(b.s)||a.h-b.h);
  return{newSchedule,changes,achieved:best.st.sum,reachable:best.ok,lo,hi};
}
// 단가 효과 = 시나리오 단가로 바꿨을 때의 매출 변화. 선택 계절·시간대 단가에
// 같은 배율 f를 곱하면 효과가 f에 정비례하므로 한 번에 풀린다.
function solveRates(target,e,seasons,bands){
  const R=seasons.reduce((s,se)=>s+bands.reduce((a,p)=>a+e.rat.bySP[se][p],0),0);
  if(!bands.length)return{error:"단가 조정 대상 시간대를 하나 이상 선택해야 함"};
  if(R<=0)return{error:"선택한 계절·시간대에 해당하는 매출이 없어 단가로 조정할 수 없음"};
  const f=1+(target-e.dr)/R;
  if(f<=0)return{error:`목표를 맞추려면 단가를 ${pct(1-f)} 이상 내려야 해 불가능함`};
  const out={f,rateMult:clone(state.rateMult),scenarioRates:clone(state.scenarioRates),before:{},after:{}};
  for(const s of seasons)for(const p of bands){
    out.before[s+p]=state.scenarioRates[s][p];
    if(selectedCategoryId()==="ALL_TOU")out.rateMult[s][p]*=f;else out.scenarioRates[s][p]=Math.round(state.scenarioRates[s][p]*f*10)/10;
  }
  if(selectedCategoryId()==="ALL_TOU")withOverride({rateMult:out.rateMult},()=>{out.scenarioRates=allTouDisplayRates()});
  for(const s of seasons)for(const p of bands)out.after[s+p]=out.scenarioRates[s][p];
  return out;
}
// 주말 할인 효과 = −(할인율 × 할인 구간 매출). 선택 계절에 같은 할인율 d를 쓰면
// 효과가 d에 정비례하므로 바로 풀린다. 0~100% 밖이면 경계값으로 맞추고 알린다.
function solveDiscount(target,e,seasons){
  if(!selectedCats().some(c=>DISCOUNT_CATS.includes(c)))return{error:"선택한 종별은 주말 할인 대상(산업용(을)·EV)이 아님"};
  const W=seasons.reduce((s,se)=>s+e.fin.win[se],0);
  if(W<=0)return{error:"선택한 계절에 주말 할인 구간(토·일·공휴일 11~14시) 매출이 없음"};
  const now=seasons.reduce((s,se)=>s+state.discountRates[se]*e.fin.win[se],0);
  let d=(e.dd-target+now)/W,clamped=null;
  if(d<0){clamped=`할인율을 0%까지 낮춰도 목표에 못 미침`;d=0}
  if(d>1){clamped=`할인율을 100%로 올려도 목표에 못 미침`;d=1}
  d=Math.round(d*1000)/1000;
  const rates={...state.discountRates};for(const s of seasons)rates[s]=d;
  return{d,rates,clamped,before:Object.fromEntries(seasons.map(s=>[s,state.discountRates[s]]))};
}
function runSolver(){
  const box=$("solverResult");
  if(!state.lastResult){box.innerHTML="";return}
  const seasons=solverSeasons(),targets={ds:readTarget("targetSchedule"),dr:readTarget("targetRate"),dd:readTarget("targetDiscount")};
  if(targets.ds===null&&targets.dr===null&&targets.dd===null){box.innerHTML='<div class="callout warning">목표값을 하나 이상 입력해야 함(억원 단위)</div>';return}
  if($("scope").value!=="연간"&&!seasons.includes($("scope").value)){box.innerHTML=`<div class="callout warning">1번 분석 범위가 ${seasonLabel($("scope").value)}이므로 다른 계절은 조정해도 결과가 바뀌지 않음. 조정 계절을 맞추거나 분석 범위를 연간으로 바꿔야 함</div>`;return}
  const before=computeEffects(),p={scenarioSchedule:clone(state.scenarioSchedule),scenarioRates:clone(state.scenarioRates),rateMult:clone(state.rateMult),discountRates:{...state.discountRates},discountForce:null},detail={};
  if(targets.ds!==null){const r=withOverride(p,()=>solveSchedule(targets.ds-before.ds,seasons,$("solveKeepNight").checked));p.scenarioSchedule=r.newSchedule;detail.ds=r}
  if(targets.dr!==null){const e=withOverride(p,computeEffects),r=withOverride(p,()=>solveRates(targets.dr,e,seasons,solverBands()));if(!r.error){p.rateMult=r.rateMult;p.scenarioRates=r.scenarioRates}detail.dr=r}
  if(targets.dd!==null){p.discountForce=true;const e=withOverride(p,computeEffects),r=withOverride(p,()=>solveDiscount(targets.dd,e,seasons));if(!r.error)p.discountRates=r.rates;detail.dd=r}
  const after=withOverride(p,computeEffects);
  state.solverProposal={p,seasons};
  renderSolverResult(before,after,targets,detail,seasons);
  $("applySolution").disabled=false;
}
function renderSolverResult(before,after,t,detail,seasons){
  const row=(label,key)=>{const tv=t[key],av=after[key],gap=tv===null?null:av-tv;return`<tr><td>${label}</td><td>${signedEok(before[key])}</td><td>${tv===null?'<span class="muted">유지</span>':signedEok(tv)}</td><td class="${cssSign(av)}"><b>${signedEok(av)}</b></td><td>${gap===null?"-":(Math.abs(gap)<=Math.max(5e6,Math.abs(tv)*0.005)?"달성":signedEok(gap))}</td></tr>`};
  const tot=e=>e.fin.totalRev-e.base.totalRev;
  let h=`<h3>계산 결과(${seasons.map(seasonLabel).join("·")} 조정)</h3><div class="scroll"><table class="table solver-out"><thead><tr><th>효과(억원)</th><th>현재</th><th>목표</th><th>조정 후</th><th>목표와 차이</th></tr></thead><tbody>${row("시간대 구분 효과","ds")}${row("단가 조정 효과","dr")}${row("주말 할인 효과","dd")}<tr class="total"><td>매출 증감 합계</td><td>${signedEok(tot(before))}</td><td>-</td><td class="${cssSign(tot(after))}"><b>${signedEok(tot(after))}</b></td><td>-</td></tr></tbody></table></div>`;
  h+=`<div class="solver-levers">`;
  if(detail.ds){const r=detail.ds;h+=`<div class="lever"><b>① 시간대 조정</b>`;
    if(!r.changes.length)h+=`<p>바꿀 필요 없음(현재 시간대 유지)</p>`;
    else{h+=`<p>평일 ${r.changes.length}개 시간 변경(토요일은 평일 최대부하를 중간부하로 두는 규칙 적용, 일·공휴일은 그대로)</p><ul>${r.changes.map(c=>`<li>${seasonLabel(c.s)} ${String(c.h).padStart(2,"0")}~${String((c.h+1)%24).padStart(2,"0")}시: <span class="pill ${PERIOD_CLASS[c.from]}">${c.from}</span> → <span class="pill ${PERIOD_CLASS[c.to]}">${c.to}</span>${c.satFrom!==c.satTo?` <small>(토 ${c.satFrom}→${c.satTo})</small>`:""}</li>`).join("")}</ul>`}
    if(!r.reachable)h+=`<p class="warn-text">선택 계절 시간대만으로는 목표에 닿지 않아 가장 가까운 값으로 맞춤(가능 범위: 현재 대비 ${signedEok(r.lo)} ~ ${signedEok(r.hi)}억원)</p>`;
    h+=`</div>`}
  if(detail.dr){const r=detail.dr;h+=`<div class="lever"><b>② 단가 조정</b>`;
    if(r.error)h+=`<p class="warn-text">${r.error}</p>`;
    else{h+=`<p>선택 계절 ${solverBands().join("·")} 단가 ${r.f>=1?"+":""}${((r.f-1)*100).toFixed(2)}% 조정</p><ul>${Object.keys(r.after).map(k=>{const s=SEASONS.find(x=>k.startsWith(x)),pp=k.slice(s.length);return`<li>${seasonLabel(s)} ${pp}: ${r.before[k].toFixed(1)} → <b>${r.after[k].toFixed(1)}</b> 원/kWh</li>`}).join("")}</ul>`;
      if(selectedCategoryId()==="ALL_TOU")h+=`<p class="muted">전체종별은 6개 종별 단가에 같은 비율을 곱함(표의 값은 판매량 가중평균)</p>`}
    h+=`</div>`}
  if(detail.dd){const r=detail.dd;h+=`<div class="lever"><b>③ 주말 할인율 조정</b>`;
    if(r.error)h+=`<p class="warn-text">${r.error}</p>`;
    else{h+=`<ul>${Object.keys(r.before).map(s=>`<li>${seasonLabel(s)} 주말 11~14시 할인율: ${pct(r.before[s])} → <b>${pct(r.d)}</b></li>`).join("")}</ul>`;if(r.clamped)h+=`<p class="warn-text">${r.clamped}</p>`}
    h+=`</div>`}
  h+=`</div><p class="muted">시간대 → 단가 → 할인 순서로 하나씩 맞춤. 앞 단계를 바꾸면 뒤 효과도 같이 움직이므로, 목표를 비워둔 효과도 값이 달라질 수 있음. 단가는 0.1원, 할인율은 0.1%p 단위로 반올림해 목표와 약간 차이가 날 수 있음.</p>`;
  $("solverResult").innerHTML=h;
}
function applySolution(){
  const sp=state.solverProposal;if(!sp)return;
  const {discountForce,...rest}=sp.p;Object.assign(state,clone(rest));
  if(discountForce)$("useScenarioDiscount").checked=true;
  state.activeSeason=sp.seasons[0];document.querySelectorAll("#seasonTabs .tab").forEach(x=>x.classList.toggle("active",x.dataset.season===state.activeSeason));
  state.solverProposal=null;$("applySolution").disabled=true;
  renderAll();
  $("solverResult").insertAdjacentHTML("afterbegin",'<div class="callout">조정안을 시나리오에 적용함. 2·3번 패널과 5번 요인분해에서 바뀐 내용을 확인할 수 있음.</div>');
}
function clearSolverResult(){state.solverProposal=null;if($("applySolution"))$("applySolution").disabled=true;if($("solverResult"))$("solverResult").innerHTML=""}
function updateSolverCurrent(e){
  $("solveCurSchedule").textContent=signedEok(e.ds);$("solveCurRate").textContent=signedEok(e.dr);$("solveCurDiscount").textContent=signedEok(e.dd);
  for(const [id,v] of [["targetSchedule",e.ds],["targetRate",e.dr],["targetDiscount",e.dd]])$(id).placeholder=`현재 ${eok(v)}`;
}
$("solveTargets").onclick=runSolver;
$("applySolution").onclick=applySolution;
["solveSeason","solveKeepNight"].forEach(id=>$(id).addEventListener("change",clearSolverResult));
document.querySelectorAll(".solve-band").forEach(x=>x.addEventListener("change",clearSolverResult));
document.querySelectorAll(".disc-rate").forEach(i=>i.onchange=()=>{const v=Math.min(100,Math.max(0,number(i.value)));state.discountRates[i.dataset.season]=v/100;i.value=v;clearSolverResult();updateWeekendDiscountText();calculate()});

$("region").onchange=()=>{loadScenarioPreset();updateWarning()};
$("category").onchange=selectCategory;
$("customerScope").onchange=()=>{updateCustomerScopeControls();loadScenarioPreset();updateWarning()};
$("contract").onchange=()=>{populateVoltage();loadScenarioPreset();updateWarning()};
$("voltage").onchange=()=>{populateChoice();loadScenarioPreset();updateWarning()};
$("choice").onchange=()=>{loadScenarioPreset();updateWarning()};
$("baseVersion").onchange=()=>{clearSolverResult();renderAll()};
$("scenarioVersion").onchange=loadScenarioPreset;
$("year").onchange=()=>{loadScenarioPreset();updateWarning()};
$("scope").onchange=()=>{clearSolverResult();calculate();updateWarning()};
$("dayType").onchange=renderAll;
$("useScenarioDiscount").onchange=()=>{clearSolverResult();calculate()};
$("useAnnualCalibration").onchange=()=>{clearSolverResult();calculate();updateWarning()};
$("loadScenario").onclick=loadScenarioPreset;
$("copyBase").onclick=copyBase;
$("exportCsv").onclick=exportCsv;
$("yAxisMode").onchange=()=>{updateYAxisControlState();if(state.lastResult)renderChart()};
$("applyYAxis").onclick=()=>{if(state.lastResult)renderChart()};
$("yAxisMin").onchange=()=>{if(state.lastResult)renderChart()};
$("yAxisMax").onchange=()=>{if(state.lastResult)renderChart()};
document.querySelectorAll("#seasonTabs .tab").forEach(btn=>btn.onclick=()=>{state.activeSeason=btn.dataset.season;document.querySelectorAll("#seasonTabs .tab").forEach(x=>x.classList.toggle("active",x===btn));renderAll()});
document.querySelectorAll("#graphDayTabs .tab").forEach(btn=>btn.onclick=()=>{state.graphDayType=btn.dataset.graphDay;document.querySelectorAll("#graphDayTabs .tab").forEach(x=>x.classList.toggle("active",x===btn));renderChart()});
document.querySelectorAll(".rate-adjust").forEach(btn=>btn.onclick=()=>{if(selectedCategoryId()==="ALL_TOU")return;const m=number(btn.dataset.mult);for(const s of SEASONS)for(const p of PERIODS)state.scenarioRates[s][p]=Math.round(state.scenarioRates[s][p]*m*10)/10;clearSolverResult();renderAll()});

// 창 크기 변경(브라우저 전체화면 진입/해제 포함)마다 그래프의 viewBox 폭을
// 다시 맞춘다 — renderChart()가 매번 svg의 "그 시점" 렌더링 폭을 읽어
// viewBox에 반영하므로, 전체화면으로 넓어진 뒤에도 숫자 글자가 가로로
// 늘어나지 않고 항상 원래 비율로 보인다. resize 이벤트가 짧은 시간에
// 여러 번 발생할 수 있어 150ms 디바운스로 과도한 재계산을 막는다.
let resizeChartTimer=null;
window.addEventListener("resize",()=>{
  if(!state.lastResult)return;
  clearTimeout(resizeChartTimer);
  resizeChartTimer=setTimeout(renderChart,150);
});

updateYAxisControlState();
loadGithubData();
