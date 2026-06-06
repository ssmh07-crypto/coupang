const won=new Intl.NumberFormat('ko-KR',{maximumFractionDigits:0});
const money=n=>`${won.format(Math.round(n))}원`;
const minus=n=>`- ${money(n)}`;
const VAT_RATE=0.1;
const XLSX_SRC='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
const periodIds=['periodTotalQuantity','periodCategoryFeeRate','periodProductUnitCost','periodReturnCount','periodCouponUnitPrice','periodShippingCost','periodOtherCost'];
const periodEl=Object.fromEntries(periodIds.map(id=>[id,document.getElementById(id)]));
const uiIds=['periodProfit','periodOrganicQuantity','periodAdRevenue','periodReturnDisplay','periodMargin','periodProfitRate','periodSummaryState','periodProfitRing','periodBreakdown','periodFormulaNumbers','periodFormulaResult','periodAdSpend','periodSoldQuantity','periodRowCount','periodFileName','periodCampaignRange','periodUploadMessage'];
const ui=Object.fromEntries(uiIds.map(id=>[id,document.getElementById(id)]));
let adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:'',productPrice:0,fileName:''};
let totalQuantityTouched=false;
const periodVal=id=>Math.max(0,Number(periodEl[id].value)||0);
const withVat=amount=>amount*(1+VAT_RATE);
const setText=(id,text)=>{const node=ui[id]||document.getElementById(id);if(node)node.textContent=text};
let xlsxLoadPromise;
function loadXlsx(){
  if(window.XLSX)return Promise.resolve(window.XLSX);
  if(!xlsxLoadPromise){
    xlsxLoadPromise=new Promise((resolve,reject)=>{
      const script=document.createElement('script');
      script.src=XLSX_SRC;
      script.async=true;
      script.onload=()=>window.XLSX?resolve(window.XLSX):reject(new Error('엑셀 분석 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.'));
      script.onerror=()=>reject(new Error('엑셀 분석 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.'));
      document.head.append(script);
    });
  }
  return xlsxLoadPromise;
}
function periodMoney(amount){return money(amount)}
function syncTotalQuantity(){
  if(!totalQuantityTouched&&periodEl.periodTotalQuantity)periodEl.periodTotalQuantity.value=adReport.quantity;
}
function renderPeriod(){
  syncTotalQuantity();
  const adQuantity=adReport.quantity;
  const totalQuantity=periodVal('periodTotalQuantity');
  const categoryRate=periodVal('periodCategoryFeeRate');
  const unitCost=periodVal('periodProductUnitCost');
  const couponUnit=periodVal('periodCouponUnitPrice');
  const shippingUnit=periodVal('periodShippingCost');
  const other=periodVal('periodOtherCost');
  const returnCount=Math.max(0,adQuantity-totalQuantity);
  const organicQuantity=Math.max(0,totalQuantity-adQuantity);
  periodEl.periodReturnCount.value=returnCount;
  const netProductPrice=Math.max(0,adReport.productPrice-couponUnit);
  const adjustedAdRevenue=Math.max(0,adReport.revenue-couponUnit*adQuantity);
  const returnRevenue=returnCount*netProductPrice;
  const organicRevenue=organicQuantity*netProductPrice;
  const totalRevenue=adjustedAdRevenue+organicRevenue-returnRevenue;
  const adSpendVat=withVat(adReport.spend);
  const categoryFee=totalRevenue*categoryRate/100;
  const categoryFeeVat=withVat(categoryFee);
  const productCostTotal=unitCost*totalQuantity;
  const shippingTotal=shippingUnit*totalQuantity;
  const shippingTotalVat=withVat(shippingTotal);
  const profit=totalRevenue-adSpendVat-categoryFeeVat-productCostTotal-shippingTotalVat-other;
  const margin=totalRevenue?profit/totalRevenue*100:0;
  const ringValue=Math.max(0,Math.min(100,margin));
  const profitEl=ui.periodProfit;
  if(profitEl){
    profitEl.textContent=periodMoney(profit);
    profitEl.classList.toggle('loss',profit<0);
  }
  setText('periodOrganicQuantity',`${won.format(organicQuantity)}개`);
  setText('periodAdRevenue',periodMoney(adjustedAdRevenue));
  setText('periodReturnDisplay',`${won.format(returnCount)}개`);
  setText('periodMargin',`총 전환매출 대비 순이익률 ${margin.toFixed(1)}%`);
  setText('periodProfitRate',`${margin.toFixed(0)}%`);
  setText('periodSummaryState',!adReport.rows?'보고서 대기':profit>=0?'수익 구간':'손실 구간');
  const ring=ui.periodProfitRing;
  if(ring)ring.style.setProperty('--rate',`${ringValue}%`);
  const revenueTotal=totalRevenue;
  const costTotal=adSpendVat+categoryFeeVat+productCostTotal+shippingTotalVat+other;
  const breakdownGroups=[
    ['매출',[
      ['광고 전환매출',periodMoney(adjustedAdRevenue),'plus'],
      ['자연판매 매출',periodMoney(organicRevenue),'plus'],
      ['반품/취소 매출',minus(returnRevenue),'minus'],
      ['매출 합계',periodMoney(revenueTotal),'subtotal']
    ]],
    ['비용',[
      ['광고비(VAT 포함)',minus(adSpendVat),'minus'],
      [`카테고리 수수료(${categoryRate.toFixed(1)}%, VAT 포함)`,minus(categoryFeeVat),'minus'],
      ['최종 매입원가',minus(productCostTotal),'minus'],
      ['입출고비 / 배송비(VAT 포함)',minus(shippingTotalVat),'minus'],
      ['기타 비용',minus(other),'minus'],
      ['비용 합계',minus(costTotal),'subtotal']
    ]],
    ['최종',[
      ['기간 순이익',periodMoney(profit),profit<0?'loss':'total']
    ]]
  ];
  ui.periodBreakdown.innerHTML=breakdownGroups.map(([title,items])=>`<section class="breakdown-section"><h3>${title}</h3>${items.map(([label,value,type])=>`<div class="result-cost-row ${type}"><span>${label}</span><strong>${value}</strong></div>`).join('')}</section>`).join('');
  ui.periodFormulaNumbers.innerHTML=[
    ['매출 합계',`= 광고 전환매출 + 자연판매 매출 - 반품/취소 매출<br><b>= ${periodMoney(adjustedAdRevenue)} + ${periodMoney(organicRevenue)} - ${periodMoney(returnRevenue)}</b>`],
    ['비용 합계',`= 광고비 + 수수료 + 매입원가 + 입출고비 + 기타 비용<br><b>= ${periodMoney(costTotal)}</b>`],
    ['기간 순이익',`= 매출 합계 - 비용 합계<br><b>= ${periodMoney(revenueTotal)} - ${periodMoney(costTotal)}</b>`]
  ].map(([title,body])=>`<div><span>${title}</span><p>${body}</p></div>`).join('');
  setText('periodFormulaResult',`= ${periodMoney(profit)}`);
}
function reportNumber(value){if(typeof value==='number')return value;return Number(String(value??'').replace(/,/g,''))||0}
function sum(rows,key){return rows.reduce((total,row)=>total+reportNumber(row[key]),0)}
function reportText(value){return String(value??'').trim()}
function campaignRange(rows){const starts=[...new Set(rows.map(row=>reportText(row['캠페인 시작일'])).filter(value=>value&&value!=='-'))].sort(),ends=[...new Set(rows.map(row=>reportText(row['캠페인 종료일'])).filter(value=>value&&value!=='-'))].sort();return{start:starts[0]||'',end:ends.at(-1)||''}}
function renderAdReport(){
  setText('periodAdSpend',periodMoney(adReport.spend));
  setText('periodSoldQuantity',`${won.format(adReport.quantity)}개`);
  setText('periodRowCount',adReport.rows?`${won.format(adReport.rows)}개 행 분석 완료`:'0개 행');
  setText('periodFileName',adReport.fileName||'선택된 파일 없음');
  setText('periodCampaignRange',adReport.start?`${adReport.start} ~ ${adReport.end||'진행 중'}`:'보고서 업로드 후 자동 표시');
  renderPeriod();
}
async function analyzeAdReport(event){const file=event.target.files[0];if(!file)return;try{setText('periodUploadMessage','보고서를 분석하고 있습니다...');const XLSX=await loadXlsx();const workbook=XLSX.read(await file.arrayBuffer(),{type:'array'}),sheet=workbook.Sheets[workbook.SheetNames[0]],rows=XLSX.utils.sheet_to_json(sheet,{defval:0});if(!rows.length||!Object.prototype.hasOwnProperty.call(rows[0],'광고비'))throw new Error('쿠팡 광고보고서 형식을 확인할 수 없습니다. 광고비 열이 필요합니다.');const range=campaignRange(rows);adReport={spend:sum(rows,'광고비'),revenue:sum(rows,'총 전환매출액(14일)')||sum(rows,'총 전환매출액(1일)'),quantity:sum(rows,'총 판매수량(14일)')||sum(rows,'총 판매수량(1일)'),rows:rows.length,fileName:file.name,...range,productPrice:0};adReport.productPrice=adReport.quantity?adReport.revenue/adReport.quantity:0;if(!totalQuantityTouched)periodEl.periodTotalQuantity.value=adReport.quantity;setText('periodUploadMessage','보고서 분석이 완료되었습니다.');renderAdReport()}catch(error){adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:'',productPrice:0,fileName:''};setText('periodUploadMessage',error.message);renderAdReport()}}
const periodCategory=document.getElementById('periodCategory');
if(periodCategory){
  periodCategory.addEventListener('change',()=>{
    if(periodCategory.value!=='custom')periodEl.periodCategoryFeeRate.value=periodCategory.value;
    renderPeriod();
  });
}
document.getElementById('adReportFile').addEventListener('change',analyzeAdReport);
Object.entries(periodEl).forEach(([id,input])=>input.addEventListener('input',()=>{
  if(id==='periodTotalQuantity')totalQuantityTouched=true;
  renderPeriod();
}));
periodEl.periodCategoryFeeRate.addEventListener('input',()=>{
  if(periodCategory&&periodCategory.value!=='custom'&&Number(periodCategory.value)!==Number(periodEl.periodCategoryFeeRate.value))periodCategory.value='custom';
});
renderAdReport();
