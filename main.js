const won=new Intl.NumberFormat('ko-KR',{maximumFractionDigits:0});
const money=n=>`${won.format(Math.round(n))}원`;
const minus=n=>`- ${money(n)}`;
const VAT_RATE=0.1;
const periodIds=['periodTotalQuantity','periodCategoryFeeRate','periodProductUnitCost','periodReturnCount','periodCouponUnitPrice','periodShippingCost','periodOtherCost'];
const periodEl=Object.fromEntries(periodIds.map(id=>[id,document.getElementById(id)]));
let adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:'',productPrice:0,fileName:''};
let totalQuantityTouched=false;
const periodVal=id=>Math.max(0,Number(periodEl[id].value)||0);
const withVat=amount=>amount*(1+VAT_RATE);
const setText=(id,text)=>{const node=document.getElementById(id);if(node)node.textContent=text};
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
  const netProductPrice=adReport.productPrice-couponUnit;
  const adjustedAdRevenue=Math.max(0,adReport.revenue-couponUnit*adQuantity);
  const organicRevenue=organicQuantity*netProductPrice;
  const totalRevenue=adjustedAdRevenue+organicRevenue;
  const adSpendVat=withVat(adReport.spend);
  const categoryFee=totalRevenue*categoryRate/100;
  const categoryFeeVat=withVat(categoryFee);
  const productCostTotal=unitCost*totalQuantity;
  const returnRate=adQuantity?returnCount/adQuantity*100:0;
  const returnTotal=returnCount*netProductPrice;
  const shippingTotal=shippingUnit*totalQuantity;
  const shippingTotalVat=withVat(shippingTotal);
  const profit=totalRevenue-adSpendVat-categoryFeeVat-productCostTotal-shippingTotalVat-other;
  const margin=totalRevenue?profit/totalRevenue*100:0;
  const ringValue=Math.max(0,Math.min(100,margin));
  const profitEl=document.getElementById('periodProfit');
  profitEl.textContent=periodMoney(profit);
  profitEl.classList.toggle('loss',profit<0);
  setText('periodOrganicQuantity',`${won.format(organicQuantity)}개`);
  setText('periodOrganicRevenue',periodMoney(organicRevenue));
  setText('periodAdRevenue',periodMoney(adjustedAdRevenue));
  setText('periodMargin',`총 전환매출 대비 순이익률 ${margin.toFixed(1)}%`);
  setText('periodProfitRate',`${margin.toFixed(0)}%`);
  setText('periodSummaryState',!adReport.rows?'보고서 대기':profit>=0?'수익 구간':'손실 구간');
  setText('periodReturnRate',`${returnRate.toFixed(1)}%`);
  setText('periodReturnTotal',periodMoney(returnTotal));
  setText('periodReturnTotalInput',periodMoney(returnTotal));
  const ring=document.getElementById('periodProfitRing');
  if(ring)ring.style.setProperty('--rate',`${ringValue}%`);
  const rows=[
    ['광고 전환매출',periodMoney(adjustedAdRevenue),'plus'],
    ['자연전환매출',periodMoney(organicRevenue),'plus'],
    ['광고비 (VAT 포함)',minus(adSpendVat),'minus'],
    [`카테고리 수수료 (${categoryRate.toFixed(1)}%, VAT 포함)`,minus(categoryFeeVat),'minus'],
    ['최종 매입원가',minus(productCostTotal),'minus'],
    ['로켓그로스 입출고비(윙 배송비, VAT 포함)',minus(shippingTotalVat),'minus'],
    ['기타 비용 합계',minus(other),'minus'],
    ['기간 순이익',periodMoney(profit),profit<0?'loss':'total']
  ];
  document.getElementById('periodBreakdown').innerHTML=rows.map(([label,value,type])=>`<div class="result-cost-row ${type}"><span>${label}</span><strong>${value}</strong></div>`).join('');
  setText('periodFormulaNumbers',`= ${won.format(Math.round(adjustedAdRevenue))} + (${won.format(Math.round(netProductPrice))} × ${won.format(organicQuantity)}) - ${won.format(Math.round(adReport.spend))} × 1.1 - ${won.format(Math.round(categoryFee))} × 1.1 - ${won.format(Math.round(productCostTotal))} - (${won.format(Math.round(shippingUnit))} × ${won.format(totalQuantity)} × 1.1) - ${won.format(Math.round(other))}`);
  setText('periodFormulaResult',`= ${periodMoney(profit)}`);
}
function reportNumber(value){if(typeof value==='number')return value;return Number(String(value??'').replace(/,/g,''))||0}
function sum(rows,key){return rows.reduce((total,row)=>total+reportNumber(row[key]),0)}
function reportText(value){return String(value??'').trim()}
function campaignRange(rows){const starts=[...new Set(rows.map(row=>reportText(row['캠페인 시작일'])).filter(value=>value&&value!=='-'))].sort(),ends=[...new Set(rows.map(row=>reportText(row['캠페인 종료일'])).filter(value=>value&&value!=='-'))].sort();return{start:starts[0]||'',end:ends.at(-1)||''}}
function renderAdReport(){
  setText('periodAdSpend',periodMoney(withVat(adReport.spend)));
  setText('periodSoldQuantity',`${won.format(adReport.quantity)}개`);
  setText('periodRowCount',adReport.rows?`${won.format(adReport.rows)}개 행 분석 완료`:'0개 행');
  setText('periodFileName',adReport.fileName||'선택된 파일 없음');
  setText('periodCampaignRange',adReport.start?`${adReport.start} ~ ${adReport.end||'진행 중'}`:'보고서 업로드 후 자동 표시');
  renderPeriod();
}
async function analyzeAdReport(event){const file=event.target.files[0];if(!file)return;try{if(!window.XLSX)throw new Error('엑셀 분석 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.');setText('periodUploadMessage','보고서를 분석하고 있습니다...');const workbook=XLSX.read(await file.arrayBuffer(),{type:'array'}),sheet=workbook.Sheets[workbook.SheetNames[0]],rows=XLSX.utils.sheet_to_json(sheet,{defval:0});if(!rows.length||!Object.prototype.hasOwnProperty.call(rows[0],'광고비'))throw new Error('쿠팡 광고보고서 형식을 확인할 수 없습니다. 광고비 열이 필요합니다.');const range=campaignRange(rows);adReport={spend:sum(rows,'광고비'),revenue:sum(rows,'총 전환매출액(14일)')||sum(rows,'총 전환매출액(1일)'),quantity:sum(rows,'총 판매수량(14일)')||sum(rows,'총 판매수량(1일)'),rows:rows.length,fileName:file.name,...range,productPrice:0};adReport.productPrice=adReport.quantity?adReport.revenue/adReport.quantity:0;if(!totalQuantityTouched)periodEl.periodTotalQuantity.value=adReport.quantity;setText('periodUploadMessage','보고서 분석이 완료되었습니다.');renderAdReport()}catch(error){adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:'',productPrice:0,fileName:''};setText('periodUploadMessage',error.message);renderAdReport()}}
function resetPeriodInputs(){Object.entries(periodEl).forEach(([id,input])=>input.value=id==='periodShippingCost'?3650:id==='periodCategoryFeeRate'?10.6:0);const category=document.getElementById('periodCategory');if(category)category.value='10.6';totalQuantityTouched=false;renderPeriod()}
document.querySelectorAll('.tab-button').forEach(button=>button.addEventListener('click',()=>{document.querySelectorAll('.tab-button,.tab-panel').forEach(item=>item.classList.remove('active'));button.classList.add('active');document.getElementById(button.dataset.tab).classList.add('active')}));
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
