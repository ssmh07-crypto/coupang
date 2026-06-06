const won=new Intl.NumberFormat('ko-KR',{maximumFractionDigits:0});
const money=n=>`${won.format(Math.round(n))}원`;
const minus=n=>`- ${money(n)}`;
const periodIds=['periodProductUnitCost','periodReturnCount','periodCouponUnitPrice','periodShippingCost','periodOtherCost'];
const periodEl=Object.fromEntries(periodIds.map(id=>[id,document.getElementById(id)]));
let adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:'',productPrice:0,fileName:''};
const periodVal=id=>Math.max(0,Number(periodEl[id].value)||0);
const setText=(id,text)=>{const node=document.getElementById(id);if(node)node.textContent=text};
function periodMoney(amount){return money(amount)}
function renderPeriod(){
  const unitCost=periodVal('periodProductUnitCost');
  const returnCount=periodVal('periodReturnCount');
  const couponUnit=periodVal('periodCouponUnitPrice');
  const shipping=periodVal('periodShippingCost');
  const other=periodVal('periodOtherCost');
  const productCostTotal=unitCost*adReport.quantity;
  const returnRate=(adReport.quantity+returnCount)?returnCount/(adReport.quantity+returnCount)*100:0;
  const returnTotal=returnCount*adReport.productPrice;
  const couponTotal=couponUnit*adReport.quantity;
  const profit=adReport.revenue-adReport.spend-productCostTotal-returnTotal-couponTotal-shipping-other;
  const margin=adReport.revenue?profit/adReport.revenue*100:0;
  const ringValue=Math.max(0,Math.min(100,margin));
  const profitEl=document.getElementById('periodProfit');
  profitEl.textContent=periodMoney(profit);
  profitEl.classList.toggle('loss',profit<0);
  setText('periodMargin',`광고 전환매출 대비 순이익률 ${margin.toFixed(1)}%`);
  setText('periodProfitRate',`${margin.toFixed(0)}%`);
  setText('periodSummaryState',!adReport.rows?'보고서 대기':profit>=0?'수익 구간':'손실 구간');
  setText('periodReturnRate',`${returnRate.toFixed(1)}%`);
  setText('periodReturnTotal',periodMoney(returnTotal));
  setText('periodReturnTotalInput',periodMoney(returnTotal));
  const ring=document.getElementById('periodProfitRing');
  if(ring)ring.style.setProperty('--rate',`${ringValue}%`);
  const rows=[
    ['광고 전환매출',periodMoney(adReport.revenue),'plus'],
    ['광고비 (VAT 포함)',minus(adReport.spend),'minus'],
    ['최종 매입원가',minus(productCostTotal),'minus'],
    ['반품된 상품 총 가격',minus(returnTotal),'minus'],
    ['할인쿠폰 차감액',minus(couponTotal),'minus'],
    ['배송비 합계',minus(shipping),'minus'],
    ['기타 비용 합계',minus(other),'minus'],
    ['기간 순이익',periodMoney(profit),profit<0?'loss':'total']
  ];
  document.getElementById('periodBreakdown').innerHTML=rows.map(([label,value,type])=>`<div class="result-cost-row ${type}"><span>${label}</span><strong>${value}</strong></div>`).join('');
  setText('periodFormulaNumbers',`= ${won.format(Math.round(adReport.revenue))} - ${won.format(Math.round(adReport.spend))} - ${won.format(Math.round(productCostTotal))} - ${won.format(Math.round(returnTotal))} - (${won.format(Math.round(couponUnit))} × ${won.format(adReport.quantity)}) - ${won.format(Math.round(shipping))} - ${won.format(Math.round(other))}`);
  setText('periodFormulaResult',`= ${periodMoney(profit)}`);
}
function reportNumber(value){if(typeof value==='number')return value;return Number(String(value??'').replace(/,/g,''))||0}
function sum(rows,key){return rows.reduce((total,row)=>total+reportNumber(row[key]),0)}
function reportText(value){return String(value??'').trim()}
function campaignRange(rows){const starts=[...new Set(rows.map(row=>reportText(row['캠페인 시작일'])).filter(value=>value&&value!=='-'))].sort(),ends=[...new Set(rows.map(row=>reportText(row['캠페인 종료일'])).filter(value=>value&&value!=='-'))].sort();return{start:starts[0]||'',end:ends.at(-1)||''}}
function renderAdReport(){
  setText('periodAdSpend',periodMoney(adReport.spend));
  setText('periodAdRevenue',periodMoney(adReport.revenue));
  setText('periodSoldQuantity',`${won.format(adReport.quantity)}개`);
  setText('periodRowCount',adReport.rows?`${won.format(adReport.rows)}개 행 분석 완료`:'0개 행');
  setText('periodFileName',adReport.fileName||'선택된 파일 없음');
  setText('periodCampaignRange',adReport.start?`${adReport.start} ~ ${adReport.end||'진행 중'}`:'보고서 업로드 후 자동 표시');
  renderPeriod();
}
async function analyzeAdReport(event){const file=event.target.files[0];if(!file)return;try{if(!window.XLSX)throw new Error('엑셀 분석 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.');setText('periodUploadMessage','보고서를 분석하고 있습니다...');const workbook=XLSX.read(await file.arrayBuffer(),{type:'array'}),sheet=workbook.Sheets[workbook.SheetNames[0]],rows=XLSX.utils.sheet_to_json(sheet,{defval:0});if(!rows.length||!Object.prototype.hasOwnProperty.call(rows[0],'광고비'))throw new Error('쿠팡 광고보고서 형식을 확인할 수 없습니다. 광고비 열이 필요합니다.');const range=campaignRange(rows);adReport={spend:sum(rows,'광고비'),revenue:sum(rows,'총 전환매출액(14일)')||sum(rows,'총 전환매출액(1일)'),quantity:sum(rows,'총 판매수량(14일)')||sum(rows,'총 판매수량(1일)'),rows:rows.length,fileName:file.name,...range,productPrice:0};adReport.productPrice=adReport.quantity?adReport.revenue/adReport.quantity:0;setText('periodUploadMessage','보고서 분석이 완료되었습니다.');renderAdReport()}catch(error){adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:'',productPrice:0,fileName:''};setText('periodUploadMessage',error.message);renderAdReport()}}
function resetPeriodInputs(){Object.values(periodEl).forEach(input=>input.value=0);renderPeriod()}
document.querySelectorAll('.tab-button').forEach(button=>button.addEventListener('click',()=>{document.querySelectorAll('.tab-button,.tab-panel').forEach(item=>item.classList.remove('active'));button.classList.add('active');document.getElementById(button.dataset.tab).classList.add('active')}));
document.getElementById('adReportFile').addEventListener('change',analyzeAdReport);Object.values(periodEl).forEach(input=>input.addEventListener('input',renderPeriod));document.getElementById('periodRecalculate').addEventListener('click',renderPeriod);renderAdReport();
