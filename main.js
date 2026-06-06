const MONTHLY_FEE=55000;
const ids=['feeRate','salePrice','sellerCoupon','productCost','quantity','adCost','otherCost','monthlyQuantity','wingShipping','wingShippingRevenue','wingPacking','rocketHandling','rocketShipping','rocketInbound','rocketOther'];
const el=Object.fromEntries(ids.map(id=>[id,document.getElementById(id)]));
const category=document.getElementById('category'),monthly=document.getElementById('applyMonthlyFee');
const won=new Intl.NumberFormat('ko-KR',{maximumFractionDigits:0});
const val=id=>Math.max(0,Number(el[id].value)||0),money=n=>`${won.format(Math.round(n))}원`,minus=n=>`- ${money(n)}`;
const coupangFee=(amount,rate,qty=1)=>{const beforeVat=Math.round(amount/qty*rate/100)*qty;return beforeVat+Math.round(beforeVat*.1)};
function breakdown(target,r){target.innerHTML=`<div><dt>총 결제금액</dt><dd>${money(r.revenue)}</dd></div><div><dt>판매자 할인쿠폰</dt><dd>${minus(r.coupon)}</dd></div><div><dt>판매 수수료(VAT 포함)</dt><dd>${minus(r.fee)}</dd></div><div><dt>상품 원가</dt><dd>${minus(r.cost)}</dd></div><div><dt>${r.label}</dt><dd>${minus(r.delivery)}</dd></div><div><dt>공통·기타 비용</dt><dd>${minus(r.other)}</dd></div>`}
function render(name,r){const profit=r.revenue-r.coupon-r.fee-r.cost-r.delivery-r.other,margin=r.revenue?profit/r.revenue*100:0;document.getElementById(`${name}Profit`).textContent=money(profit);document.getElementById(`${name}Margin`).textContent=`${margin.toFixed(1)}%`;breakdown(document.getElementById(`${name}Breakdown`),r);document.getElementById(`${name}Card`).classList.toggle('negative',profit<0);return profit}
function calculate(){const qty=Math.max(1,val('quantity')),revenue=val('salePrice')*qty,coupon=Math.min(revenue,val('sellerCoupon')*qty),fee=coupangFee(revenue-coupon,val('feeRate'),qty),cost=val('productCost')*qty,other=(val('adCost')+val('otherCost'))*qty+(monthly.checked?MONTHLY_FEE/Math.max(1,val('monthlyQuantity'))*qty:0);const wingShippingRevenue=val('wingShippingRevenue')*qty,wingShippingFee=coupangFee(wingShippingRevenue,3,qty);const wing=render('wing',{revenue:revenue+wingShippingRevenue,coupon,fee:fee+wingShippingFee,cost,other,label:'배송·포장비',delivery:(val('wingShipping')+val('wingPacking'))*qty});const rocket=render('rocket',{revenue,coupon,fee,cost,other,label:'풀필먼트 비용',delivery:(val('rocketHandling')+val('rocketShipping')+val('rocketInbound')+val('rocketOther'))*qty});const diff=Math.abs(wing-rocket),wingCard=document.getElementById('wingCard'),rocketCard=document.getElementById('rocketCard');wingCard.classList.remove('best');rocketCard.classList.remove('best');document.getElementById('difference').textContent=money(diff);const winner=document.getElementById('winner');if(!val('salePrice'))winner.textContent='판매가를 입력해 주세요.';else if(wing===rocket)winner.textContent='두 판매 방식의 예상 이익이 같습니다.';else{const wingWins=wing>rocket;(wingWins?wingCard:rocketCard).classList.add('best');winner.textContent=`${wingWins?'판매자배송':'로켓그로스'}이 ${money(diff)} 더 유리합니다.`}}
function fillExample(){Object.entries({feeRate:10.6,salePrice:18900,sellerCoupon:0,productCost:7200,quantity:100,adCost:900,otherCost:150,monthlyQuantity:500,wingShipping:3000,wingShippingRevenue:3000,wingPacking:300,rocketHandling:650,rocketShipping:1550,rocketInbound:180,rocketOther:100}).forEach(([id,v])=>el[id].value=v);category.value='10.6';monthly.checked=true;calculate()}
function reset(){Object.values(el).forEach(input=>input.value=input.id==='quantity'?1:input.id==='monthlyQuantity'?100:0);el.rocketHandling.value=650;el.rocketShipping.value=1550;category.value='10.6';el.feeRate.value=10.6;monthly.checked=false;calculate()}
category.addEventListener('change',()=>{if(category.value!=='custom')el.feeRate.value=category.value;calculate()});el.feeRate.addEventListener('input',()=>{if(category.value!=='custom'&&Number(category.value)!==Number(el.feeRate.value))category.value='custom'});document.querySelectorAll('input').forEach(input=>input.addEventListener('input',calculate));document.getElementById('fillExample').addEventListener('click',fillExample);document.getElementById('reset').addEventListener('click',reset);calculate();

const periodIds=['periodProductUnitCost','periodShippingCost','periodOtherCost'];
const periodEl=Object.fromEntries(periodIds.map(id=>[id,document.getElementById(id)]));
let adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:''};
const periodVal=id=>Math.max(0,Number(periodEl[id].value)||0);
const setText=(id,text)=>{const node=document.getElementById(id);if(node)node.textContent=text};
function periodMoney(amount){return money(amount)}
function renderPeriod(){
  const product=periodVal('periodProductUnitCost')*adReport.quantity;
  const shipping=periodVal('periodShippingCost');
  const other=periodVal('periodOtherCost');
  const totalCost=adReport.spend+product+shipping+other;
  const profit=adReport.revenue-totalCost;
  const margin=adReport.revenue?profit/adReport.revenue*100:0;
  const roas=adReport.spend?adReport.revenue/adReport.spend*100:0;
  const costPerSale=adReport.quantity?adReport.spend/adReport.quantity:0;
  const profitEl=document.getElementById('periodProfit');
  profitEl.textContent=periodMoney(profit);
  profitEl.classList.toggle('loss',profit<0);
  setText('periodMargin',`광고 전환매출 대비 ${margin.toFixed(1)}%`);
  setText('periodRoas',`${roas.toFixed(1)}%`);
  setText('periodProductCostTotal',`총 매입원가 ${periodMoney(product)}`);
  setText('periodCostTotal',`총 비용 ${periodMoney(totalCost)}`);
  setText('periodCostPerSale',periodMoney(costPerSale));
  setText('periodNetProfitInline',periodMoney(profit));
  setText('periodSummaryState',!adReport.rows?'보고서 대기':profit>=0?'수익 구간':'손실 구간');
  const rows=[
    ['광고 전환매출',periodMoney(adReport.revenue),'plus'],
    ['광고비 (VAT 포함)',minus(adReport.spend),'minus'],
    ['최종 매입원가',minus(product),'minus'],
    ['배송비 합계',minus(shipping),'minus'],
    ['기타 비용',minus(other),'minus'],
    ['예상 순이익',periodMoney(profit),profit<0?'loss':'total']
  ];
  document.getElementById('periodBreakdown').innerHTML=rows.map(([label,value,type])=>`<div class="ledger-row ${type}"><span>${label}</span><strong>${value}</strong></div>`).join('');
}
function reportNumber(value){if(typeof value==='number')return value;return Number(String(value??'').replace(/,/g,''))||0}
function sum(rows,key){return rows.reduce((total,row)=>total+reportNumber(row[key]),0)}
function reportText(value){return String(value??'').trim()}
function campaignRange(rows){const starts=[...new Set(rows.map(row=>reportText(row['캠페인 시작일'])).filter(value=>value&&value!=='-'))].sort(),ends=[...new Set(rows.map(row=>reportText(row['캠페인 종료일'])).filter(value=>value&&value!=='-'))].sort();return{start:starts[0]||'',end:ends.at(-1)||''}}
function renderAdReport(){
  setText('periodAdSpend',periodMoney(adReport.spend));
  setText('periodAdRevenue',periodMoney(adReport.revenue));
  setText('periodSoldQuantity',`${won.format(adReport.quantity)}개`);
  setText('periodRowCount',`${won.format(adReport.rows)}개 행`);
  setText('periodCampaignRange',adReport.start?`${adReport.start} ~ ${adReport.end||'진행 중'}`:'보고서 업로드 후 자동 표시');
  renderPeriod();
}
async function analyzeAdReport(event){const file=event.target.files[0],status=document.getElementById('adReportStatus');if(!file)return;try{if(!window.XLSX)throw new Error('엑셀 분석 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.');status.textContent='보고서를 분석하고 있습니다...';status.classList.remove('complete');const workbook=XLSX.read(await file.arrayBuffer(),{type:'array'}),sheet=workbook.Sheets[workbook.SheetNames[0]],rows=XLSX.utils.sheet_to_json(sheet,{defval:0});if(!rows.length||!Object.prototype.hasOwnProperty.call(rows[0],'광고비'))throw new Error('쿠팡 광고보고서 형식을 확인할 수 없습니다. 광고비 열이 필요합니다.');const range=campaignRange(rows);adReport={spend:sum(rows,'광고비'),revenue:sum(rows,'총 전환매출액(14일)')||sum(rows,'총 전환매출액(1일)'),quantity:sum(rows,'총 판매수량(14일)')||sum(rows,'총 판매수량(1일)'),rows:rows.length,...range};status.textContent=`${file.name} · ${won.format(adReport.rows)}개 행 분석 완료`;status.classList.add('complete');renderAdReport()}catch(error){adReport={spend:0,revenue:0,quantity:0,rows:0,start:'',end:''};status.textContent=error.message;status.classList.remove('complete');renderAdReport()}}
document.querySelectorAll('.tab-button').forEach(button=>button.addEventListener('click',()=>{document.querySelectorAll('.tab-button,.tab-panel').forEach(item=>item.classList.remove('active'));button.classList.add('active');document.getElementById(button.dataset.tab).classList.add('active')}));
document.getElementById('adReportFile').addEventListener('change',analyzeAdReport);Object.values(periodEl).forEach(input=>input.addEventListener('input',renderPeriod));renderAdReport();
