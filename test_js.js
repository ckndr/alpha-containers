function selectL3NativeMonths() {
  const months = getNativeCustMonths(_nativeSelectedCust);
  const l3 = months.slice(-3);
  _nativeSelectedMonths = new Set(l3);
  renderNativeCustMain();
}

function toggleNativeMonth(m) {
  if (!_nativeSelectedMonths) {
    _nativeSelectedMonths = new Set([m]);
  } else {
    if (_nativeSelectedMonths.has(m)) {
      _nativeSelectedMonths.delete(m);
      if (_nativeSelectedMonths.size === 0) {
        _nativeSelectedMonths = null;
      }
    } else {
      _nativeSelectedMonths.add(m);
    }
  }
  renderNativeCustMain();
}

function renderNativeCustMain() {
  const container = document.getElementById('nativeCustMainContent');
  if (!container || !_nativeSelectedCust) return;
  
  const data = (typeof CUSTOMER_REPORT_DATA !== 'undefined') ? CUSTOMER_REPORT_DATA : [];
  const isAllCust = (_nativeSelectedCust === "ALL CUSTOMERS (Factory Total)");
  const months = getNativeCustMonths(_nativeSelectedCust);
  const l3 = months.slice(-3);
  
  const isAll = (!_nativeSelectedMonths || _nativeSelectedMonths.size === 0);
  const isL3  = (!isAll && l3.length > 0 && l3.every(m => _nativeSelectedMonths.has(m)) && _nativeSelectedMonths.size === l3.length);
  
  const filteredRecs = data.filter(r => {
    if (!isAllCust && r.customer !== _nativeSelectedCust) return false;
    if (_nativeTypeFilter !== 'all' && r.type !== _nativeTypeFilter) return false;
    if (!isAll && !_nativeSelectedMonths.has(r.month)) return false;
    return true;
  });
  
  const tubeProdTotal = filteredRecs.filter(r => r.type === 'TUBE').reduce((s,r) => s + (r.produced || r.good || 0), 0);
  const tubeDispTotal = filteredRecs.filter(r => r.type === 'TUBE').reduce((s,r) => s + (r.dispatched || 0), 0);
  const tubeRejectTotal = filteredRecs.filter(r => r.type === 'TUBE').reduce((s,r) => s + (r.reject || 0), 0);
  
  const petProdTotal  = filteredRecs.filter(r => r.type === 'PET').reduce((s,r) => s + (r.produced || r.good || 0), 0);
  const petDispTotal  = filteredRecs.filter(r => r.type === 'PET').reduce((s,r) => s + (r.dispatched || 0), 0);
  const petRejectTotal = filteredRecs.filter(r => r.type === 'PET').reduce((s,r) => s + (r.reject || 0), 0);
  
  const grandProdTotal = tubeProdTotal + petProdTotal;
  const grandDispTotal = tubeDispTotal + petDispTotal;
  
  const tubeRejPct = tubeProdTotal > 0 ? ((tubeRejectTotal / tubeProdTotal) * 100).toFixed(1) : '0.0';
  const petRejPct  = petProdTotal > 0 ? ((petRejectTotal / petProdTotal) * 100).toFixed(1) : '0.0';
  
  let selectionSubtext = `All ${months.length} months`;
  if (isL3) {
    selectionSubtext = "Last 3 months (L3)";
  } else if (!isAll) {
    selectionSubtext = `Custom (${_nativeSelectedMonths.size} months selected)`;
  }

  let periodBtns = `<button class="filter-btn ${isAll ? 'active' : ''}" onclick="resetNativeMonths()">All Months</button>`;
  periodBtns += `<button class="filter-btn ${isL3 ? 'active' : ''}" onclick="selectL3NativeMonths()">Last 3 Months</button>`;
  
  months.forEach(m => {
    const active = _nativeSelectedMonths && _nativeSelectedMonths.has(m);
    periodBtns += `<button class="filter-btn ${active ? 'active' : ''}" onclick="toggleNativeMonth('${m}')">${m}</button>`;
  });
  
  const typeBtns = ['all','TUBE','PET'].map(t =>
    `<button class="filter-btn ${_nativeTypeFilter===t?'active':''}" onclick="setNativeTypeFilter('${t}')">${t==='all'?'All Types':t}</button>`
  ).join('');

  const monthSummary = {};
  data.filter(r => isAllCust || r.customer === _nativeSelectedCust).forEach(r => {
    if (!monthSummary[r.month]) monthSummary[r.month] = {TUBE_P:0, TUBE_D:0, PET_P:0, PET_D:0, TUBE_R:0, PET_R:0};
    if (r.type === 'TUBE') {
      monthSummary[r.month].TUBE_P += (r.produced || r.good || 0);
      monthSummary[r.month].TUBE_D += (r.dispatched || 0);
      monthSummary[r.month].TUBE_R += (r.reject || 0);
    } else {
      monthSummary[r.month].PET_P  += (r.produced || r.good || 0);
      monthSummary[r.month].PET_D  += (r.dispatched || 0);
      monthSummary[r.month].PET_R  += (r.reject || 0);
    }
  });
  
  const monthRows = months.map(m => {
    const d = monthSummary[m] || {TUBE_P:0, TUBE_D:0, PET_P:0, PET_D:0, TUBE_R:0, PET_R:0};
    const tubeRej = d.TUBE_P > 0 ? ((d.TUBE_R / d.TUBE_P) * 100).toFixed(1) : '-';
    const petRej  = d.PET_P > 0 ? ((d.PET_R / d.PET_P) * 100).toFixed(1) : (d.PET_D > 0 ? '*' : '-');
    const isSelected = isAll || (_nativeSelectedMonths && _nativeSelectedMonths.has(m));
    const bgStyle = isSelected ? (isAll ? '' : 'background:rgba(35,85,160,0.08);font-weight:600') : 'opacity:0.4';
    return `<tr style="${bgStyle}">
      <td style="font-weight:600">${m} ${isSelected && !isAll ? '✓' : ''}</td>
      <td class="num" style="color:var(--blue-light)">${d.TUBE_P.toLocaleString()}</td>
      <td class="num" style="color:var(--blue-light);font-size:11px">${d.TUBE_D.toLocaleString()}</td>
      <td class="num" style="color:var(--red);font-size:11px">${tubeRej === '-' ? '-' : tubeRej + '%'}</td>
      <td class="num" style="color:var(--green)">${d.PET_P.toLocaleString()}</td>
      <td class="num" style="color:var(--green);font-size:11px">${d.PET_D.toLocaleString()}</td>
      <td class="num" style="color:var(--red);font-size:11px">${petRej === '-' ? '-' : (petRej === '*' ? '<span title="PET production data not available for this month">*</span>' : petRej + '%')}</td>
    </tr>`;
  }).join('');
  
  const allM = Object.values(monthSummary);
  const sumTUBEP   = allM.reduce((s,d)=>s+d.TUBE_P,0);
  const sumTUBED   = allM.reduce((s,d)=>s+d.TUBE_D,0);
  const sumPETP    = allM.reduce((s,d)=>s+d.PET_P,0);
  const sumPETD    = allM.reduce((s,d)=>s+d.PET_D,0);
  const sumTUBER   = allM.reduce((s,d)=>s+d.TUBE_R,0);
  const sumPETR    = allM.reduce((s,d)=>s+d.PET_R,0);
  const sumTubeRejPct = sumTUBEP > 0 ? ((sumTUBER/sumTUBEP)*100).toFixed(1) : '0.0';
  const sumPetRejPct  = sumPETP > 0 ? ((sumPETR/sumPETP)*100).toFixed(1) : '0.0';

  const products = {};
  filteredRecs.forEach(r => {
    const k = r.product + '||' + r.type + (isAllCust ? '||' + r.customer : '');
    if (!products[k]) products[k] = {product:r.product, type:r.type, customer:r.customer, produced:0, dispatched:0};
    products[k].produced   += (r.produced || r.good || 0);
    products[k].dispatched += (r.dispatched || 0);
  });
  const prodMax = Math.max(...Object.values(products).map(p=>p.produced), 1);
  const prodRows = Object.values(products)
    .sort((a,b)=>b.produced-a.produced)
    .map(p => {
      const pct = ((p.produced/prodMax)*100).toFixed(0);
      const custTag = isAllCust ? `<div style="font-size:10px;color:var(--muted)">${p.customer}</div>` : '';
      return `<tr>
        <td><span class="type-badge ${p.type}">${p.type}</span></td>
        <td style="font-weight:500">${p.product}${custTag}</td>
        <td class="num" style="font-weight:600;color:var(--navy)">${p.produced.toLocaleString()}</td>
        <td class="num" style="color:var(--accent)">${p.dispatched.toLocaleString()}</td>
        <td style="width:120px">
          <div style="display:flex;align-items:center;gap:6px">
            <div style="flex:1;height:6px;background:var(--light);border-radius:3px;overflow:hidden">
              <div style="height:100%;width:${pct}%;background:${p.type==='TUBE'?'var(--blue)':'var(--green)'};border-radius:3px"></div>
            </div>
            <span style="font-size:11px;color:var(--muted);width:28px;text-align:right">${pct}%</span>
          </div>
        </td>
      </tr>`;
    }).join('');

  const displayCustTitle = isAllCust ? '🌐 ALL CUSTOMERS (Factory Total)' : _nativeSelectedCust;

  container.innerHTML = `
    <div class="cust-hdr-bar">
      <div>
        <div class="cust-title-text">${displayCustTitle}</div>
        <div style="font-size:11px;color:var(--muted);margin-top:2px">Historical Production & Dispatch Log (Records since November 2025)</div>
      </div>
      <span class="cust-badge-pill">${selectionSubtext}</span>
    </div>

    <div class="cust-kpi-grid">
      <div class="cust-kpi">
        <div class="lbl">TUBE Produced</div>
        <div class="val" style="color:var(--blue)">${tubeProdTotal.toLocaleString()}</div>
      </div>
      <div class="cust-kpi">
        <div class="lbl">TUBE Dispatched</div>
        <div class="val" style="color:var(--blue)">${tubeDispTotal.toLocaleString()}</div>
      </div>
      <div class="cust-kpi reject-card">
        <div class="lbl">TUBE Reject</div>
        <div class="val" style="color:var(--red)">${tubeRejectTotal.toLocaleString()}</div>
        <div class="sub">${tubeRejPct}% of tube production</div>
      </div>
      <div class="cust-kpi">
        <div class="lbl">PET Produced</div>
        <div class="val" style="color:var(--green)">${petProdTotal.toLocaleString()}</div>
      </div>
      <div class="cust-kpi">
        <div class="lbl">PET Dispatched</div>
        <div class="val" style="color:var(--green)">${petDispTotal.toLocaleString()}</div>
      </div>
      <div class="cust-kpi reject-card">
        <div class="lbl">PET Reject</div>
        <div class="val" style="color:var(--red)">${petRejectTotal.toLocaleString()}</div>
        <div class="sub">${petRejPct}% of PET production</div>
      </div>
    </div>

    <div class="cust-filter-bar">
      <span style="font-size:11px;font-weight:600;color:var(--muted)">PERIOD (MULTI-SELECT):</span>
      ${periodBtns}
      <div style="margin-left:auto;display:flex;align-items:center;gap:6px">
        <span style="font-size:11px;font-weight:600;color:var(--muted)">TYPE:</span>
        ${typeBtns}
      </div>
    </div>

    <div class="cust-two-col-grid">
      <div class="cust-tbl-card">
        <h3>Monthly Production & Dispatch Breakdown</h3>
        <table class="cust-tbl">
          <thead>
            <tr>
              <th>Month</th>
              <th class="num">TUBE Prod</th>
              <th class="num">TUBE Disp</th>
              <th class="num" style="color:var(--red)">Tube Rej%</th>
              <th class="num">PET Prod</th>
              <th class="num">PET Disp</th>
              <th class="num" style="color:var(--red)">PET Rej%</th>
            </tr>
          </thead>
          <tbody>
            ${monthRows}
            <tr class="tot">
              <td>ALL MONTHS TOTAL</td>
              <td class="num">${sumTUBEP.toLocaleString()}</td>
              <td class="num">${sumTUBED.toLocaleString()}</td>
              <td class="num" style="color:var(--red)">${sumTubeRejPct}%</td>
              <td class="num">${sumPETP.toLocaleString()}</td>
              <td class="num">${sumPETD.toLocaleString()}</td>
              <td class="num" style="color:var(--red)">${sumPetRejPct}%</td>
            </tr>
          </tbody>
        </table>
        <div class="cust-footnote" style="margin-top:12px">
          <span class="fn-icon">ℹ️</span>
          <div><strong>Note:</strong> PET bottle production data is available from <strong>January 2026</strong> onward. PET production for November & December 2025 is not tracked as ERP does not record PET production and manual tracking began in January 2026. PET dispatch data from ERP is available for all months. Where PET Rej% shows <strong>*</strong>, production data is unavailable for that month.</div>
        </div>
      </div>

      <div class="cust-tbl-card">
        <h3>Product SKU Breakdown (${selectionSubtext})</h3>
        <table class="cust-tbl">
          <thead><tr><th>Type</th><th>Product</th><th class="num">Produced</th><th class="num">Dispatched</th><th>Share</th></tr></thead>
          <tbody>${prodRows || '<tr><td colspan="5" class="no-data">No records match filter</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `;
}

function setNativeMonthFilter(k) {
