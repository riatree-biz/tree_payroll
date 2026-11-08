document.addEventListener('DOMContentLoaded', () => {
  
  const defaults = { mealRate: 0
                   , vehicleRate: 0
                   , mealLimit: 200000
                   , vehicleLimit: 200000
                   , npsRate: 4.75
                   , nhisRate: 3.595
                   , ltciRate: 13.14
                   , eiRate: .9
                   , incomeTaxRate: 3.3
                   , localIncomeTaxRate: 10 
                  };

  const key = 'treePayrollSettingsV1'; 
  
  let settings = { ...defaults }; 
  
  try { 
    settings = { ...settings, ...JSON.parse(localStorage.getItem(key) || '{}') } 
  } catch { 

  }
  
  const engine = new PayrollEngine(settings); 
  let workbook = null, workers = [], calculated = false;
  const $ = id => document.getElementById(id), tbody = $('payroll-tbody'), preview = $('excel-preview-container'), sheet = $('excel-sheet-select');
  const num = v => { const n = Number(String(v ?? 0).replace(/[,₩원\s]/g, '')); return Number.isFinite(n) ? n : 0 }, fmt = v => Math.round(num(v)).toLocaleString('ko-KR');
  
  document.querySelectorAll('.tab-btn').forEach(b => b.addEventListener('click', ( ) => { 
    document.querySelectorAll('.tab-btn').forEach(x => { x.classList.toggle('active', x === b); 
      x.setAttribute('aria-selected', String(x === b)) }); 
      document.querySelectorAll('.tab-content').forEach(x => x.classList.toggle('active', x.id === b.dataset.target)) }));
  
  
  function sync() { 
    document.querySelectorAll('[data-setting]').forEach(i => i.value = settings[i.dataset.setting]) 
  } sync();
  
  $('btn-save-settings').addEventListener('click', () => { const next = { ...settings }; for (const i of document.querySelectorAll('[data-setting]')) { const v = Number(i.value); if (!Number.isFinite(v) || v < 0 || v > 100 && /Rate$/.test(i.dataset.setting)) return alert('설정값은 0 이상이어야 하며 요율은 100% 이하여야 합니다.'); next[i.dataset.setting] = v } settings = next; localStorage.setItem(key, JSON.stringify(settings)); engine.updateRules(settings); if (workers.length) calculate(false); alert(workers.length ? '계산 설정을 저장하고 전체 직원의 계산값을 갱신했습니다.' : '계산 설정을 저장했습니다. 직원 데이터를 불러온 뒤 계산하면 적용됩니다.') });
  $('btn-reset-settings').addEventListener('click', () => { if (!confirm('기본 설정으로 복원할까요?')) return; settings = { ...defaults }; localStorage.setItem(key, JSON.stringify(settings)); engine.updateRules(settings); sync(); if (workers.length) calculate(false) });
  $('excel-file-input').addEventListener('change', async e => { const file = e.target.files?.[0]; if (!file) return; try { const r = await ExcelHandler.parseExcelFile(file); workbook = r.workbook; sheet.replaceChildren(); r.sheetNames.forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; sheet.append(o) }); $('excel-sheet-selector').hidden = !r.sheetNames.length; preview.hidden = false; preview.textContent = file.name + ' · ' + r.sheetNames.length + '개 시트를 읽었습니다.' } catch (err) { alert('엑셀 읽기 실패: ' + err.message) } });
  $('btn-parse-sheet').addEventListener('click', () => { if (!workbook || !sheet.value) return alert('엑셀 파일과 시트를 선택해주세요.'); try { workers = ExcelHandler.getSheetData(workbook, sheet.value).map(w => ({ ...w })); calculated = false; render(); preview.textContent = sheet.value + ' · ' + workers.length + '명 불러오기 완료. 금액 셀을 수정할 수 있습니다.'; document.querySelector('[data-target="tab-payroll"]').click() } catch (e) { alert('데이터 분석 실패: ' + e.message) } });
  
  const fields = [['basePay', '기본급'], ['mealAllowance', '식대(원본)'], ['vehicleAllowance', '차량운전보조금(원본)'], ['overtimePay', '연장근로수당'], ['nightPay', '야간근로수당'], ['holidayPay', '휴일근로수당'], ['loadingAllowance', '상하차수당'], ['annualLeavePay', '연차수당'], ['incentive', '인센티브']];
  function add(tr, v, cls = '') { const td = document.createElement('td'); td.textContent = v ?? ''; if (cls) td.className = cls; tr.append(td); return td }
  
  
  function render() {
    tbody.replaceChildren(); workers.forEach((w, idx) => {

      const tr = document.createElement('tr');

      [['name', '성명'], ['dept', '부서']].forEach(([f, label]) => { const td = add(tr, '', 'editable-cell'), 
        input = document.createElement('input'); 
        input.type = 'text'; 
        input.value = w[f] || ''; input.setAttribute('aria-label', (w.name || '신규 직원') + ' ' + label); 
        input.addEventListener('change', () => { workers[idx][f] = input.value.trim(); }); 
        td.append(input) });

      fields.forEach(([f, label]) => { const td = add(tr, '', 'number editable-cell'), 
        input = document.createElement('input'); 
        input.type = 'number'; 
        input.min = '0'; 
        input.step = '1'; 
        input.value = num(w[f]); 
        input.setAttribute('aria-label', (w.name || '직원') + ' ' + label); 
        input.addEventListener('change', () => { const v = Number(input.value); if (!Number.isFinite(v) || v < 0) { alert('0 이상의 숫자를 입력해주세요.'); input.value = num(w[f]); return } workers[idx][f] = v; calculated = false; render() }); td.append(input) });
      add(tr, calculated ? fmt(w.calculatedMealAllowance) : '—', 'number'); add(tr, calculated ? fmt(w.calculatedVehicleAllowance) : '—', 'number'); const pay = calculated ? w.totalPayment : fields.reduce((s, [f]) => s + num(w[f]), 0); add(tr, fmt(pay), 'number');
      ['nps', 'nhis', 'ltci', 'ei', 'incomeTax', 'localIncomeTax'].forEach(f => add(tr, calculated ? fmt(w[f]) : '0', 'number')); add(tr, calculated ? fmt(w.totalDeduction) : '0', 'number'); add(tr, calculated ? fmt(w.netPayment) : '0', 'number'); tbody.append(tr)
    })
  }
  
  
  function calculate(show = true) { 
    if (!workers.length) { 
      if (show) alert('먼저 엑셀 데이터를 불러와주세요.'); return 
    } 
    
    workers = workers.map(w => ({ ...w, ...engine.calculateWorkerSalary(w) })); 
    
    calculated = true; 
    
    render(); if (show) alert(workers.length + '명의 급여 계산이 완료되었습니다.') 
  } // calculate() end

  /*급여행추가*/
  $('btn-add-row').addEventListener('click', () => { 
    workers.push({ name: '', dept: '', employeeNo: '', basePay: 0, mealAllowance: 0, vehicleAllowance: 0, overtimePay: 0, loadingAllowance: 0, incentive: 0 }); 
    calculated = false; 
    render(); const inputs = tbody.querySelectorAll('tr:last-child input'); 
    inputs[0]?.focus(); 
  });


  $('btn-calculate-all').addEventListener('click', () => calculate());   
  $('footer-btn-calc').addEventListener('click', () => calculate());


  $('footer-btn-export').addEventListener('click', () => { 
    
    if (!workers.length) return alert('내보낼 데이터가 없습니다.'); 
    
    const headers = ['성명', '부서', ...fields.map(x => x[1]), '계산 식대', '계산 차량운전보조금', '지급합계', '국민연금', '건강보험', '장기요양', '고용보험', '소득세', '지방소득세', '공제합계', '실지급액']; 
    const rows = workers.map(w => [w.name, w.dept, ...fields.map(([f]) => w[f]), w.calculatedMealAllowance ?? '', w.calculatedVehicleAllowance ?? '', w.totalPayment ?? '', w.nps ?? '', w.nhis ?? '', w.ltci ?? '', w.ei ?? '', w.incomeTax ?? '', w.localIncomeTax ?? '', w.totalDeduction ?? '', w.netPayment ?? '']); 
    const csv = '\uFEFF' + [headers, ...rows].map(r => r.map(v => '"' + String(v ?? '').replace(/"/g, '""') + '"').join(',')).join('\r\n'); 
    
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' })), a = document.createElement('a'); a.href = url; a.download = '급여대장.csv'; a.click(); URL.revokeObjectURL(url) 
  });
  
  
  
  $('btn-save-html').addEventListener('click', async () => {
    try {
      const clone = document.documentElement.cloneNode(true);
      
      document.querySelectorAll('input,select,textarea').forEach((source, index) => {
        const copy = clone.querySelectorAll('input,select,textarea')[index]; if (!copy) return;
        if (source.tagName === 'SELECT') { Array.from(copy.options).forEach((o, i) => o.selected = source.options[i]?.selected ?? false) }
        else if (source.type === 'checkbox' || source.type === 'radio') { if (source.checked) copy.setAttribute('checked', ''); else copy.removeAttribute('checked') }
        else copy.setAttribute('value', source.value);
        if (source.tagName === 'TEXTAREA') copy.textContent = source.value;
      });
      
      const link = document.querySelector('link[rel="stylesheet"]');
      if (link) {
        try { const res = await fetch(link.href); if (res.ok) { const style = document.createElement('style'); style.textContent = await res.text(); clone.querySelector('link[rel="stylesheet"]')?.replaceWith(style) } }
        catch (_) {
          try {
            const sheet = Array.from(document.styleSheets).find(s => s.href === link.href);
            const rules = sheet ? Array.from(sheet.cssRules).map(rule => rule.cssText).join('\\n') : '';
            if (rules) { const style = document.createElement('style'); style.textContent = rules; clone.querySelector('link[rel="stylesheet"]')?.replaceWith(style) }
          } catch (__) {/* Preserve the link as a last-resort fallback; saving must continue. */ }
        }
      }
      // Keep the snapshot self-contained as a view; no scripts are needed to display its captured table.
      clone.querySelectorAll('script').forEach(s => s.remove());
      const blob = new Blob(['<!DOCTYPE html>\\n' + clone.outerHTML], { type: 'text/html;charset=utf-8' }), url = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = url; a.download = 'tree-payroll-screen.html'; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1500);
    } catch (e) { alert('HTML 저장 실패: ' + (e?.message || String(e))) }
  });
  render();
});