/* 임금명세서(근로기준법 제48조 제2항) — 정산 결과로 사람별 PDF를 만들고, 카카오톡·문자로 공유하거나 저장한다.
   일급은 '주휴수당 포함' 계약 기준: 통상시급 = 일급 × 5 ÷ 48 (주 40시간 + 주휴 8시간), 기본일급 = 통상시급 × 8. */
(function () {
  const LIBS = [
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
  ];
  let loading = null, last = null;
  const n = v => Number(v || 0).toLocaleString('ko-KR');
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function loadLibs() {
    if (window.html2canvas && window.jspdf) return Promise.resolve();
    if (loading) return loading;
    loading = Promise.all(LIBS.map(src => new Promise((ok, fail) => {
      const s = document.createElement('script'); s.src = src; s.onload = ok;
      s.onerror = () => { loading = null; fail(Error('PDF 도구를 불러오지 못했습니다. 인터넷 연결을 확인하세요.')); };
      document.head.append(s);
    })));
    return loading;
  }

  /** 주휴 포함 일급을 기본일급·주휴수당으로 나눈다 (원 단위 반올림, 합계는 일급과 같게) */
  function split(wage) {
    const hourly = wage * 5 / 48, base = Math.round(hourly * 8);
    return { hourly: Math.round(hourly), base, weekly: wage - base };
  }

  function data(p, month, payday) {
    const s = split(p.wage), d = p.dayCount, ded = p.incomeTax + p.localTax + p.employment;
    const b = (window.LK_CONFIG && window.LK_CONFIG.business) || {};
    return { p, month, payday, s, d, ded, biz: b,
      baseTotal: s.base * d, weeklyTotal: s.weekly * d,
      empNo: String(p.workerId || p.visitorId || '').replace(/[^0-9a-zA-Z]/g, '').slice(0, 8).toUpperCase(),
      issued: new Date().toLocaleDateString('ko-KR', { timeZone: 'Asia/Seoul' }),
      mw: window.minWageOf ? window.minWageOf(Number(month.slice(0, 4))) : null };
  }

  function html(x) {
    const [y, m] = x.month.split('-'), p = x.p, dayList = p.days.map(d => Number(d.day.slice(8)) + '일').join(', ');
    const bizLine = [x.biz.name || '사업장', x.biz.ceo && '대표 ' + x.biz.ceo, x.biz.bizNo && '사업자등록번호 ' + x.biz.bizNo].filter(Boolean).map(esc).join(' · ');
    const tr = (a, b, c = '') => `<tr><th>${a}</th><td class="r">${b}</td><td class="c">${c}</td></tr>`;
    return `<div class="ps-doc">
<h1>임금명세서</h1><p class="ps-sub">${esc(y)}년 ${Number(m)}월분 · 지급일 ${esc(x.payday)}</p>
<table class="ps-info"><tr><th>사업장</th><td colspan="3">${bizLine}</td></tr>
<tr><th>성명</th><td>${esc(p.name)}</td><th>사원번호</th><td>${esc(x.empNo)}</td></tr>
<tr><th>숙련 단계</th><td>${esc(p.grade || '-')}</td><th>일급</th><td>${n(p.wage)}원 (주휴수당 포함)</td></tr></table>
<h2>지급 내역</h2><table class="ps-t"><tr><th>항목</th><th class="r">금액(원)</th><th class="c">계산 방법</th></tr>
${tr('기본급', n(x.baseTotal), `기본일급 ${n(x.s.base)}원 × ${x.d}일`)}
${tr('주휴수당', n(x.weeklyTotal), `일급 포함분 ${n(x.s.weekly)}원 × ${x.d}일`)}
${tr('연장·야간·휴일수당', '0', '해당 없음')}
<tr class="tot"><th>지급 총액</th><td class="r">${n(p.gross)}</td><td class="c">일급 ${n(p.wage)}원 × ${x.d}일</td></tr></table>
<h2>공제 내역</h2><table class="ps-t"><tr><th>항목</th><th class="r">금액(원)</th><th class="c">계산 방법</th></tr>
${tr('소득세', n(p.incomeTax), '(일급 − 15만원) × 6% × 45% × 근무일수 · 1,000원 미만 소액부징수')}
${tr('지방소득세', n(p.localTax), '소득세의 10%')}
${tr('고용보험', n(p.employment), '지급 총액 × 0.9%')}
<tr class="tot"><th>공제 합계</th><td class="r">${n(x.ded)}</td><td class="c"></td></tr></table>
<table class="ps-t net"><tr><th>실지급액</th><td class="r">${n(p.net)}원</td><td class="c">지급 총액 − 공제 합계</td></tr></table>
<h2>근무 내역</h2><p class="ps-note">근무일수 <b>${x.d}일</b> · 근로시간 합계 <b>${p.hours}시간</b><br>근무일: ${esc(dayList)}<br>
통상시급 ${n(x.s.hourly)}원 = 일급 ${n(p.wage)}원 × 5 ÷ 48 (1일 8시간 · 주 40시간 + 주휴 8시간 기준)<br>최저임금 확인: ${x.mw ? `통상시급 ${n(x.s.hourly)}원 ${p.wage * 5 / 48 >= x.mw ? '≥' : '<'} ${esc(x.month.slice(0, 4))}년 최저시급 ${n(x.mw)}원 ${p.wage * 5 / 48 >= x.mw ? '✓ 충족' : '✗ 미달'}` : '-'}</p>
<p class="ps-foot">이 명세서는 근로기준법 제48조 제2항에 따라 교부합니다. 발급일 ${esc(x.issued)} · 문의: ${esc(x.biz.phone || x.biz.name || '담당자')}</p></div>`;
  }

  const CSS = `.ps-doc{width:794px;padding:40px 48px;box-sizing:border-box;background:#fff;color:#111;font:13px/1.6 system-ui,sans-serif}
.ps-doc h1{text-align:center;font-size:24px;margin:0}.ps-sub{text-align:center;color:#555;margin:2px 0 16px}
.ps-doc h2{font-size:14px;margin:16px 0 6px}.ps-doc table{border-collapse:collapse;width:100%}
.ps-doc th,.ps-doc td{border:1px solid #c9d2dc;padding:6px 9px;text-align:left}.ps-doc th{background:#f2f5f8;white-space:nowrap;font-weight:700}
.ps-doc .r{text-align:right;white-space:nowrap}.ps-doc .c{color:#444;font-size:12px}.ps-doc tr.tot th,.ps-doc tr.tot td{background:#eef6f2;font-weight:800}
.ps-doc .net{margin-top:10px}.ps-doc .net th,.ps-doc .net td{font-size:15px;font-weight:800;background:#e3f2ea}
.ps-t{table-layout:fixed}.ps-t th:first-child{width:28%}.ps-t td.r,.ps-t th.r{width:22%}.ps-note{margin:0;color:#333}.ps-foot{margin-top:18px;color:#777;font-size:11px}`;

  async function makePdf(x) {
    await loadLibs();
    const stage = document.createElement('div');
    stage.style.cssText = 'position:fixed;left:-10000px;top:0;background:#fff';
    stage.innerHTML = `<style>${CSS}</style>${html(x)}`;
    document.body.append(stage);
    try {
      const canvas = await window.html2canvas(stage.querySelector('.ps-doc'), { scale: 2, backgroundColor: '#ffffff' });
      const { jsPDF } = window.jspdf, pdf = new jsPDF({ unit: 'mm', format: 'a4', compress: true });
      const img = canvas.toDataURL('image/jpeg', 0.9), w = 210, h = canvas.height * w / canvas.width, f = Math.min(1, 297 / h);
      pdf.addImage(img, 'JPEG', (210 - w * f) / 2, 0, w * f, h * f);
      return pdf.output('arraybuffer');
    } finally { stage.remove(); }
  }

  function save(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }
  const fileOf = r => new File([r.blob], r.fileName, { type: 'application/pdf' });
  const canShareFile = r => !!(navigator.canShare && navigator.canShare({ files: [fileOf(r)] }));

  function showResult(r) {
    const box = document.getElementById('psResult'); box.hidden = false; box.replaceChildren();
    const t = document.createElement('p'); t.innerHTML = `✅ <b>${esc(r.name)}</b>님 ${esc(r.monthLabel)} 임금명세서를 발급했습니다. <small>(실지급액 ${n(r.net)}원 · 발급 기록 저장됨)</small>`;
    const share = document.createElement('button'); share.type = 'button'; share.textContent = '카카오톡·문자로 보내기';
    const dl = document.createElement('button'); dl.type = 'button'; dl.className = 'secondary'; dl.textContent = 'PDF 저장'; dl.style.marginLeft = '6px';
    const msg = document.createElement('p'); msg.className = 'muted'; msg.style.margin = '6px 0 0';
    share.hidden = !canShareFile(r);
    if (share.hidden) msg.textContent = '이 기기에서는 바로 보내기를 지원하지 않습니다. PDF를 저장한 뒤 카카오톡·문자에 첨부해 보내 주세요.';
    share.onclick = async () => {
      try { await navigator.share({ files: [fileOf(r)], title: '임금명세서', text: `${r.name}님, ${r.biz} ${r.monthLabel} 임금명세서입니다. (지급일 ${r.payday})` }); msg.textContent = '보냈습니다.'; }
      catch (e) { if (e.name !== 'AbortError') msg.textContent = '보내지 못했습니다. PDF를 저장해 직접 보내 주세요.'; }
    };
    dl.onclick = () => save(r.blob, r.fileName);
    box.append(t, share, dl, msg);
    box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  /** 정산표의 '명세서' 버튼에서 호출. 성공하면 서버 발급 기록을 반환 */
  window.issuePayslip = async function (p, month, payday) {
    if (!p || !p.registered || p.gross == null) throw Error('일급이 등록된 작업자만 발급할 수 있습니다.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(payday || '')) throw Error('지급일을 선택하세요.');
    if (p.minOk === false) throw Error(p.name + '님 일급이 최저임금에 미달합니다. 일급을 고친 뒤 발급하세요.');
    const x = data(p, month, payday), buf = await makePdf(x);
    const sha = [...new Uint8Array(await crypto.subtle.digest('SHA-256', buf))].map(v => v.toString(16).padStart(2, '0')).join('');
    const log = await api('payslips', 'POST', { month, visitorId: p.visitorId, payday, sha256: sha });
    const [y, m] = month.split('-'), monthLabel = `${y}년 ${Number(m)}월분`;
    const r = { blob: new Blob([buf], { type: 'application/pdf' }), fileName: `임금명세서_${p.name}_${month}.pdf`, name: p.name, monthLabel, payday, net: log.net, biz: x.biz.name || '' };
    last = r; showResult(r); return log;
  };
  window.payslipSplit = split;
})();
