/* 전자 근로계약서 — 관리자 휴대폰에서 작성 → 사업주·근로자 대면 서명 → PDF 보관·공유 → 근로자 링크 다운로드
   index.html 의 전역 함수($, api, me, roster, fmtPhone, showPanel)를 사용합니다. */
(() => {
  const LIBS = [
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
  ];
  const biz = () => window.LK_CONFIG?.business || {};
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const won = n => Number(n || 0).toLocaleString('ko-KR');
  const today = () => new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Seoul' }).format(new Date());
  const korDate = d => { const [y, m, dd] = String(d).split('-'); return y ? `${y}년 ${Number(m)}월 ${Number(dd)}일` : ''; };
  let libsLoaded = null, draft = null, pads = {}, result = null;

  function loadLibs() {
    if (libsLoaded) return libsLoaded;
    libsLoaded = Promise.all(LIBS.map(src => new Promise((ok, fail) => {
      const s = document.createElement('script'); s.src = src; s.onload = ok;
      s.onerror = () => { libsLoaded = null; fail(Error('PDF 도구를 불러오지 못했습니다. 인터넷 연결을 확인하세요.')); };
      document.head.append(s);
    })));
    return libsLoaded;
  }

  /* ---------- 서명 패드 ---------- */
  function signaturePad(canvas) {
    const ctx = canvas.getContext('2d'); let drawing = false, empty = true, last = null;
    function fit() {
      const r = window.devicePixelRatio || 1, w = canvas.clientWidth, h = canvas.clientHeight;
      canvas.width = w * r; canvas.height = h * r; ctx.setTransform(r, 0, 0, r, 0, 0);
      ctx.lineWidth = 2.6; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#111';
      empty = true;
    }
    const pos = e => { const b = canvas.getBoundingClientRect(); return { x: e.clientX - b.left, y: e.clientY - b.top }; };
    canvas.addEventListener('pointerdown', e => { e.preventDefault(); canvas.setPointerCapture(e.pointerId); drawing = true; last = pos(e); ctx.beginPath(); ctx.arc(last.x, last.y, 1.2, 0, Math.PI * 2); ctx.fill(); empty = false; });
    canvas.addEventListener('pointermove', e => { if (!drawing) return; const p = pos(e); ctx.beginPath(); ctx.moveTo(last.x, last.y); ctx.lineTo(p.x, p.y); ctx.stroke(); last = p; });
    ['pointerup', 'pointercancel', 'pointerleave'].forEach(t => canvas.addEventListener(t, () => { drawing = false; }));
    fit();
    return { clear: fit, isEmpty: () => empty, image: () => canvas.toDataURL('image/png') };
  }

  /* ---------- 계약서 본문(HTML) — 고용노동부 표준근로계약서 항목 기준 ---------- */
  function contractHtml(d, sig = {}) {
    const B = biz(), g = roster?.grades?.[d.grade];
    const period = d.start === d.end ? `${korDate(d.start)} (1일)` : `${korDate(d.start)}부터 ${korDate(d.end)}까지 (근로일 단위로 근로를 제공함)`;
    const sigImg = (src, label) => src ? `<img src="${src}" alt="${label}" style="height:46px;max-width:160px;vertical-align:middle">` : '<span style="color:#999">(서명)</span>';
    const row = (k, v) => `<tr><th>${k}</th><td>${v}</td></tr>`;
    return `<div class="ct-doc">
<h1>일용근로자 근로계약서</h1>
<p class="ct-sub">고용노동부 표준근로계약서 항목 기준 · 전자문서(근로기준법 제17조 제2항)</p>
<p><b>${esc(B.name || '')}</b> (이하 "사업주"라 함)과(와) <b>${esc(d.name)}</b> (이하 "근로자"라 함)은 다음과 같이 근로계약을 체결한다.</p>
<ol>
<li><b>근로계약기간:</b> ${period}</li>
<li><b>근무장소:</b> ${esc(d.site)}<br><small>※ 근로자는 사업주가 도급받은 업무를 수행하며, 업무 지시는 사업주의 현장책임자를 통해서만 받는다.</small></li>
<li><b>업무의 내용:</b> 반도체 자동화설비 조립 (자재 입고·부품 분류·조립·설치·셋업 등 전 공정을 순환하며 수행)<br>숙련 단계: <b>${esc(g ? g.label : d.grade)}</b><br><small>※ 일급은 그날 맡은 공정과 관계없이 근로자의 숙련 단계를 기준으로 하며, 성별 등에 따른 차별을 두지 않는다.</small></li>
<li><b>소정근로시간:</b> ${esc(d.timeStart)} ~ ${esc(d.timeEnd)} (휴게시간 ${esc(d.breakTime)})</li>
<li><b>근무일 및 휴일:</b> 매주 ${esc(d.days)} 근무, 주휴일 매주 ${esc(d.weeklyHoliday)}</li>
<li><b>임금</b><br>
 - 일급: <b>${won(d.wage)}원</b> (세전, 식대·교통비 포함)<br>
 - 주휴수당: ${esc(d.weeklyPay)}<br>
 - 연장·야간·휴일근로: 근로기준법에 따라 지급 (상시 5인 이상 사업장은 통상임금의 50% 가산)<br>
 - 임금지급일: ${esc(d.payday)} (휴일인 경우 전일 지급) · 지급방법: 근로자 명의 예금통장 입금<br>
 - 사업주는 임금 지급 시 임금명세서를 교부하며, 숙련 단계가 오르면 변경된 일급을 서면(전자문서)으로 통지한다.</li>
<li><b>연차유급휴가:</b> 근로기준법에서 정하는 바에 따라 부여한다.</li>
<li><b>사회보험:</b> 고용보험·산재보험 가입. 국민연금·건강보험은 법정 요건(월 8일 또는 60시간 이상 등) 충족 시 가입.</li>
<li><b>안전·보건:</b> 사업주는 채용 시 안전보건교육을 실시하고 보호구를 지급하며, 근로자는 안전수칙과 현장책임자의 안전 지시를 준수한다.</li>
<li><b>근로계약서 교부:</b> 이 계약서는 전자문서로 작성·서명하며, 사업주는 서명 완료 즉시 근로자에게 사본(PDF)을 교부한다. (근로기준법 제17조)</li>
<li><b>기타:</b> 이 계약에 정함이 없는 사항은 근로기준법령에 의한다.</li>
</ol>
<p class="ct-date">${korDate(d.signDate)}</p>
<div class="ct-sign">
<table><caption>사업주</caption>
${row('사업체명', esc(B.name || ''))}${row('대표자', esc(B.ceo || ''))}${row('사업자번호', esc(B.bizNo || ''))}${row('주소', esc(B.address || ''))}${row('전화', esc(B.phone || ''))}
${row('서명', `${esc(d.employerSigner)} ${sigImg(sig.employer, '사업주 서명')}`)}</table>
<table><caption>근로자</caption>
${row('성명', esc(d.name))}${row('생년월일', esc(d.birth || ''))}${row('연락처', esc(fmtPhone(d.phone)))}${row('주소', esc(d.address || ''))}
${row('서명', sigImg(sig.worker, '근로자 서명'))}</table>
</div>
<p class="ct-foot">계약번호 ${esc(d.id)} · ${sig.time ? '전자서명 완료 ' + esc(sig.time) + ' · ' : ''}현장 대면 서명${sig.witness ? ' (확인: ' + esc(sig.witness) + ')' : ''} · 원본 PDF의 SHA-256 고유값을 사업주가 보관합니다.</p>
</div>`;
  }

  /* ---------- 화면 ---------- */
  function step(n) {
    document.querySelectorAll('#contractPanel [data-step]').forEach(el => { el.hidden = Number(el.dataset.step) !== n; });
    $('#contractPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function readForm() {
    const v = id => $('#ct_' + id).value.trim();
    const d = { ...draft, name: v('name'), phone: v('phone'), birth: v('birth'), address: v('address'), grade: v('grade'), wage: Number(v('wage')),
      start: v('start'), end: v('end') || v('start'), site: v('site'), timeStart: v('timeStart'), timeEnd: v('timeEnd'), breakTime: v('breakTime'),
      days: v('days'), weeklyHoliday: v('weeklyHoliday'), weeklyPay: v('weeklyPay'), payday: v('payday'), employerSigner: v('employerSigner'), signDate: today() };
    const missing = [['name', '근로자 이름'], ['start', '계약 시작일'], ['site', '근무장소'], ['employerSigner', '사업주 서명자']].filter(([k]) => !d[k]).map(([, l]) => l);
    if (missing.length) throw Error(missing.join(', ') + '을(를) 입력하세요.');
    if (!Number.isInteger(d.wage) || d.wage < 10000 || d.wage > 1000000) throw Error('일급을 1만~100만원 사이로 입력하세요.');
    if (d.end < d.start) throw Error('계약 종료일이 시작일보다 빠릅니다.');
    return d;
  }

  window.openContract = function (worker) {
    if (!me?.admin) { $('#status').textContent = '승인된 관리자만 계약서를 작성할 수 있습니다.'; return; }
    const B = biz();
    draft = { id: crypto.randomUUID(), workerId: worker?.id || '' };
    result = null;
    const set = (id, val) => { $('#ct_' + id).value = val ?? ''; };
    $('#ct_grade').replaceChildren(...Object.entries(roster.grades || {}).map(([k, g]) => { const o = document.createElement('option'); o.value = k; o.textContent = g.label; return o; }));
    set('name', worker?.name); set('phone', fmtPhone(worker?.phone || '')); set('birth', ''); set('address', '');
    set('grade', worker?.grade || '입문'); set('wage', worker?.wage || roster.grades?.[worker?.grade || '입문']?.wage || '');
    set('start', today()); set('end', ''); set('site', B.site || '');
    set('timeStart', '08:30'); set('timeEnd', '17:30'); set('breakTime', '12:00~13:00 (1시간)');
    set('days', '월~금'); set('weeklyHoliday', '일요일'); set('weeklyPay', '요건 충족 시 별도 지급'); set('payday', '매월 10일');
    set('employerSigner', B.ceo || '');
    $('#ct_warn').textContent = (!B.ceo || !B.bizNo || !B.address) ? '⚠ config.js에 대표자·사업자번호·주소가 비어 있습니다. 계약서 사업주란이 빈칸으로 나옵니다.' : '';
    $('#ct_err').textContent = '';
    showPanel('contractPanel'); step(1);
  };

  function bind() {
    $('#ct_toPreview').onclick = () => {
      try { draft = readForm(); $('#ct_err').textContent = ''; } catch (e) { $('#ct_err').textContent = e.message; return; }
      $('#ct_preview').innerHTML = contractHtml(draft); step(2);
      pads.employer = signaturePad($('#ct_padEmployer')); $('#ct_signerName').textContent = draft.employerSigner;
    };
    $('#ct_back1').onclick = () => step(1);
    $('#ct_clearEmployer').onclick = () => pads.employer.clear();
    $('#ct_toWorker').onclick = () => {
      if (pads.employer.isEmpty()) { $('#ct_err2').textContent = '사업주 서명을 해 주세요.'; return; }
      $('#ct_err2').textContent = ''; draft.employerSig = pads.employer.image();
      $('#ct_workerName').textContent = draft.name; $('#ct_agree').checked = false; $('#ct_preview2').innerHTML = contractHtml(draft, { employer: draft.employerSig }); step(3);
      pads.worker = signaturePad($('#ct_padWorker'));
    };
    $('#ct_back2').onclick = () => step(2);
    $('#ct_clearWorker').onclick = () => pads.worker.clear();
    $('#ct_finish').onclick = finish;
    $('#ct_download').onclick = () => result && saveBlob(result.blob, result.fileName);
    $('#ct_share').onclick = share;
    $('#ct_copyLink').onclick = async () => { try { await navigator.clipboard.writeText(result.link); $('#ct_doneMsg').textContent = '다운로드 링크를 복사했습니다. 카카오톡·문자에 붙여넣어 보내세요.'; } catch { prompt('아래 링크를 복사하세요', result.link); } };
    $('#ct_mail').onclick = () => { location.href = 'mailto:?subject=' + encodeURIComponent('[근로계약서] ' + draft.name) + '&body=' + encodeURIComponent(shareText()); };
    $('#ct_toList').onclick = () => { showPanel('managePanel'); loadContracts().catch(() => {}); };
    $('#cd_get').onclick = publicDownload;
    $('#ct_agree').onchange = () => { $('#ct_err3').textContent = ''; };
  }

  async function finish() {
    if (!$('#ct_agree').checked) { $('#ct_err3').textContent = '계약 내용 확인에 체크해 주세요.'; return; }
    if (pads.worker.isEmpty()) { $('#ct_err3').textContent = '근로자 서명을 해 주세요.'; return; }
    const btn = $('#ct_finish'); btn.disabled = true; $('#ct_err3').textContent = 'PDF를 만들고 저장하는 중입니다…';
    try {
      await loadLibs();
      const sig = { employer: draft.employerSig, worker: pads.worker.image(), time: new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul', hour12: false }), witness: me.email || '' };
      const stage = $('#ct_render'); stage.innerHTML = contractHtml(draft, sig);
      const canvas = await window.html2canvas(stage.firstElementChild, { scale: 2, backgroundColor: '#ffffff' });
      const { jsPDF } = window.jspdf, pdf = new jsPDF({ unit: 'mm', format: 'a4', compress: true });
      const img = canvas.toDataURL('image/jpeg', 0.88), w = 210, h = canvas.height * w / canvas.width;
      for (let y = 0; y < h; y += 297) { if (y) pdf.addPage(); pdf.addImage(img, 'JPEG', 0, -y, w, h); }
      const buf = pdf.output('arraybuffer'); stage.innerHTML = '';
      const sha = [...new Uint8Array(await crypto.subtle.digest('SHA-256', buf))].map(x => x.toString(16).padStart(2, '0')).join('');
      const saved = await api('contracts', 'POST', { id: draft.id, workerId: draft.workerId, name: draft.name, phone: draft.phone, grade: draft.grade, wage: draft.wage,
        start: draft.start, end: draft.end, employerSigner: draft.employerSigner, pdf: toBase64(buf), sha256: sha });
      const link = location.origin + location.pathname + '#contract=' + saved.code;
      result = { blob: new Blob([buf], { type: 'application/pdf' }), fileName: saved.fileName, link };
      $('#ct_doneName').textContent = draft.name; $('#ct_doneLink').textContent = link; $('#ct_doneMsg').textContent = '';
      $('#ct_share').hidden = !(navigator.canShare && navigator.canShare({ files: [new File([result.blob], saved.fileName, { type: 'application/pdf' })] }));
      $('#ct_err3').textContent = ''; step(4);
    } catch (e) { $('#ct_err3').textContent = e.message; } finally { btn.disabled = false; }
  }

  const shareText = () => `${draft.name}님, ${biz().name || ''} 근로계약서입니다.\n아래 링크에서 언제든 다시 받을 수 있습니다.\n${result.link}`;
  async function share() {
    try { await navigator.share({ files: [new File([result.blob], result.fileName, { type: 'application/pdf' })], title: '근로계약서', text: shareText() }); }
    catch (e) { if (e.name !== 'AbortError') $('#ct_doneMsg').textContent = '공유하지 못했습니다. 다운로드 후 직접 보내 주세요.'; }
  }
  function toBase64(buf) { let s = ''; const b = new Uint8Array(buf); for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return btoa(s); }
  function fromBase64(b64) { const s = atob(b64), b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i); return new Blob([b], { type: 'application/pdf' }); }
  function saveBlob(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }

  /* ---------- 관리자: 계약서 목록 ---------- */
  window.loadContracts = async function () {
    const box = $('#contractList'); if (!box) return;
    const list = await api('contracts');
    box.replaceChildren();
    if (!list.length) { box.textContent = '아직 작성된 계약서가 없습니다.'; return; }
    const t = document.createElement('table'); t.innerHTML = '<tr><th>완료일</th><th>근로자</th><th>기간</th><th class="num">일급</th><th></th></tr>';
    for (const c of list) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${esc(new Date(c.completed).toLocaleDateString('ko-KR'))}</td><td>${esc(c.name)}<br><small>${esc(fmtPhone(c.phone))}</small></td><td>${esc(c.start)}${c.end !== c.start ? '<br>~ ' + esc(c.end) : ''}</td><td class="num">${won(c.wage)}</td>`;
      const td = document.createElement('td'), b = document.createElement('button'); b.type = 'button'; b.className = 'secondary'; b.textContent = 'PDF 받기';
      b.onclick = async () => { b.disabled = true; try { const f = await api('contracts/file?id=' + encodeURIComponent(c.id)); saveBlob(fromBase64(f.base64), f.fileName); } catch (e) { $('#payNote').textContent = e.message; } finally { b.disabled = false; } };
      td.append(b); tr.append(td); t.append(tr);
    }
    box.append(t);
  };

  /* ---------- 근로자: 링크로 다운로드 ---------- */
  async function publicDownload() {
    const code = (location.hash.match(/contract=([A-Za-z0-9]{12})/) || [])[1];
    const btn = $('#cd_get'); btn.disabled = true; $('#cd_msg').textContent = '불러오는 중…';
    try { const f = await api('contract/public', 'POST', { code }); saveBlob(fromBase64(f.base64), f.fileName); $('#cd_msg').textContent = f.name + '님 근로계약서를 내려받았습니다.'; }
    catch (e) { $('#cd_msg').textContent = e.message; } finally { btn.disabled = false; }
  }

  bind();
  const route = () => { if (/^#contract=[A-Za-z0-9]{12}$/.test(location.hash)) { showPanel('contractDownload'); $('#cd_msg').textContent = ''; } };
  window.addEventListener('hashchange', route); route();
})();
