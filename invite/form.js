/* 참석 폼 + 투표 현황 + 참석자 명단. 디자인은 각 초안의 CSS가 .rv-* 클래스로 입힙니다. */
(function () {
  const C = window.INVITE, R = window.RSVP, esc = R.esc;
  let rows = [];
  const openDates = new Set();

  function formHTML(L) {
    const open = R.voteOpen();
    return `
    <form class="rv-form" novalidate>
      <label class="rv-field"><span class="rv-label">${L.name}</span>
        <input name="name" maxlength="20" autocomplete="name" placeholder="이름 (같은 이름으로 다시 내면 수정돼요)" required></label>
      <div class="rv-field"><span class="rv-label">${L.attend}</span>
        <div class="rv-seg">
          <label><input type="radio" name="attend" value="yes" checked><span>갈게요</span></label>
          <label><input type="radio" name="attend" value="no"><span>이번엔 못 가요</span></label>
        </div></div>
      <div class="rv-field rv-when"><span class="rv-label">${open ? L.dates : '날짜'}</span>
        ${open ? `<p class="rv-hint">되는 날을 <b>전부</b> 골라주세요. 가장 많이 되는 날로 정할게요.</p>
        <div class="rv-chips">${C.dates.map(d => `<label><input type="checkbox" name="dates" value="${d}"><span>${R.fmt.mdw(d)}</span></label>`).join('')}</div>`
        : `<p class="rv-hint">${C.confirmedDate ? R.fmt.mdw(C.confirmedDate) + '로 확정됐어요' : '투표가 마감됐어요'}</p>`}
      </div>
      <div class="rv-field rv-yes"><span class="rv-label">${L.count}</span>
        <div class="rv-step"><button type="button" data-d="-1" aria-label="한 명 빼기">−</button><output name="countOut">1명</output><button type="button" data-d="1" aria-label="한 명 더">+</button></div>
        <input type="hidden" name="count" value="1"></div>
      <div class="rv-field rv-yes"><span class="rv-label">${L.car}</span>
        <div class="rv-seg">
          <label><input type="radio" name="hasCar" value="0" checked><span>대중교통·택시</span></label>
          <label><input type="radio" name="hasCar" value="1"><span>차 가져가요</span></label>
        </div>
        <input class="rv-car" name="car" maxlength="12" placeholder="차량번호 (예: 12가 3456)" hidden></div>
      <label class="rv-field"><span class="rv-label">${L.message}</span>
        <input name="message" maxlength="60" placeholder="선택 · 먹고 싶은 배달 메뉴도 환영"></label>
      <p class="rv-err" role="alert"></p>
      <button class="rv-submit" type="submit">${L.submit}</button>
      ${R.preview ? '<p class="rv-preview">미리보기 모드: 지금은 이 폰에만 저장돼요</p>' : ''}
    </form>`;
  }

  function resultsHTML(L) {
    const t = R.tally(rows);
    const open = R.voteOpen();
    const max = Math.max(1, ...Object.values(t.counts));
    const going0 = rows.filter(r => r.attend !== 'no');
    const voters = d => going0.filter(r => (r.dates || []).includes(d));
    const bars = C.confirmedDate ? '' : `
      <p class="rv-title">투표 현황 <span>날짜를 누르면 누가 골랐는지 보여요</span></p>
      <div class="rv-bars">${C.dates.map(d => {
        const vs = voters(d), on = openDates.has(d);
        return `
        <div class="rv-row ${on ? 'is-open' : ''}">
          <button type="button" class="rv-bar ${d === t.majority ? 'is-major' : ''} ${d === t.best && t.counts[d] ? 'is-best' : ''}" data-date="${d}" aria-expanded="${on}">
            <span class="rv-bar-l">${R.fmt.mdw(d)}</span>
            <span class="rv-bar-t"><i style="width:${(t.counts[d] / max) * 100}%"></i></span>
            <span class="rv-bar-n">${t.counts[d]}</span>
          </button>
          <div class="rv-voters" ${on ? '' : 'hidden'}>${vs.length ? vs.map(r => `<span>${esc(r.name)}${Number(r.count) > 1 ? ` +${Number(r.count) - 1}` : ''}</span>`).join('') : '<em>아직 아무도 안 골랐어요</em>'}</div>
        </div>`; }).join('')}
      </div>
      <p class="rv-sum">${t.n ? `${t.n}명 응답 · ${t.majority ? `<b>${R.fmt.md(t.majority)}</b>이 과반이에요` : '아직 과반인 날이 없어요'}` : '아직 아무도 안 냈어요. 1등 해주세요'}${open && C.voteDeadline ? ` · ${R.fmt.md(C.voteDeadline)} 마감` : ''}${R.preview ? '' : ` · <span class="rv-live">실시간</span>`}</p>`;
    const going = rows.filter(r => r.attend !== 'no');
    const cars = going.filter(r => r.car);
    const no = rows.filter(r => r.attend === 'no');
    return `${bars}
      <div class="rv-people">
        <p class="rv-sub">${L.going} <b>${t.people}</b>명</p>
        <ul class="rv-names">${going.map(r => `<li>${esc(r.name)}${Number(r.count) > 1 ? ` <small>+${Number(r.count) - 1}</small>` : ''}${r.message ? `<em>${esc(r.message)}</em>` : ''}</li>`).join('') || '<li class="rv-empty">아직 없어요</li>'}</ul>
        ${cars.length ? `<p class="rv-sub">방문차량 등록 <b>${cars.length}</b>대</p>
        <ul class="rv-cars">${cars.map(r => `<li><span>${esc(r.name)}</span><code>${esc(R.fmt.car(r.car))}</code></li>`).join('')}</ul>
        <p class="rv-hint">차량번호는 뒷자리만 보여요. 전체 번호는 저희만 봐요.</p>` : ''}
        ${no.length ? `<p class="rv-sub rv-no">아쉽지만 다음에: ${no.map(r => esc(r.name)).join(', ')}</p>` : ''}
      </div>`;
  }

  window.RSVPUI = {
    async mount({ form, results, labels = {}, onChange }) {
      const L = Object.assign({ name: '이름', attend: '참석', dates: '되는 날', count: '인원', car: '오는 방법', message: '한마디', submit: '보내기', going: '오는 사람' }, labels);
      form.innerHTML = formHTML(L);
      const f = form.querySelector('form');
      const render = () => { results.innerHTML = resultsHTML(L); onChange && onChange(R.tally(rows), rows); };
      const sync = () => {
        const yes = f.elements.attend.value === 'yes';
        f.querySelectorAll('.rv-yes, .rv-when').forEach(el => el.hidden = !yes);
        f.elements.car.hidden = f.elements.hasCar.value !== '1';
      };
      f.addEventListener('change', sync);
      f.querySelectorAll('.rv-step button').forEach(b => b.onclick = () => {
        const v = Math.max(1, Math.min(10, Number(f.elements.count.value) + Number(b.dataset.d)));
        f.elements.count.value = v; f.elements.countOut.value = v + '명';
      });
      f.elements.name.addEventListener('change', () => {
        const mine = rows.find(r => r.name.replace(/\s+/g, '') === f.elements.name.value.replace(/\s+/g, ''));
        if (!mine) return;
        f.elements.attend.value = mine.attend; f.elements.count.value = mine.count || 1; f.elements.countOut.value = (mine.count || 1) + '명';
        f.querySelectorAll('[name=dates]').forEach(c => c.checked = (mine.dates || []).includes(c.value));
        f.elements.hasCar.value = mine.car ? '1' : '0'; f.elements.car.value = mine.car || ''; f.elements.message.value = mine.message || '';
        sync();
      });
      f.onsubmit = async e => {
        e.preventDefault();
        const err = f.querySelector('.rv-err'); err.textContent = '';
        const yes = f.elements.attend.value === 'yes';
        const dates = [...f.querySelectorAll('[name=dates]:checked')].map(c => c.value);
        if (!f.elements.name.value.trim()) { err.textContent = '이름을 적어주세요'; f.elements.name.focus(); return; }
        if (yes && R.voteOpen() && !dates.length) { err.textContent = '되는 날을 하나 이상 골라주세요'; return; }
        if (yes && f.elements.hasCar.value === '1' && f.elements.car.value.replace(/\s/g, '').length < 7) { err.textContent = '차량번호를 끝까지 적어주세요'; f.elements.car.focus(); return; }
        const btn = f.querySelector('.rv-submit'); btn.disabled = true; const label = btn.textContent; btn.textContent = '보내는 중…';
        try {
          rows = await R.submit({ name: f.elements.name.value, attend: f.elements.attend.value, dates: yes ? dates : [], count: yes ? f.elements.count.value : 0,
            car: yes && f.elements.hasCar.value === '1' ? f.elements.car.value.trim() : '', message: f.elements.message.value.trim() });
          render(); btn.textContent = '받았어요! (고치려면 다시 보내기)';
          results.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } catch { err.textContent = '저장이 안 됐어요. 잠시 후 다시 눌러주세요'; btn.textContent = label; }
        btn.disabled = false;
      };
      sync();
      results.innerHTML = '<p class="rv-hint">불러오는 중…</p>';
      results.addEventListener('click', e => {
        const b = e.target.closest('.rv-bar[data-date]'); if (!b) return;
        const d = b.dataset.date; openDates.has(d) ? openDates.delete(d) : openDates.add(d);
        const row = b.parentElement, on = openDates.has(d);
        row.classList.toggle('is-open', on); b.setAttribute('aria-expanded', on); row.querySelector('.rv-voters').hidden = !on;
      });
      try { rows = await R.list(); } catch { rows = []; }
      render();
      if (!R.preview) {
        const refresh = async () => {
          if (document.hidden) return;
          try { const next = await R.list(); if (JSON.stringify(next) !== JSON.stringify(rows)) { rows = next; render(); } } catch {}
        };
        setInterval(refresh, 15000);
        document.addEventListener('visibilitychange', refresh);
      }
    },
  };
})();
