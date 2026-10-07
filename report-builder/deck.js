/* =====================================================================
   AX Insight to Concept · 단계별 결과 슬라이드 생성기 (Apple 키노트 스타일)
   YouTube 수집 → VOC 분석 → 페르소나 → 인터뷰 (→ 스토리보드) 순서로
   각 단계 결과의 핵심을 슬라이드에 담습니다.
   브라우저(window.buildResearchDeck)와 Node(module.exports) 양쪽에서 동작.
   ===================================================================== */
(function (root) {
  'use strict';

  const W = 13.333, H = 7.5, MX = 0.85;
  const FONT = 'Pretendard';
  const C = {
    ink: '1D1D1F', ink2: '424245', gray: '6E6E73', gray2: '86868B', line: 'D2D2D7',
    card: 'F5F5F7', white: 'FFFFFF', black: '000000', blue: '0071E3', blueDark: '2997FF',
  };
  const AVATAR = ['0071E3', 'FF9F0A', '34C759', 'BF5AF2'];
  const STEP_ORDER = ['youtube', 'voc', 'persona', 'interview', 'storyboard'];
  const STEP_DEFAULT = {
    youtube: { name: 'YouTube VOC 수집', tagline: '실제 사용자의 목소리를 모읍니다' },
    voc: { name: 'VOC 분석', tagline: '목소리 속 반복되는 패턴을 찾습니다' },
    persona: { name: '타깃 페르소나', tagline: '누구를 위해 만들지 정합니다' },
    interview: { name: '디자인 리서치 인터뷰', tagline: '숨은 니즈와 컨셉 방향을 확인합니다' },
    storyboard: { name: '스토리보드', tagline: '컨셉이 쓰이는 장면을 그립니다' },
  };

  const str = v => (v == null ? '' : String(v)).replace(/[ \t]+/g, ' ').trim();
  const one = v => str(v).replace(/\s*\n\s*/g, ' ');
  const clip = (v, n) => { const s = one(v); return s.length > n ? s.slice(0, n - 1) + '…' : s; };
  const arr = (v, n) => (Array.isArray(v) ? v : []).filter(x => x != null && x !== '').slice(0, n);
  const T = o => Object.assign({ fontFace: FONT, isTextBox: true, margin: 0, valign: 'top' }, o);
  // 원본 비율을 유지한 채 상자 안에 맞춘 위치·크기
  function fit(img, x, y, w, h, alignTop, alignLeft) {
    const r = (img.w > 0 && img.h > 0) ? img.w / img.h : 4 / 3;
    let iw = w, ih = w / r;
    if (ih > h) { ih = h; iw = h * r; }
    return { x: alignLeft ? x : x + (w - iw) / 2, y: alignTop ? y : y + (h - ih) / 2, w: iw, h: ih };
  }

  function makeBuilder(PptxGenJS) {
    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_WIDE';
    pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
    let page = 0;

    function slide(dark) {
      page += 1;
      const s = pres.addSlide();
      s.background = { color: dark ? C.black : C.white };
      if (page > 1) {
        s.addText(String(page), T({ x: W - MX - 1, y: 6.95, w: 1, h: 0.25, fontSize: 9, color: C.gray2, align: 'right', objectName: 'page-number' }));
      }
      return s;
    }

    // 본문 슬라이드 공통 머리: 단계 라벨 + 헤드라인
    function head(s, eyebrow, headline, size) {
      s.addText(clip(eyebrow, 30), T({ x: MX, y: 0.6, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.blue, objectName: 'eyebrow' }));
      s.addText(clip(headline, 60), T({ x: MX, y: 0.98, w: W - 2 * MX, h: 1.3, fontSize: size || 30, bold: true, color: C.ink, lineSpacingMultiple: 1.08, objectName: 'headline' }));
    }

    // ── 표지 ──────────────────────────────────────────────────────
    function cover(spec) {
      const s = slide(true);
      s.addText('AX Insight to Concept', T({ x: MX, y: 2.2, w: 8, h: 0.4, fontSize: 15, bold: true, color: C.gray2 }));
      s.addText(clip(spec.title, 46), T({ x: MX, y: 2.7, w: W - 2 * MX, h: 2.1, fontSize: 52, bold: true, color: C.white, lineSpacingMultiple: 1.05 }));
      s.addText(clip(spec.subtitle, 80), T({ x: MX, y: 4.95, w: 10.5, h: 0.9, fontSize: 20, color: 'A1A1A6', lineSpacingMultiple: 1.25 }));
      s.addText(clip(spec.date, 30), T({ x: MX, y: 6.55, w: 6, h: 0.3, fontSize: 12, color: C.gray }));
    }

    // ── 전체 흐름 ────────────────────────────────────────────────
    function overview(steps) {
      const s = slide(false);
      head(s, 'Process', '단계별로 정리한 리서치 결과');
      const n = steps.length, gap = 0.45;
      const w = (W - 2 * MX - gap * (n - 1)) / n;
      steps.forEach((st, i) => {
        const x = MX + i * (w + gap), y = 3.25, has = st.slides.length > 0 || (st.images || []).length > 0;
        s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color: has ? C.ink : C.line, width: 1.5 } });
        s.addText(String(i + 1).padStart(2, '0'), T({ x, y: y + 0.3, w, h: 0.8, fontSize: 40, bold: true, color: has ? C.blue : C.line }));
        const small = n >= 5;
        s.addText(clip(st.name, 16), T({ x, y: y + 1.25, w, h: 0.75, fontSize: small ? 16 : 18, bold: true, color: has ? C.ink : C.gray2, lineSpacingMultiple: 1.05 }));
        s.addText(has ? clip(st.tagline, 40) : '자료 없음', T({ x, y: y + 2.05, w, h: 1.0, fontSize: small ? 11.5 : 12.5, color: C.gray, lineSpacingMultiple: 1.3 }));
      });
    }

    // ── 단계 간지 ────────────────────────────────────────────────
    function divider(st, no) {
      const s = slide(true);
      s.addText('Step ' + String(no).padStart(2, '0'), T({ x: MX, y: 2.35, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.blueDark }));
      s.addText(clip(st.name, 22), T({ x: MX, y: 2.85, w: W - 2 * MX, h: 1.3, fontSize: 58, bold: true, color: C.white }));
      s.addText(clip(st.tagline, 60), T({ x: MX, y: 4.25, w: 11, h: 0.6, fontSize: 22, color: C.gray2 }));
      const files = arr(st.sources, 4).map(f => clip(f, 40));
      if (files.length) s.addText('근거 자료  ' + files.join('  ·  '), T({ x: MX, y: 6.5, w: 10.5, h: 0.3, fontSize: 11, color: C.gray }));
    }

    // ── 본문 슬라이드 유형 ───────────────────────────────────────
    const TYPES = {
      // 한 문장 메시지
      statement(s, d, eb) {
        s.addText(clip(eb, 30), T({ x: MX, y: 0.6, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.blue }));
        s.addText(clip(d.headline, 70), T({ x: MX, y: 1.9, w: W - 2 * MX, h: 2.6, fontSize: 42, bold: true, color: C.ink, valign: 'middle', lineSpacingMultiple: 1.1 }));
        if (str(d.body)) s.addText(clip(d.body, 160), T({ x: MX, y: 4.8, w: 10.5, h: 1.5, fontSize: 18, color: C.gray, lineSpacingMultiple: 1.4 }));
      },

      // 큰 숫자 2~3개
      numbers(s, d, eb) {
        head(s, eb, d.headline);
        const it = arr(d.items, 3), gap = 0.6;
        const w = (W - 2 * MX - gap * (it.length - 1)) / Math.max(1, it.length);
        it.forEach((x, i) => {
          const cx = MX + i * (w + gap);
          s.addText(clip(x.value, 9), T({ x: cx, y: 2.95, w, h: 1.45, fontSize: 72, bold: true, color: i === 0 ? C.blue : C.ink, fit: 'shrink' }));
          s.addShape(pres.shapes.LINE, { x: cx, y: 4.6, w: Math.min(w, 3.2), h: 0, line: { color: C.line, width: 1 } });
          s.addText(clip(x.label, 60), T({ x: cx, y: 4.8, w, h: 1.2, fontSize: 15, color: C.gray, lineSpacingMultiple: 1.35 }));
        });
      },

      // 카드 2~4개
      cards(s, d, eb) {
        head(s, eb, d.headline);
        const cs = arr(d.cards, 4), n = cs.length;
        const cols = n === 4 ? 2 : Math.max(1, n), rows = n === 4 ? 2 : 1, gap = 0.3;
        const top = 2.55, avail = 6.7 - top;
        const w = (W - 2 * MX - gap * (cols - 1)) / cols, h = rows === 1 ? Math.min(avail, 3.3) : (avail - gap * (rows - 1)) / rows;
        const narrow = cols >= 3, tH = narrow ? 0.9 : 0.5;
        cs.forEach((c, i) => {
          const x = MX + (i % cols) * (w + gap), y = top + Math.floor(i / cols) * (h + gap);
          s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.18, fill: { color: C.card }, line: { color: C.card } });
          s.addText(clip(c.title, 24), T({ x: x + 0.4, y: y + 0.35, w: w - 0.8, h: tH, fontSize: narrow ? 17 : 19, bold: true, color: C.ink, lineSpacingMultiple: 1.05 }));
          s.addText(clip(c.text, rows === 2 ? 90 : 130), T({ x: x + 0.4, y: y + 0.45 + tH, w: w - 0.8, h: h - 0.7 - tH, fontSize: 13.5, color: C.ink2, lineSpacingMultiple: 1.4 }));
        });
      },

      // 대표 인용 1~3개
      quotes(s, d, eb) {
        const qs = arr(d.quotes, 3);
        if (qs.length === 1) {
          s.addText(clip(eb, 30), T({ x: MX, y: 0.6, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.blue }));
          s.addText('“', T({ x: MX, y: 1.3, w: 1.2, h: 1.2, fontSize: 96, bold: true, color: C.blue, fontFace: 'Georgia' }));
          s.addText(clip(qs[0].text, 110), T({ x: MX, y: 2.45, w: W - 2 * MX, h: 2.8, fontSize: 32, bold: true, color: C.ink, lineSpacingMultiple: 1.25 }));
          s.addText(clip(qs[0].who, 50), T({ x: MX, y: 5.6, w: 10, h: 0.4, fontSize: 15, color: C.gray }));
          return;
        }
        head(s, eb, d.headline);
        const top = 2.6, rh = (6.7 - top) / Math.max(1, qs.length);
        qs.forEach((q, i) => {
          const y = top + i * rh;
          if (i > 0) s.addShape(pres.shapes.LINE, { x: MX, y: y - 0.1, w: W - 2 * MX, h: 0, line: { color: C.line, width: 0.75 } });
          s.addText('“', T({ x: MX, y: y, w: 0.6, h: 0.7, fontSize: 44, bold: true, color: C.blue, fontFace: 'Georgia' }));
          s.addText(clip(q.text, 95), T({ x: MX + 0.75, y: y + 0.08, w: W - 2 * MX - 0.75, h: rh - 0.55, fontSize: 18, color: C.ink, lineSpacingMultiple: 1.3 }));
          s.addText(clip(q.who, 50), T({ x: MX + 0.75, y: y + rh - 0.5, w: 8, h: 0.3, fontSize: 11.5, color: C.gray2 }));
        });
      },

      // 페르소나 카드 최대 3명
      personas(s, d, eb) {
        head(s, eb, d.headline);
        const ps = arr(d.personas, 3), gap = 0.3;
        const w = (W - 2 * MX - gap * (ps.length - 1)) / Math.max(1, ps.length), y = 2.4, h = 4.4;
        ps.forEach((p, i) => {
          const x = MX + i * (w + gap), col = AVATAR[i % AVATAR.length];
          s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.18, fill: { color: C.card }, line: { color: C.card } });
          s.addShape(pres.shapes.OVAL, { x: x + 0.4, y: y + 0.35, w: 0.7, h: 0.7, fill: { color: col }, line: { color: col } });
          s.addText(clip(one(p.name).charAt(0) || '·', 1), T({ x: x + 0.4, y: y + 0.35, w: 0.7, h: 0.7, fontSize: 19, bold: true, color: C.white, align: 'center', valign: 'middle' }));
          s.addText(clip(p.name, 18), T({ x: x + 0.4, y: y + 1.2, w: w - 0.8, h: 0.42, fontSize: 19, bold: true, color: C.ink }));
          s.addText(clip(p.meta, 40), T({ x: x + 0.4, y: y + 1.62, w: w - 0.8, h: 0.35, fontSize: 11.5, color: C.gray }));
          s.addText('니즈', T({ x: x + 0.4, y: y + 2.1, w: w - 0.8, h: 0.25, fontSize: 10.5, bold: true, color: col }));
          s.addText(clip(p.need, 60), T({ x: x + 0.4, y: y + 2.37, w: w - 0.8, h: 0.8, fontSize: 12.5, color: C.ink2, lineSpacingMultiple: 1.3 }));
          s.addText('페인 포인트', T({ x: x + 0.4, y: y + 3.25, w: w - 0.8, h: 0.25, fontSize: 10.5, bold: true, color: col }));
          s.addText(clip(p.pain, 60), T({ x: x + 0.4, y: y + 3.52, w: w - 0.8, h: 0.8, fontSize: 12.5, color: C.ink2, lineSpacingMultiple: 1.3 }));
        });
      },

      // 번호 목록 3~5개
      list(s, d, eb) {
        head(s, eb, d.headline);
        const it = arr(d.items, 5), top = 2.55, rh = Math.min(1.0, (6.75 - top) / Math.max(1, it.length));
        it.forEach((x, i) => {
          const y = top + i * rh;
          s.addShape(pres.shapes.LINE, { x: MX, y, w: W - 2 * MX, h: 0, line: { color: C.line, width: 0.75 } });
          s.addText(String(i + 1).padStart(2, '0'), T({ x: MX, y: y + 0.16, w: 0.8, h: rh - 0.25, fontSize: 22, bold: true, color: C.blue }));
          s.addText(clip(x.title, 30), T({ x: MX + 0.95, y: y + 0.18, w: 3.9, h: rh - 0.25, fontSize: 17, bold: true, color: C.ink }));
          s.addText(clip(x.text, 110), T({ x: MX + 5.0, y: y + 0.2, w: W - 2 * MX - 5.0, h: rh - 0.25, fontSize: 13, color: C.ink2, lineSpacingMultiple: 1.3 }));
        });
      },
    };

    // ── 스토리보드 장면 (실제 이미지) ───────────────────────────
    function sceneGrid(imgs, captions, eb) {
      const s = slide(false);
      head(s, eb, '장면으로 보는 사용 시나리오');
      const list = imgs.slice(0, 6), n = list.length;
      const cols = n <= 4 ? n : 3, rows = Math.ceil(n / cols), gap = 0.3;
      const top = 2.45, capH = 0.6;
      const w = (W - 2 * MX - gap * (cols - 1)) / cols;
      const cellH = (6.75 - top - gap * (rows - 1)) / rows, imgH = cellH - capH;
      list.forEach((img, i) => {
        const x = MX + (i % cols) * (w + gap);
        let y = top + Math.floor(i / cols) * (cellH + gap);
        if (rows === 1) y += Math.max(0, (cellH - capH - fit(img, x, 0, w, imgH, true).h) / 2) - 0.2;
        const f = fit(img, x, y, w, imgH, true);
        s.addImage({ data: img.data, x: f.x, y: f.y, w: f.w, h: f.h, objectName: 'scene-thumb-' + (i + 1) });
        s.addText([
          { text: String(i + 1).padStart(2, '0') + '  ', options: { bold: true, color: C.blue } },
          { text: clip(captions[i], cols >= 3 ? 34 : 50), options: { color: C.ink2 } },
        ], T({ x: f.x, y: f.y + f.h + 0.1, w: Math.max(f.w, 1.5), h: capH - 0.1, fontSize: 11, lineSpacingMultiple: 1.25 }));
      });
    }

    function scene(img, caption, no, eb) {
      const s = slide(false);
      s.addText(clip(eb, 30), T({ x: MX, y: 0.6, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.blue }));
      const f = fit(img, MX, 1.15, 8.3, 5.6, false, true);
      s.addImage({ data: img.data, x: f.x, y: f.y, w: f.w, h: f.h, objectName: 'scene-' + no });
      const rx = f.x + f.w + 0.6, rw = W - MX - rx;
      s.addText('Scene ' + String(no).padStart(2, '0'), T({ x: rx, y: 1.25, w: rw, h: 0.35, fontSize: 14, bold: true, color: C.blue }));
      s.addText(clip(caption, 90), T({ x: rx, y: 1.75, w: rw, h: 4.6, fontSize: 20, bold: true, color: C.ink, lineSpacingMultiple: 1.35 }));
    }

    // ── 종합 요약 ────────────────────────────────────────────────
    function summary(sm) {
      const s = slide(true);
      s.addText('Summary', T({ x: MX, y: 0.6, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.blueDark }));
      s.addText(clip(sm.headline, 60), T({ x: MX, y: 0.98, w: W - 2 * MX, h: 1.4, fontSize: 34, bold: true, color: C.white, lineSpacingMultiple: 1.08 }));
      const pts = arr(sm.points, 4), top = 2.75, rh = 0.82;
      pts.forEach((p, i) => {
        const y = top + i * rh;
        s.addText(String(i + 1).padStart(2, '0'), T({ x: MX, y, w: 0.8, h: rh - 0.15, fontSize: 18, bold: true, color: C.blueDark, valign: 'middle' }));
        s.addText(clip(p, 80), T({ x: MX + 0.9, y, w: W - 2 * MX - 0.9, h: rh - 0.15, fontSize: 18, color: C.white, valign: 'middle' }));
      });
      const nx = arr(sm.nextSteps, 4).map(x => clip(x, 30));
      if (nx.length) {
        s.addShape(pres.shapes.LINE, { x: MX, y: 6.15, w: W - 2 * MX, h: 0, line: { color: '333336', width: 0.75 } });
        s.addText([
          { text: '다음 단계   ', options: { bold: true, color: C.white } },
          { text: nx.join('   ·   '), options: { color: C.gray2 } },
        ], T({ x: MX, y: 6.3, w: W - 2 * MX - 1.2, h: 0.4, fontSize: 13, valign: 'middle' }));
      }
    }

    return {
      pres,
      build(spec, extras) {
        const given = Array.isArray(spec.steps) ? spec.steps : [];
        const imgs = arr(extras && extras.storyboardImages, 8).filter(im => im && im.data);
        const steps = STEP_ORDER.map(key => {
          const g = given.find(x => one(x && x.key).toLowerCase() === key) || {};
          const def = STEP_DEFAULT[key];
          return {
            key, name: one(g.name) || def.name, tagline: one(g.tagline) || def.tagline, sources: g.sources,
            slides: arr(g.slides, key === 'storyboard' && imgs.length ? 2 : 4).filter(sl => TYPES[one(sl && sl.type)]),
            images: key === 'storyboard' ? imgs : [],
            captions: key === 'storyboard' ? imgs.map((im, i) => {
              const sc = arr(g.scenes, 8)[i];
              return one(sc && (sc.caption || sc.text)) || one(im.desc);
            }) : [],
          };
        }).filter(st => st.key !== 'storyboard' || st.slides.length || st.images.length);

        cover(spec);
        overview(steps);
        let no = 0;
        steps.forEach(st => {
          if (!st.slides.length && !st.images.length) return;
          no += 1;
          divider(st, no);
          const eb = 'Step ' + String(no).padStart(2, '0') + ' · ' + st.name;
          st.slides.forEach(sl => {
            const s = slide(false);
            TYPES[one(sl.type)](s, sl, eb);
          });
          if (st.images.length) {
            if (st.images.length > 1) sceneGrid(st.images, st.captions, eb);
            st.images.forEach((im, i) => scene(im, st.captions[i], i + 1, eb));
          }
        });
        if (spec.summary && (str(spec.summary.headline) || arr(spec.summary.points, 4).length)) summary(spec.summary);
        return pres;
      },
    };
  }

  function buildResearchDeck(PptxGenJS, spec, extras) {
    return makeBuilder(PptxGenJS).build(spec || {}, extras || {});
  }
  buildResearchDeck.SLIDE_TYPES = ['statement', 'numbers', 'cards', 'quotes', 'personas', 'list'];
  buildResearchDeck.STEP_ORDER = STEP_ORDER;

  if (typeof module !== 'undefined' && module.exports) module.exports = buildResearchDeck;
  else root.buildResearchDeck = buildResearchDeck;
})(typeof window !== 'undefined' ? window : globalThis);
