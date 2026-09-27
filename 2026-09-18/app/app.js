(function (global) {
  'use strict';

  // ⚠️ 미검증: 실제 배민·쿠팡이츠·요기요 정산 파일 샘플을 확보하지 못한 상태의 가정 열 이름.
  // 실제 파일 헤더와 다를 수 있다. 샘플 확보 후 이 객체만 교정하면 된다.
  // 쿠팡이츠 헤더는 1차 DIFF의 가정(주문ID·결제금액·중개수수료·프로모션비용)에서
  // 샘플3 기준(주문번호·매출·중개이용료·프로모션)으로 교체됨 — 어느 쪽이든 추측이다.
  var PLATFORM_CONFIG = {
    baemin: {
      label: '배달의민족',
      headers: {
        orderId: '주문번호',
        revenue: '매출',
        commission: '중개수수료',
        delivery: '배달비',
        payment: '결제수수료',
        promo: '프로모션'
      }
    },
    coupangeats: {
      label: '쿠팡이츠',
      headers: {
        orderId: '주문번호',
        revenue: '매출',
        commission: '중개이용료',
        delivery: '배달비',
        payment: '결제대행수수료',
        promo: '프로모션'
      }
    },
    yogiyo: {
      label: '요기요',
      headers: {
        orderId: '주문번호',
        revenue: '매출',
        commission: '중개수수료',
        delivery: '배달비',
        payment: 'PG수수료',
        promo: '프로모션'
      }
    }
  };

  var FIELD_KEYS = ['orderId', 'revenue', 'commission', 'delivery', 'payment', 'promo'];
  var NUMERIC_KEYS = ['revenue', 'commission', 'delivery', 'payment', 'promo'];
  var FIELD_LABELS = {
    revenue: '매출',
    commission: '중개수수료',
    delivery: '배달비',
    payment: '결제수수료',
    promo: '프로모션'
  };

  var MAX_FILE_BYTES = 10 * 1024 * 1024;

  function fail(reason, row) {
    var r = { error: true, reason: reason };
    if (row !== undefined) r.row = row;
    return r;
  }

  function safeAdd(a, b) {
    var sum = a + b;
    return Number.isSafeInteger(sum) ? sum : null;
  }

  function fnv1aHash(str) {
    var hash = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) {
      hash ^= str.charCodeAt(i);
      hash = Math.imul(hash, 0x01000193);
    }
    return (hash >>> 0).toString(16).padStart(8, '0');
  }

  function storageKey(platform, fileText) {
    return 'baedal-deposit:' + platform + ':' + fnv1aHash(fileText);
  }

  // 엑셀에서 저장한 한국어 CSV는 EUC-KR인 경우가 흔하다. UTF-8로 엄격하게 먼저 읽고 실패하면 EUC-KR.
  function decodeFileBuffer(buffer) {
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (e) {
      return new TextDecoder('euc-kr').decode(bytes);
    }
  }

  // RFC 4180 파서. 구조가 어긋나면 추측하지 않고 오류를 돌려준다.
  // - 따옴표는 필드 맨 앞에서만 열 수 있고, 닫힌 뒤에는 구분자(, 또는 줄바꿈)만 올 수 있다.
  // - 파일 끝까지 닫히지 않은 따옴표는 오류.
  // 각 레코드에는 파일상의 시작 줄 번호(line)를 붙인다 — 오류 안내에 쓴다.
  function parseCSVRecords(text) {
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);

    var records = [];
    var fields = [];
    var field = '';
    var line = 1;
    var recordLine = 1;
    var inQuotes = false;
    var quotedJustClosed = false;
    var fieldStarted = false;

    function endField() {
      fields.push(field);
      field = '';
      quotedJustClosed = false;
      fieldStarted = false;
    }
    function endRecord() {
      endField();
      records.push({ line: recordLine, fields: fields });
      fields = [];
    }

    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else { inQuotes = false; quotedJustClosed = true; }
        } else {
          if (ch === '\n') line++;
          field += ch;
        }
        continue;
      }
      if (ch === ',') { endField(); continue; }
      if (ch === '\r') {
        if (text[i + 1] === '\n') continue;
        return fail('줄바꿈 형식이 잘못됐습니다', line);
      }
      if (ch === '\n') {
        endRecord();
        line++;
        recordLine = line;
        continue;
      }
      if (quotedJustClosed) return fail('따옴표 뒤에 구분자가 없습니다', line);
      if (ch === '"') {
        if (fieldStarted) return fail('필드 중간에 따옴표가 있습니다', line);
        inQuotes = true;
        fieldStarted = true;
        continue;
      }
      field += ch;
      fieldStarted = true;
    }
    if (inQuotes) return fail('닫히지 않은 따옴표가 있습니다', recordLine);
    if (field.length > 0 || fields.length > 0 || quotedJustClosed) endRecord();

    var cleaned = records
      .map(function (r) {
        return { line: r.line, fields: r.fields.map(function (v) { return v.trim(); }) };
      })
      .filter(function (r) { return !(r.fields.length === 1 && r.fields[0] === ''); });
    return { error: false, records: cleaned };
  }

  var INT_PATTERN = /^-?(\d{1,3}(,\d{3})+|\d+)$/;
  var DECIMAL_PATTERN = /^-?(\d{1,3}(,\d{3})+|\d+)\.\d+$/;

  // 금액 문자열 → 정수. 천 단위 콤마는 3자리 간격일 때만 허용한다("1,00,0"은 오류).
  // 빈 칸·소수·범위 초과·그 밖의 형식은 전부 오류 — 0이나 반올림으로 바꾸지 않는다.
  function toAmount(str) {
    var s = String(str === undefined || str === null ? '' : str).trim();
    if (s === '') return { error: '빈 칸' };
    if (DECIMAL_PATTERN.test(s)) return { error: '정수가 아닌 금액' };
    if (!INT_PATTERN.test(s)) return { error: '숫자가 아님' };
    var n = Number(s.replace(/,/g, ''));
    if (!Number.isSafeInteger(n)) return { error: '처리할 수 없는 큰 금액' };
    return { value: n === 0 ? 0 : n };
  }

  function parseCSV(text, platform) {
    var config = PLATFORM_CONFIG[platform];
    if (!config) return fail('알 수 없는 플랫폼: ' + platform);

    var parsed = parseCSVRecords(text);
    if (parsed.error) return parsed;
    var records = parsed.records;
    if (records.length < 2) return fail('데이터 행이 없습니다');

    var header = records[0].fields;
    var seenHeader = Object.create(null);
    for (var d = 0; d < header.length; d++) {
      if (header[d] === '') continue;
      if (header[d] in seenHeader) return fail('헤더 중복: ' + header[d], records[0].line);
      seenHeader[header[d]] = true;
    }
    var idx = {};
    for (var h = 0; h < FIELD_KEYS.length; h++) {
      var name = config.headers[FIELD_KEYS[h]];
      var pos = header.indexOf(name);
      if (pos === -1) return fail('필수 헤더 누락: ' + name + ' (' + config.label + ' 기준)');
      idx[FIELD_KEYS[h]] = pos;
    }

    var rows = [];
    var seenOrders = Object.create(null);
    for (var i = 1; i < records.length; i++) {
      var fields = records[i].fields;
      var line = records[i].line;
      if (fields.length !== header.length) {
        return fail('열 개수 불일치 (헤더 ' + header.length + '개, 이 행 ' + fields.length + '개)', line);
      }

      var orderId = fields[idx.orderId];
      if (orderId === '') return fail('주문번호 없음', line);

      var amounts = {};
      for (var n = 0; n < NUMERIC_KEYS.length; n++) {
        var key = NUMERIC_KEYS[n];
        var amt = toAmount(fields[idx[key]]);
        if (amt.error) return fail(FIELD_LABELS[key] + ' 값 오류: ' + amt.error, line);
        amounts[key] = amt.value;
      }

      // 같은 주문번호가 다른 내용으로 다시 나오면 분할정산인지 알 수 없으므로 중단.
      // 완전히 같은 행의 재등장은 computeSettlement에서 합계 제외로 처리한다.
      var rawKey = JSON.stringify(fields);
      if (orderId in seenOrders) {
        if (seenOrders[orderId] !== rawKey) return fail('같은 주문번호(' + orderId + ')에 다른 내용', line);
      } else {
        seenOrders[orderId] = rawKey;
      }

      rows.push({
        rowNum: line,
        orderId: orderId,
        revenue: amounts.revenue,
        commission: amounts.commission,
        deliveryFee: amounts.delivery,
        paymentFee: amounts.payment,
        promotion: amounts.promo,
        rawKey: rawKey
      });
    }
    return { error: false, rows: rows };
  }

  function computeSettlement(rows) {
    var seen = Object.create(null);
    var duplicates = [];
    var uniqueRows = [];
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].rawKey in seen) {
        duplicates.push(rows[i].rowNum);
        continue;
      }
      seen[rows[i].rawKey] = true;
      uniqueRows.push(rows[i]);
    }

    var sums = { revenue: 0, commission: 0, delivery: 0, payment: 0, promo: 0 };

    for (var j = 0; j < uniqueRows.length; j++) {
      var row = uniqueRows[j];
      var fees = [row.commission, row.deliveryFee, row.paymentFee, row.promotion];

      // 부호 규칙: 매출 양수 행은 수수료 모두 0 이상, 환불(매출 음수) 행은 모두 0 이하.
      // 그 밖(매출 0 포함)은 파일 형식을 모르는 상태에서 해석할 수 없으므로 중단.
      if (row.revenue === 0) return fail('매출 0', row.rowNum);
      for (var f = 0; f < fees.length; f++) {
        if (row.revenue > 0 && fees[f] < 0) return fail('매출 양수인데 수수료 음수', row.rowNum);
        if (row.revenue < 0 && fees[f] > 0) return fail('매출 음수(환불)인데 수수료 양수', row.rowNum);
      }

      sums.revenue += row.revenue;
      sums.commission += row.commission;
      sums.delivery += row.deliveryFee;
      sums.payment += row.paymentFee;
      sums.promo += row.promotion;
      for (var k in sums) {
        if (!Number.isSafeInteger(sums[k])) return fail('합계가 처리할 수 있는 범위를 넘었습니다', row.rowNum);
      }
    }

    var feeSum = sums.commission;
    if (feeSum !== null) feeSum = safeAdd(feeSum, sums.delivery);
    if (feeSum !== null) feeSum = safeAdd(feeSum, sums.payment);
    if (feeSum !== null) feeSum = safeAdd(feeSum, sums.promo);
    if (feeSum === null) return fail('합계가 처리할 수 있는 범위를 넘었습니다');
    var netAmount = safeAdd(sums.revenue, -feeSum);
    if (netAmount === null) return fail('합계가 처리할 수 있는 범위를 넘었습니다');
    // 잠식률은 표시용 비율이라 소수 1자리 반올림. 매출 합계가 0 이하면 비율이 의미 없어 null.
    var erosionRate = sums.revenue > 0 ? Math.round((feeSum / sums.revenue) * 1000) / 10 : null;

    return {
      error: false,
      orderCount: uniqueRows.length,
      revenueSum: sums.revenue,
      commissionSum: sums.commission,
      deliverySum: sums.delivery,
      paymentSum: sums.payment,
      promoSum: sums.promo,
      feeSum: feeSum,
      netAmount: netAmount,
      erosionRate: erosionRate,
      duplicates: duplicates
    };
  }

  // 사용자가 형식을 볼 수 있도록 플랫폼별 예시 CSV를 만든다. 값은 예시일 뿐 실제 정산이 아니다.
  function templateCSV(platform) {
    var hd = PLATFORM_CONFIG[platform].headers;
    return [
      FIELD_KEYS.map(function (k) { return hd[k]; }).join(','),
      'EX001,25000,1950,3400,375,2000',
      'EX002,"18,000",1404,3400,270,0',
      'EX003,-8000,-624,0,-120,0'
    ].join('\n') + '\n';
  }

  function won(n) {
    return n.toLocaleString('ko-KR') + '원';
  }

  function el(tag, className, text) {
    var e = document.createElement(tag);
    if (className) e.className = className;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  function renderError(container, result) {
    container.textContent = '';
    var box = el('div', 'error');
    box.appendChild(el('strong', null, '계산하지 않았습니다'));
    box.appendChild(el('p', null, (result.row ? result.row + '번째 줄: ' : '') + result.reason));
    box.appendChild(el('p', 'hint', '형식을 모르는 값을 추측해서 계산하지 않도록 멈췄습니다. 파일의 해당 줄을 확인해 주세요.'));
    container.appendChild(box);
  }

  function renderResult(container, result, key, fileName) {
    container.textContent = '';

    container.appendChild(el('p', 'meta', fileName + ' · 주문 ' + result.orderCount + '건 계산'));

    var table = el('table', 'summary');
    var lines = [
      ['매출 합계', won(result.revenueSum), ''],
      ['중개수수료', '− ' + won(result.commissionSum), 'fee'],
      ['배달비', '− ' + won(result.deliverySum), 'fee'],
      ['결제수수료', '− ' + won(result.paymentSum), 'fee'],
      ['프로모션 부담', '− ' + won(result.promoSum), 'fee'],
      ['수수료 합계', '− ' + won(result.feeSum), 'subtotal'],
      ['실수령액', won(result.netAmount), 'total']
    ];
    lines.forEach(function (ln) {
      var tr = el('tr', ln[2]);
      tr.appendChild(el('th', null, ln[0]));
      tr.appendChild(el('td', null, ln[1]));
      table.appendChild(tr);
    });
    container.appendChild(table);

    container.appendChild(el('p', 'erosion', result.erosionRate === null
      ? '수수료 잠식률: 계산 불가 (매출 합계가 0 이하)'
      : '수수료 잠식률: 매출의 ' + result.erosionRate + '%'));

    if (result.duplicates.length > 0) {
      container.appendChild(el('p', 'warn',
        '완전히 같은 행 ' + result.duplicates.length + '건을 합계에서 뺐습니다 (' +
        result.duplicates.join(', ') + '번째 줄).'));
    }

    var label = el('label', 'deposit');
    var checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = 'depositConfirm';
    var status = el('p', 'storage-status');
    try {
      checkbox.checked = global.localStorage.getItem(key) === '1';
    } catch (e) {
      status.textContent = '이 브라우저에서는 입금 확인을 기기에 저장할 수 없습니다. 체크는 이 화면에서만 유지됩니다.';
    }
    checkbox.addEventListener('change', function () {
      try {
        if (checkbox.checked) global.localStorage.setItem(key, '1');
        else global.localStorage.removeItem(key);
        status.textContent = '';
      } catch (e) {
        status.textContent = '입금 확인을 기기에 저장하지 못했습니다(브라우저 저장공간 설정 확인). 체크는 이 화면에서만 유지됩니다.';
      }
    });
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(' 이 정산 입금 확인함 (이 기기에만 저장)'));
    container.appendChild(label);
    container.appendChild(status);
  }

  function init() {
    var fileInput = document.getElementById('csvFile');
    var platformSelect = document.getElementById('platformSelect');
    var resultEl = document.getElementById('result');
    var templateLink = document.getElementById('templateLink');
    var headerHint = document.getElementById('headerHint');
    if (!fileInput || !platformSelect || !resultEl) return;

    var currentReadToken = 0;

    function showHeaders() {
      if (!headerHint) return;
      var hd = PLATFORM_CONFIG[platformSelect.value].headers;
      headerHint.textContent = FIELD_KEYS.map(function (k) { return hd[k]; }).join(' · ');
    }

    function processFile(file) {
      // 토큰은 파일 유무와 상관없이 먼저 올린다 — 선택 해제 시 진행 중이던 읽기 결과도 버려야 한다.
      currentReadToken++;
      var myToken = currentReadToken;
      resultEl.textContent = '';
      if (!file) return;
      if (file.size > MAX_FILE_BYTES) {
        renderError(resultEl, fail('파일이 너무 큽니다 (10MB 이하만 처리)'));
        return;
      }

      var reader = new FileReader();
      reader.onload = function () {
        if (myToken !== currentReadToken) return;
        var text = decodeFileBuffer(reader.result);
        var platform = platformSelect.value;
        var parsed = parseCSV(text, platform);
        if (parsed.error) { renderError(resultEl, parsed); return; }
        var settled = computeSettlement(parsed.rows);
        if (settled.error) { renderError(resultEl, settled); return; }
        renderResult(resultEl, settled, storageKey(platform, text), file.name);
      };
      reader.onerror = function () {
        if (myToken !== currentReadToken) return;
        renderError(resultEl, fail('파일을 읽지 못했습니다. 다시 선택해 주세요.'));
      };
      reader.readAsArrayBuffer(file);
    }

    platformSelect.addEventListener('change', function () {
      showHeaders();
      processFile(fileInput.files[0]);
    });
    fileInput.addEventListener('change', function () {
      processFile(fileInput.files[0]);
    });

    if (templateLink) {
      templateLink.addEventListener('click', function (e) {
        e.preventDefault();
        var platform = platformSelect.value;
        var blob = new Blob(['﻿' + templateCSV(platform)], { type: 'text/csv;charset=utf-8' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = '배달정산-예시-' + platform + '.csv';
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
      });
    }

    showHeaders();
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', init);
  }

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
      PLATFORM_CONFIG: PLATFORM_CONFIG,
      fnv1aHash: fnv1aHash,
      storageKey: storageKey,
      decodeFileBuffer: decodeFileBuffer,
      parseCSVRecords: parseCSVRecords,
      toAmount: toAmount,
      parseCSV: parseCSV,
      computeSettlement: computeSettlement,
      templateCSV: templateCSV
    };
  }
})(this);
