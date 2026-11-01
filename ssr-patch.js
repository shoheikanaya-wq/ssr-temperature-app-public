/* ssr-patch.js : SSR温測 追加機能
 *  1. 上から下へのスワイプで画面が再読み込みされるのを止める
 *  2. 入力した値を自動で保存し、再読み込みしても復元する
 *  3. ブラウザで開いた人に「アプリとして追加する方法」を案内する
 *  4. アプリとして開いたとき、配布用QR欄を隠す(下の設定で切り替え可能)
 */
(function () {
  'use strict';

  /* ★設定 ------------------------------------------------ */
  var HIDE_DISTRIBUTE_IN_APP = false; // true: アプリとして開いたとき配布用QR欄を隠す / false: 表示する
  var STORE_KEY = 'ssr_temp_autosave_v1';
  /* ------------------------------------------------------ */

  var ua = navigator.userAgent || '';
  var isStandalone =
    (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) ||
    window.navigator.standalone === true;
  var isIOS =
    /iPhone|iPad|iPod/.test(ua) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  var isAndroid = /Android/.test(ua);
  var inAppBrowser = /Line\/|FBAN|FBAV|Instagram|MicroMessenger|Twitter/i.test(ua);

  /* ===== 1. 引っ張って更新(上から下へのスワイプ)を止める ===== */
  var st = document.createElement('style');
  st.textContent = 'html,body{overscroll-behavior-y:none;}';
  document.head.appendChild(st);

  // 指の下にある要素の親をたどり、「縦にスクロールでき、まだ上に戻れる」ものがあれば true
  function innerCanScrollUp(el) {
    while (el && el !== document.body && el !== document.documentElement) {
      if (el.nodeType === 1) {
        var oy = window.getComputedStyle(el).overflowY;
        if ((oy === 'auto' || oy === 'scroll') && el.scrollHeight > el.clientHeight && el.scrollTop > 0) {
          return true;
        }
      }
      el = el.parentNode;
    }
    return false;
  }

  var sx = 0;
  var sy = 0;
  document.addEventListener(
    'touchstart',
    function (e) {
      if (e.touches.length === 1) {
        sx = e.touches[0].clientX;
        sy = e.touches[0].clientY;
      }
    },
    { passive: true }
  );
  document.addEventListener(
    'touchmove',
    function (e) {
      if (e.touches.length !== 1) return;
      var dx = e.touches[0].clientX - sx;
      var dy = e.touches[0].clientY - sy;
      var atTop = (window.pageYOffset || document.documentElement.scrollTop || 0) <= 0;
      // 画面の一番上で、縦方向に下へ引っ張られたときだけ止める(横スクロールは邪魔しない)
      // ただし、指が置かれた場所が「中でスクロールできる表示」(出力画面など)で、上にまだ戻れる場合は止めない
      if (atTop && dy > 0 && Math.abs(dy) > Math.abs(dx) && !innerCanScrollUp(e.target)) {
        if (e.cancelable) e.preventDefault();
      }
    },
    { passive: false }
  );

  /* ===== 2. 入力の自動保存と復元 ===== */
  function fields() {
    return document.querySelectorAll('.temp-input, .ssr-input, .kounai-input');
  }
  function keyOf(inp) {
    if (inp.id) return 'id:' + inp.id;
    return (inp.dataset.panel || '') + '|' + (inp.dataset.cell || '') + '|' + (inp.dataset.role || '');
  }

  var lastJson = '';
  var VERSION = 'v4';
  var badgeEl = null;
  function badge(text) {
    try {
      if (!badgeEl) {
        badgeEl = document.createElement('div');
        badgeEl.id = 'ssrPatchBadge';
        badgeEl.style.cssText =
          'position:fixed;left:6px;bottom:6px;z-index:9990;padding:2px 8px;border-radius:10px;' +
          'background:rgba(255,255,255,.88);border:1px solid #c8d0dc;font-size:11px;color:#444;pointer-events:none;';
        document.body.appendChild(badgeEl);
      }
      badgeEl.textContent = text;
    } catch (e) {}
  }
  function hhmm() {
    var d = new Date();
    return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
  }

  function save() {
    try {
      var dateEl = document.getElementById('dateInput');
      var data = { inputs: {}, date: dateEl ? dateEl.value : '' };
      fields().forEach(function (inp) {
        if (inp.value !== '') data.inputs[keyOf(inp)] = inp.value;
      });
      // 入力が1つもないときは、保存済みの内容を消さない(消すのは「クリア」を押したときだけ)
      if (Object.keys(data.inputs).length === 0) return;
      var json = JSON.stringify(data);
      if (json === lastJson) return;
      localStorage.setItem(STORE_KEY, json);
      lastJson = json;
      badge('🛠' + VERSION + ' 💾 保存済み ' + hhmm());
    } catch (e) {
      badge('🛠' + VERSION + ' ⚠ 保存できません');
    }
  }

  // 「クリア」を押して、本当に全部が空になったときだけ、保存内容も消す
  function afterClear() {
    var any = false;
    fields().forEach(function (inp) {
      if (inp.value !== '') any = true;
    });
    if (any) return; // 確認で「いいえ」を選んだ場合
    try {
      localStorage.removeItem(STORE_KEY);
    } catch (e) {}
    lastJson = '';
    badge('🛠' + VERSION + ' 保存内容を消しました');
  }

  function notice(msg) {
    var n = document.createElement('div');
    n.style.cssText =
      'position:fixed;top:0;left:0;right:0;z-index:9998;background:#fff8c5;' +
      'border-bottom:1px solid #d4a72c;padding:8px 12px;font-size:13px;color:#333;' +
      'display:flex;gap:8px;align-items:center;justify-content:space-between;';
    var s = document.createElement('span');
    s.textContent = msg;
    var b = document.createElement('button');
    b.type = 'button';
    b.textContent = '✕';
    b.style.cssText = 'border:none;background:transparent;font-size:16px;cursor:pointer;';
    b.onclick = function () {
      if (n.parentNode) n.parentNode.removeChild(n);
    };
    n.appendChild(s);
    n.appendChild(b);
    document.body.appendChild(n);
    setTimeout(function () {
      if (n.parentNode) n.parentNode.removeChild(n);
    }, 12000);
  }

  function restore() {
    try {
      var raw = localStorage.getItem(STORE_KEY);
      if (!raw) return;
      var data = JSON.parse(raw);
      var count = 0;
      fields().forEach(function (inp) {
        var v = data.inputs[keyOf(inp)];
        if (v !== undefined) {
          inp.value = v;
          inp.dispatchEvent(new Event('input', { bubbles: true })); // 色分け表示を反映
          count++;
        }
      });
      if (data.date) {
        var d = document.getElementById('dateInput');
        if (d) d.value = data.date;
      }
      lastJson = raw;
      badge('🛠' + VERSION + ' ♻ ' + count + '欄を復元');
      if (count > 0) {
        notice(
          '前回の入力' +
            (data.date ? '(' + data.date + ')' : '') +
            'を復元しました。新しく始めるときは「クリア」を押してください。'
        );
      }
    } catch (e) {
      /* 復元できなくても、アプリは普通に動かす */
    }
  }

  /* ===== 3. アプリとして追加する案内 ===== */
  var deferredPrompt = null;
  var barEl = null;

  function removeBar() {
    if (barEl && barEl.parentNode) barEl.parentNode.removeChild(barEl);
    barEl = null;
    document.body.style.paddingBottom = '';
  }

  function renderBar() {
    if (isStandalone) return;
    try {
      if (sessionStorage.getItem('ssr_install_bar_closed') === '1') return;
    } catch (e) {}
    var message = '';
    var showButton = false;

    if (inAppBrowser) {
      message =
        '⚠ このアプリ内のブラウザでは、ホーム画面に追加できません。右上または右下の「…」から' +
        '「ブラウザで開く」(SafariまたはChrome)を選んでください。';
    } else if (isIOS) {
      message =
        '📲 iPhone: Safariで開き、画面下の共有ボタン(四角と上向き矢印)から' +
        '「ホーム画面に追加」を選ぶと、アプリとして使えます。';
    } else if (isAndroid) {
      if (deferredPrompt) {
        message = '📲 ホーム画面に追加すると、アプリとして使えます。';
        showButton = true;
      } else {
        message =
          '📲 Android: Chromeの右上の「︙」から「アプリをインストール」' +
          '(または「ホーム画面に追加」)を選ぶと、アプリとして使えます。';
      }
    } else {
      return; // パソコンでは案内を出さない
    }

    removeBar();
    barEl = document.createElement('div');
    barEl.style.cssText =
      'position:fixed;left:0;right:0;bottom:0;z-index:9999;background:#1a73e8;color:#fff;' +
      'padding:10px 12px calc(10px + env(safe-area-inset-bottom,0px));font-size:13px;line-height:1.5;' +
      'display:flex;gap:8px;align-items:center;justify-content:space-between;';
    var s = document.createElement('span');
    s.textContent = message;
    barEl.appendChild(s);

    var right = document.createElement('span');
    right.style.cssText = 'display:flex;gap:6px;align-items:center;flex:0 0 auto;';
    if (showButton) {
      var ib = document.createElement('button');
      ib.type = 'button';
      ib.textContent = '追加する';
      ib.style.cssText =
        'background:#fff;color:#1a73e8;border:none;border-radius:6px;padding:8px 12px;font-weight:bold;font-size:14px;';
      ib.onclick = function () {
        if (!deferredPrompt) return;
        deferredPrompt.prompt();
        deferredPrompt.userChoice.then(function () {
          deferredPrompt = null;
          removeBar();
        });
      };
      right.appendChild(ib);
    }
    var cb = document.createElement('button');
    cb.type = 'button';
    cb.textContent = '✕';
    cb.style.cssText = 'background:transparent;color:#fff;border:none;font-size:18px;';
    cb.onclick = function () {
      try {
        sessionStorage.setItem('ssr_install_bar_closed', '1');
      } catch (e) {}
      removeBar();
    };
    right.appendChild(cb);
    barEl.appendChild(right);
    document.body.appendChild(barEl);
    document.body.style.paddingBottom = '90px'; // 案内バーで画面の下が隠れないように
  }

  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    deferredPrompt = e;
    renderBar();
  });
  window.addEventListener('appinstalled', removeBar);

  /* ===== 4. テンキーの「Enter(次へ)」で次の欄へ進む ===== */
  var FIELD_SELECTOR = '.kounai-input, .temp-input, .ssr-input';

  function chain() {
    // 画面の上から下、左から右の順(構内温度 → 盤内温度の上・下 → SSRの各欄 → 次の行 → 次の盤)
    return Array.prototype.slice.call(document.querySelectorAll(FIELD_SELECTOR)).filter(function (el) {
      return !el.disabled && el.offsetParent !== null;
    });
  }

  function selectAll(el) {
    try {
      el.setSelectionRange(0, el.value.length);
    } catch (e) {
      try {
        el.select();
      } catch (e2) {}
    }
  }

  function goNext(from) {
    var list = chain();
    var i = list.indexOf(from);
    if (i === -1) return;
    var next = list[i + 1];
    if (!next) {
      from.blur(); // 最後の欄まで入力したら、テンキーを閉じる
      return;
    }
    next.focus(); // 同じ操作の中で移動するので、テンキーは開いたままになる
    selectAll(next);
    setTimeout(function () {
      try {
        next.scrollIntoView({ block: 'center', behavior: 'smooth' });
      } catch (e) {}
    }, 50);
  }

  function setupEnterNav() {
    // テンキーを出し、キーボードのEnterキーを「次へ」の表示にする
    document.querySelectorAll(FIELD_SELECTOR).forEach(function (el) {
      el.setAttribute('inputmode', 'decimal');
      el.setAttribute('enterkeyhint', 'next');
    });

    document.addEventListener('keydown', function (e) {
      var t = e.target;
      if (!t || !t.matches || !t.matches(FIELD_SELECTOR)) return;
      if (e.key === 'Enter' || e.keyCode === 13) {
        e.preventDefault();
        goNext(t);
      }
    });

    // 欄をタップしたら中身を全選択する(打った数字で上書きできる)。打ち忘れの行は、その行の左端をタップすれば続けて入力できる
    document.addEventListener('focusin', function (e) {
      var t = e.target;
      if (t && t.matches && t.matches(FIELD_SELECTOR)) {
        setTimeout(function () {
          selectAll(t);
        }, 0);
      }
    });
  }

  /* ===== 起動処理 ===== */
  function init() {
    badge('🛠' + VERSION + ' 動作中');
    restore();
    setupEnterNav();

    // 自動保存(入力時・一定間隔・画面を離れるとき)
    var timer = null;
    document.addEventListener('input', function () {
      clearTimeout(timer);
      timer = setTimeout(save, 400);
    });
    setInterval(save, 3000); // Word読込など、入力イベントが出ない変更も拾う
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState === 'hidden') save();
    });
    window.addEventListener('pagehide', save);

    // 「クリア」ボタンのあと、保存内容も空にする(確認で「いいえ」なら変化なし)
    var clearBtn = document.getElementById('clearAllBtn');
    if (clearBtn) {
      clearBtn.addEventListener('click', function () {
        setTimeout(afterClear, 300);
      });
    }

    // アプリとして開いたときは、配布用QR欄を隠す
    if (isStandalone && HIDE_DISTRIBUTE_IN_APP) {
      var f = document.querySelector('.footer-distribute');
      if (f) f.style.display = 'none';
    }

    renderBar();

    // アプリとして追加できるようにするための裏方(常に最新の画面を取りに行き、通信できないときだけ保存分を使う)
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      navigator.serviceWorker.register('ssr-sw.js').catch(function () {});
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
