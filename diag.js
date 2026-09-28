// Secretary Diag — отчёт о клиенте Outlook для выбора пути (web add-in / урезанный / VBA).
// Намеренно ES5 без модулей: Outlook 2016 может открыть панель в движке IE11.
// Содержимое писем не читается: только наличие API, размеры и счётчики. Адрес — только домен.
(function () {
  'use strict';

  var VERSION = '0.1.0';
  var report = { diag_version: VERSION, generated_at: new Date().toISOString() };
  var el = function (id) {
    return document.getElementById(id);
  };

  function safe(fn) {
    try {
      return fn();
    } catch (e) {
      return 'error: ' + (e && e.message);
    }
  }

  function has(obj, name) {
    return !!obj && typeof obj[name] !== 'undefined';
  }

  function host(url) {
    var m = /^https?:\/\/([^/]+)/i.exec(url || '');
    return m ? m[1] : null;
  }

  function status(text) {
    el('status').textContent = text || '';
  }

  function render() {
    el('out').value = JSON.stringify(report, null, 2);
    el('summary').textContent = verdict();
  }

  // Развилка из plan-outlook.md (шаг 0.1).
  function verdict() {
    var mb = report.requirement_sets && report.requirement_sets.Mailbox;
    if (!mb) return 'office.js ещё не готов…';
    // Только непрерывный ряд: Outlook 2016 сравнивает версии как числа («1.10» = 1.1)
    // и отвечает true на 1.10+, не поддерживая 1.5.
    var max = '—';
    for (var i = 1; i <= 16; i++) {
      if (mb['1.' + i] !== true) break;
      max = '1.' + i;
    }
    var n = +max.split('.')[1] || 0;
    mb = { '1.8': n >= 8, '1.15': n >= 15 };
    var lines = [
      'Клиент: ' + (report.mailbox && report.mailbox.hostName) + ' ' + (report.mailbox && report.mailbox.hostVersion),
      'Сервер: ' + (report.exchange && report.exchange.guess),
      'Mailbox API: до ' + max,
    ];
    var path;
    if (mb['1.15'] === true) path = 'полный web add-in (мультивыбор, EML, вложения)';
    else if (mb['1.8'] === true) path = 'урезанный add-in (одно письмо, вложения; без EML)';
    else path = 'в этом клиенте — только VBA-макрос; проверьте Outlook в браузере';
    lines.push('Путь: ' + path);
    return lines.join('\n');
  }

  function collectStatic(info) {
    var ctx = Office.context;
    var mbx = ctx.mailbox;
    report.office = {
      host: info && info.host,
      platform: info && info.platform,
      diagnostics: safe(function () {
        var d = ctx.diagnostics;
        return d ? { host: d.host, version: d.version, platform: d.platform } : null;
      }),
      displayLanguage: safe(function () {
        return ctx.displayLanguage;
      }),
    };
    report.mailbox = safe(function () {
      var d = mbx.diagnostics;
      return { hostName: d.hostName, hostVersion: d.hostVersion, OWAView: d.OWAView || null };
    });

    var sets = { Mailbox: {} };
    for (var i = 1; i <= 16; i++) {
      var v = '1.' + i;
      sets.Mailbox[v] = safe(function () {
        return ctx.requirements.isSetSupported('Mailbox', v);
      });
    }
    var other = [
      ['DialogApi', '1.1'],
      ['DialogApi', '1.2'],
      ['DialogOrigin', '1.1'],
      ['IdentityAPI', '1.3'],
      ['NestedAppAuth', '1.1'],
      ['OpenBrowserWindowApi', '1.1'],
    ];
    for (var j = 0; j < other.length; j++) {
      (function (name, ver) {
        sets[name + ' ' + ver] = safe(function () {
          return ctx.requirements.isSetSupported(name, ver);
        });
      })(other[j][0], other[j][1]);
    }
    report.requirement_sets = sets;

    report.account = safe(function () {
      var p = mbx.userProfile;
      var email = p.emailAddress || '';
      return {
        accountType: p.accountType || null, // 1.6+: enterprise / office365 / gmail / outlookCom
        email_domain: email.indexOf('@') > 0 ? email.split('@')[1] : null,
        timeZone: p.timeZone || null,
      };
    });

    var ews = safe(function () {
      return mbx.ewsUrl;
    });
    var rest = safe(function () {
      return mbx.restUrl;
    });
    var ewsHost = host(ews);
    report.exchange = {
      ews_host: ewsHost,
      rest_host: host(rest),
      guess: !ewsHost
        ? 'неизвестно'
        : /outlook\.office365\.com|outlook\.office\.com/i.test(ewsHost)
          ? 'Exchange Online (Microsoft 365)'
          : 'Exchange on-premises (' + ewsHost + ')',
    };

    var item = mbx.item;
    report.apis = {
      'mailbox.getSelectedItemsAsync': has(mbx, 'getSelectedItemsAsync'),
      'mailbox.loadItemByIdAsync': has(mbx, 'loadItemByIdAsync'),
      'mailbox.masterCategories': has(mbx, 'masterCategories'),
      'mailbox.getCallbackTokenAsync': has(mbx, 'getCallbackTokenAsync'),
      'item.getAsFileAsync': has(item, 'getAsFileAsync'),
      'item.getAttachmentContentAsync': has(item, 'getAttachmentContentAsync'),
      'item.categories': has(item, 'categories'),
      'ui.displayDialogAsync': has(ctx.ui, 'displayDialogAsync'),
      'ui.openBrowserWindow': has(ctx.ui, 'openBrowserWindow'),
      'EventType.SelectedItemsChanged': !!(Office.EventType && Office.EventType.SelectedItemsChanged),
      'EventType.ItemChanged': !!(Office.EventType && Office.EventType.ItemChanged),
    };

    report.runtime = {
      userAgent: navigator.userAgent,
      engine: /Trident\//.test(navigator.userAgent)
        ? 'IE11 (Trident)'
        : /Edg\//.test(navigator.userAgent)
          ? 'Edge Chromium / WebView2'
          : /Edge\//.test(navigator.userAgent)
            ? 'Edge Legacy'
            : 'другой',
      es2017: safe(function () {
        return typeof new Function('return async () => 1')() === 'function';
      }),
      fetch: typeof window.fetch === 'function',
      Promise: typeof window.Promise === 'function',
      TextEncoder: typeof window.TextEncoder === 'function',
      crypto_subtle: !!(window.crypto && window.crypto.subtle),
      localStorage: safe(function () {
        window.localStorage.setItem('_d', '1');
        window.localStorage.removeItem('_d');
        return true;
      }),
    };
  }

  // Открытое письмо: наличие internetMessageId, вложения (кол-во/размеры), размер EML (если есть API).
  function checkItem() {
    var item = Office.context.mailbox.item;
    if (!item) {
      report.current_item = 'нет открытого письма';
      return render();
    }
    var atts = item.attachments || [];
    var r = {
      itemType: item.itemType,
      has_internetMessageId: !!item.internetMessageId,
      attachments: atts.length,
      attachments_total_bytes: 0,
      attachment_types: {},
    };
    for (var i = 0; i < atts.length; i++) {
      r.attachments_total_bytes += atts[i].size || 0;
      var t = (atts[i].isInline ? 'inline:' : '') + String(atts[i].attachmentType);
      r.attachment_types[t] = (r.attachment_types[t] || 0) + 1;
    }
    report.current_item = r;
    render();
    if (typeof item.getAsFileAsync === 'function') {
      status('Читаю размер EML…');
      var t0 = Date.now();
      item.getAsFileAsync(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          // base64 → примерный размер в байтах; содержимое не сохраняется
          r.eml_bytes_approx = Math.round((res.value.length * 3) / 4);
          r.eml_ms = Date.now() - t0;
        } else {
          r.eml_error = res.error && res.error.message;
        }
        status('');
        render();
      });
    }
  }

  function checkSelection() {
    var mbx = Office.context.mailbox;
    if (typeof mbx.getSelectedItemsAsync !== 'function') {
      report.selection = 'getSelectedItemsAsync недоступен';
      return render();
    }
    mbx.getSelectedItemsAsync(function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        report.selection = { error: res.error && res.error.message };
      } else {
        var s = { count: res.value.length, with_attachments: 0, types: {} };
        for (var i = 0; i < res.value.length; i++) {
          if (res.value[i].hasAttachment) s.with_attachments++;
          s.types[res.value[i].itemType] = (s.types[res.value[i].itemType] || 0) + 1;
        }
        report.selection = s;
      }
      render();
    });
  }

  // Доходит ли запрос из панели до адреса (no-cors: только факт соединения, ответ не читается).
  function probe(url, done) {
    var t0 = Date.now();
    if (typeof window.fetch === 'function') {
      window
        .fetch(url, { mode: 'no-cors', cache: 'no-store' })
        .then(function () {
          done({ ok: true, ms: Date.now() - t0 });
        })
        .catch(function (e) {
          done({ ok: false, error: String(e && e.message), ms: Date.now() - t0 });
        });
    } else {
      var x = new XMLHttpRequest();
      x.onload = function () {
        done({ ok: true, http: x.status, ms: Date.now() - t0 });
      };
      x.onerror = function () {
        done({ ok: false, error: 'xhr error (сеть или CORS)', ms: Date.now() - t0 });
      };
      x.open('GET', url);
      x.send();
    }
  }

  function checkNet() {
    var targets = [
      'https://www.googleapis.com/discovery/v1/apis?name=drive&preferred=true',
      'https://accounts.google.com/.well-known/openid-configuration',
    ];
    var extra = el('probe-url').value.trim();
    if (/^https:\/\//.test(extra)) targets.push(extra);
    report.network = {};
    var left = targets.length;
    status('Проверяю связь…');
    for (var i = 0; i < targets.length; i++) {
      (function (u) {
        probe(u, function (r) {
          report.network[host(u)] = r;
          if (--left === 0) status('');
          render();
        });
      })(targets[i]);
    }
  }

  function copy() {
    var ta = el('out');
    ta.focus();
    ta.select();
    var ok;
    try {
      ok = document.execCommand('copy');
    } catch (e) {
      ok = false;
    }
    status(ok ? 'Скопировано — вставьте в чат' : 'Выделено — нажмите Ctrl+C');
  }

  el('btn-item').onclick = checkItem;
  el('btn-sel').onclick = checkSelection;
  el('btn-net').onclick = checkNet;
  el('btn-copy').onclick = copy;

  if (typeof Office === 'undefined') {
    el('summary').textContent = 'office.js не загрузился — откройте панель из Outlook.';
    return;
  }
  Office.onReady(function (info) {
    collectStatic(info);
    render();
    var mbx = Office.context.mailbox;
    if (mbx && Office.EventType && Office.EventType.SelectedItemsChanged) {
      safe(function () {
        mbx.addHandlerAsync(Office.EventType.SelectedItemsChanged, checkSelection);
      });
    }
    if (mbx && typeof mbx.getSelectedItemsAsync === 'function') checkSelection();
    else if (mbx && mbx.item) checkItem();
  });
})();
