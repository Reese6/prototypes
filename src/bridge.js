/**
 * Мост между HTML-документом в iframe и приложением.
 * Внедряется в srcdoc после script.js и до скриптов документа (см. HtmlFrame.tsx).
 * Origin приложения передаётся атрибутом data-app-origin тега <script>.
 *
 * Клик по ссылке, если документ сам не отменил переход:
 * - href="#..." — переход по фрагменту внутри iframe (:target, hashchange, история);
 * - target="_blank", клик с Ctrl/Cmd/Shift/Alt или средней кнопкой — поведение браузера;
 * - mailto:, tel: и другие схемы кроме http(s) — поведение браузера (внешнее приложение);
 * - URL на origin приложения ("/", "/about") — навигация в приложении через postMessage;
 * - прочие http(s)-ссылки — новая вкладка: иначе iframe ушёл бы с документа
 *   вместе с script.js.
 * Отправка формы методом GET на origin приложения тоже уходит в приложение;
 * остальные отправки, уводящие iframe с документа, отменяются.
 *
 * Не перехватываются: location.href = ..., form.submit(), <meta http-equiv="refresh">.
 */
(() => {
  'use strict';
  const appOrigin = document.currentScript?.dataset.appOrigin;
  const parentWindow = window.parent;
  if (!appOrigin || parentWindow === window) return;

  const isLink = node => node instanceof HTMLAnchorElement || node instanceof HTMLAreaElement;
  // Пустая цель, _self, _parent и _top — текущий контекст; остальные открывают новое окно.
  const opensNewWindow = target => {
    const value = (target ?? document.querySelector('base[target]')?.getAttribute('target') ?? '').trim().toLowerCase();
    return !['', '_self', '_parent', '_top'].includes(value);
  };

  function navigateApp(url) {
    parentWindow.postMessage({ type: 'html-frame:navigate', href: url.href }, appOrigin);
  }

  function goToFragment(fragment) {
    const before = location.href;
    try { location.hash = fragment; } catch { /* остаётся прокрутка ниже */ }
    if (location.href !== before) return;
    // Тот же фрагмент: браузер не прокручивает к нему повторно.
    let id = fragment.slice(1);
    try { id = decodeURIComponent(id); } catch { /* id без декодирования */ }
    const element = document.getElementById(id) ?? document.getElementsByName(id)[0];
    if (element) element.scrollIntoView();
    else if (!id || id.toLowerCase() === 'top') window.scrollTo(0, 0);
  }

  // Фаза всплытия: script.js (фаза перехвата) и обработчики документа уже отработали.
  window.addEventListener('click', event => {
    if (event.defaultPrevented || event.button !== 0
      || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
    const link = event.composedPath().find(isLink);
    if (!link || !link.hasAttribute('href') || link.hasAttribute('download')) return;
    const href = link.getAttribute('href').trim();
    if (href.startsWith('#')) {
      // В srcdoc "#id" разрешается от адреса <base>, и браузер загрузил бы в iframe другую страницу.
      event.preventDefault();
      goToFragment(href);
      return;
    }
    if (opensNewWindow(link.getAttribute('target'))) return;
    let url;
    try { url = new URL(link.href); } catch { return; }
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return;
    event.preventDefault();
    if (url.origin === appOrigin) navigateApp(url);
    else window.open(url.href, '_blank', 'noopener');
  });

  window.addEventListener('submit', event => {
    const form = event.target;
    if (event.defaultPrevented || !(form instanceof HTMLFormElement)) return;
    const submitter = event.submitter;
    const attr = name => submitter?.hasAttribute(`form${name}`)
      ? submitter.getAttribute(`form${name}`) : form.getAttribute(name);
    const method = attr('method')?.trim().toLowerCase();
    if (method === 'dialog' || opensNewWindow(attr('target'))) return;
    event.preventDefault();
    const action = attr('action')?.trim();
    if (method === 'post' || !action || action.startsWith('#')) return;
    let url;
    try { url = new URL(action, document.baseURI); } catch { return; }
    if (url.origin !== appOrigin) return;
    const query = new URLSearchParams();
    for (const [key, value] of new FormData(form, submitter)) {
      query.append(key, typeof value === 'string' ? value : value.name);
    }
    url.search = query.toString();
    navigateApp(url);
  });
})();
