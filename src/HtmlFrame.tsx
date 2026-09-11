import { useEffect, useEffectEvent, useRef, useState } from 'react'
import bridgeScript from './bridge.js?raw'
import downloadGuardScript from './script.js?raw'
import './HtmlFrame.css'

// Нет allow-same-origin: у документа непрозрачный origin, он не видит приложение
// (DOM, cookie, localStorage) и не может снять sandbox или обойти script.js,
// скачивая файлы через окно родителя. Нет allow-top-navigation: приложение
// навигирует только по сообщениям bridge.js.
const SANDBOX = [
  'allow-scripts',
  'allow-downloads', // фильтрует script.js
  'allow-popups', // ссылки в новой вкладке; без него Chrome блокирует mailto: и tel:
  'allow-popups-to-escape-sandbox', // новая вкладка — обычная страница без sandbox
  'allow-forms', // без него не срабатывает событие submit
  'allow-modals', // alert, confirm, prompt
].join(' ')

type Loaded = { src: string; srcDoc: string } | { src: string; error: string }

type HtmlFrameProps = {
  /** URL HTML-документа; от него разрешаются относительные ресурсы документа. */
  src: string
  title: string
  /**
   * Переход по ссылке документа на origin приложения. Путь приходит строкой
   * "/about?q=1#top" — её принимает navigate из react-router. Объект URL не подходит:
   * navigate копирует его через spread и теряет pathname, search и hash из прототипа.
   */
  onNavigate: (to: string) => void
}

export function HtmlFrame({ src, title, onNavigate }: HtmlFrameProps) {
  const frameRef = useRef<HTMLIFrameElement>(null)
  const [loaded, setLoaded] = useState<Loaded>()
  const navigate = useEffectEvent(onNavigate)

  useEffect(() => {
    const controller = new AbortController()
    async function load() {
      const response = await fetch(src, { signal: controller.signal })
      if (!response.ok) throw new Error(`HTTP ${response.status}`)
      return buildSrcDoc(await response.text(), response.url)
    }
    load()
      .then(
        (srcDoc): Loaded => ({ src, srcDoc }),
        (error: unknown): Loaded => ({ src, error: String(error) }),
      )
      .then((result) => {
        if (!controller.signal.aborted) setLoaded(result)
      })
    return () => controller.abort()
  }, [src])

  useEffect(() => {
    function handleMessage(event: MessageEvent) {
      // origin документа в sandbox — "null", поэтому отправителя сверяем по окну iframe.
      if (event.source !== frameRef.current?.contentWindow) return
      if (event.data?.type !== 'html-frame:navigate' || typeof event.data.href !== 'string') return
      const url = URL.parse(event.data.href)
      if (url?.origin === window.location.origin) navigate(url.pathname + url.search + url.hash)
    }
    window.addEventListener('message', handleMessage)
    return () => window.removeEventListener('message', handleMessage)
  }, [])

  if (loaded?.src !== src) return null
  if ('error' in loaded) {
    return (
      <p role="alert" className="html-frame-error">
        Не удалось загрузить {src}: {loaded.error}
      </p>
    )
  }
  return (
    <iframe
      ref={frameRef}
      className="html-frame"
      title={title}
      srcDoc={loaded.srcDoc}
      sandbox={SANDBOX}
      allow="fullscreen"
    />
  )
}

/** Собирает srcdoc: script.js и bridge.js выполняются раньше скриптов документа. */
function buildSrcDoc(html: string, documentUrl: string) {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  const injected: HTMLElement[] = []
  // Без <base> относительные URL в srcdoc разрешались бы от адреса приложения, а не документа.
  if (!doc.querySelector('base[href]')) {
    const base = doc.createElement('base')
    base.setAttribute('href', documentUrl)
    injected.push(base)
  }
  const bridge = inlineScript(doc, bridgeScript)
  bridge.dataset.appOrigin = window.location.origin
  injected.push(inlineScript(doc, downloadGuardScript), bridge)
  doc.head.prepend(...injected)
  const doctype = doc.doctype ? new XMLSerializer().serializeToString(doc.doctype) : ''
  return doctype + doc.documentElement.outerHTML
}

function inlineScript(doc: Document, code: string) {
  const script = doc.createElement('script')
  // "</script" внутри кода закрыл бы тег при разборе srcdoc.
  script.textContent = code.replace(/<\/(script)/gi, '<\\/$1')
  return script
}
