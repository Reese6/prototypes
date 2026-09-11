import { useSyncExternalStore } from 'react'
import { HtmlFrame } from './HtmlFrame.tsx'
import './App.css'

function subscribe(onChange: () => void) {
  window.addEventListener('popstate', onChange)
  return () => window.removeEventListener('popstate', onChange)
}

function navigate(url: URL) {
  window.history.pushState(null, '', url)
  // pushState не вызывает popstate — оповещаем подписчиков сами.
  window.dispatchEvent(new PopStateEvent('popstate'))
}

function App() {
  const pathname = useSyncExternalStore(subscribe, () => window.location.pathname)

  if (pathname === '/') {
    return <HtmlFrame src="/demo.html" title="Демо HTML" onNavigate={navigate} />
  }

  return (
    <main className="app-page">
      <h1>Страница приложения</h1>
      <p>
        <code>{pathname}</code>
      </p>
      <a href="/">Вернуться к HTML</a>
    </main>
  )
}

export default App
