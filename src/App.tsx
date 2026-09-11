import { BrowserRouter, Link, Route, Routes, useLocation, useNavigate } from 'react-router'
import { HtmlFrame } from './HtmlFrame.tsx'
import './App.css'

function DemoPage() {
  const navigate = useNavigate()
  return <HtmlFrame src="/demo.html" title="Демо HTML" onNavigate={navigate} />
}

function AppPage() {
  const { pathname, search, hash } = useLocation()
  return (
    <main className="app-page">
      <h1>Страница приложения</h1>
      <p>
        <code>{pathname + search + hash}</code>
      </p>
      <Link to="/">Вернуться к HTML</Link>
    </main>
  )
}

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<DemoPage />} />
        <Route path="*" element={<AppPage />} />
      </Routes>
    </BrowserRouter>
  )
}

export default App
