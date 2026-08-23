import { useEffect, useState } from 'react'
import './App.css'

const API_BASE = 'http://localhost:8080'

function App() {
  const [chunksStored, setChunksStored] = useState(null)
  const [ingesting, setIngesting] = useState(false)
  const [ingestMessage, setIngestMessage] = useState('')

  const [question, setQuestion] = useState('How many days of annual leave do employees get?')
  const [asking, setAsking] = useState(false)
  const [directAnswer, setDirectAnswer] = useState(null)
  const [ragAnswer, setRagAnswer] = useState(null)
  const [sources, setSources] = useState([])
  const [error, setError] = useState('')

  const refreshStatus = () => {
    fetch(`${API_BASE}/api/v1/documents/status`)
      .then((res) => res.json())
      .then((data) => setChunksStored(data.chunksStored))
      .catch(() => setChunksStored(null))
  }

  useEffect(() => {
    refreshStatus()
  }, [])

  const handleIngest = async () => {
    setIngesting(true)
    setIngestMessage('')
    try {
      const res = await fetch(`${API_BASE}/api/v1/documents/ingest`, { method: 'POST' })
      if (!res.ok) throw new Error(`Ingest failed: ${res.status}`)
      const data = await res.json()
      setIngestMessage(`Indexed ${data.chunksIndexed} chunk(s) across both documents.`)
      refreshStatus()
    } catch (err) {
      setIngestMessage(err.message)
    } finally {
      setIngesting(false)
    }
  }

  const handleAsk = async (e) => {
    e.preventDefault()
    setAsking(true)
    setError('')
    setDirectAnswer(null)
    setRagAnswer(null)
    setSources([])
    try {
      const res = await fetch(`${API_BASE}/api/v1/rag/compare`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ question }),
      })
      if (!res.ok) throw new Error(`Query failed: ${res.status}`)
      const data = await res.json()
      setDirectAnswer(data.directAnswer)
      setRagAnswer(data.ragAnswer)
      setSources(data.sources ?? [])
    } catch (err) {
      setError(err.message)
    } finally {
      setAsking(false)
    }
  }

  return (
    <div className="page">
      <h1>HR &amp; IT Policy RAG Demo</h1>
      <p className="subtitle">Spring AI + MariaDB vector store — Day 2/3 training</p>

      <section className="card">
        <h2>1. Ingest</h2>
        <p>Chunks currently stored in the vector store: <strong>{chunksStored ?? '—'}</strong></p>
        <p className="hint">Ingests both the HR Company Policy and the IT &amp; Data Security Policy.</p>
        <button onClick={handleIngest} disabled={ingesting}>
          {ingesting ? 'Ingesting…' : 'Ingest Both Policy PDFs'}
        </button>
        {ingestMessage && <p className="message">{ingestMessage}</p>}
      </section>

      <section className="card">
        <h2>2. Ask a question — with vs. without RAG</h2>
        <form onSubmit={handleAsk}>
          <textarea
            rows={3}
            value={question}
            onChange={(e) => setQuestion(e.target.value)}
            placeholder="Ask something about the HR or IT policy..."
          />
          <button type="submit" disabled={asking || !question.trim()}>
            {asking ? 'Asking…' : 'Compare'}
          </button>
        </form>
        {error && <p className="error">{error}</p>}
        {(directAnswer || ragAnswer) && (
          <div className="compare">
            <div className="answer answer-direct">
              <h3>Without RAG (plain LLM)</h3>
              <p>{directAnswer}</p>
            </div>
            <div className="answer answer-rag">
              <h3>With RAG (grounded)</h3>
              <p>{ragAnswer}</p>
              {sources.length > 0 && (
                <details>
                  <summary>{sources.length} source chunk(s) used</summary>
                  <ul>
                    {sources.map((s, i) => (
                      <li key={i}>
                        <span className="score">{s.score.toFixed(2)}</span>
                        <em>{s.source}</em>: {s.excerpt}
                      </li>
                    ))}
                  </ul>
                </details>
              )}
            </div>
          </div>
        )}
      </section>
    </div>
  )
}

export default App
