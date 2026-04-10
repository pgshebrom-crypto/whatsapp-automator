import React, { useState, useRef, useEffect, useMemo } from 'react'
import * as XLSX from 'xlsx'
import { 
  Upload, 
  FileSpreadsheet, 
  Send, 
  RefreshCw, 
  AlertCircle,
  Search,
  Trash2,
  Eye,
  Bold,
  Italic as ItalicIcon,
  Strikethrough
} from 'lucide-react'
import './index.css'

interface Contact {
  [key: string]: any
}

interface MappedContact extends Contact {
  originalIndex: number
}

interface Mapping {
  name: string
  phone: string
}

function App() {
  const [data, setData] = useState<Contact[]>([])
  const [columns, setColumns] = useState<string[]>([])
  const [mapping, setMapping] = useState<Mapping>({ name: '', phone: '' })
  const [template, setTemplate] = useState('Olá {nome}, tudo bem? Estou entrando em contato para...')
  const [sentIndices, setSentIndices] = useState<Set<number>>(new Set())
  const [error, setError] = useState<string | null>(null)
  const [searchTerm, setSearchTerm] = useState('')
  
  const fileInputRef = useRef<HTMLInputElement>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)

  // Load sent status from localStorage
  useEffect(() => {
    const saved = localStorage.getItem('sent_contacts')
    if (saved) {
      try {
        setSentIndices(new Set(JSON.parse(saved)))
      } catch (e) {
        console.error("Failed to load sent indices", e)
      }
    }
  }, [])

  useEffect(() => {
    localStorage.setItem('sent_contacts', JSON.stringify(Array.from(sentIndices)))
  }, [sentIndices])

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const binaryData = event.target?.result
        const workbook = XLSX.read(binaryData, { type: 'binary' })
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]
        const jsonData = XLSX.utils.sheet_to_json(worksheet) as Contact[]

        if (jsonData.length === 0) {
          setError('A planilha está vazia.')
          return
        }

        const colNames = Object.keys(jsonData[0])
        
        setColumns(colNames)
        setData(jsonData)
        setError(null)
        
        const nameGuess = colNames.find(c => c.toLowerCase().includes('nome')) || ''
        const phoneGuess = colNames.find(c => c.toLowerCase().includes('tel') || c.toLowerCase().includes('cel') || c.toLowerCase().includes('fone')) || ''
        setMapping({ name: nameGuess, phone: phoneGuess })
      } catch (err) {
        setError('Erro ao ler o arquivo. Verifique se é um arquivo Excel ou CSV válido.')
      }
    }
    reader.readAsBinaryString(file)
  }

  const reset = () => {
    if (window.confirm('Tem certeza que deseja carregar um novo arquivo? Todo o progresso atual será perdido.')) {
      setData([])
      setColumns([])
      setMapping({ name: '', phone: '' })
      setSentIndices(new Set())
      setSearchTerm('')
      if (fileInputRef.current) fileInputRef.current.value = ''
    }
  }

  const resetProgress = () => {
    if (window.confirm('Deseja marcar todos os contatos como "Pendente"?')) {
      setSentIndices(new Set())
    }
  }

  const formatWhatsAppNumber = (num: any) => {
    const cleaned = String(num).replace(/\D/g, '')
    if (cleaned.length === 10 || cleaned.length === 11) {
      return `55${cleaned}`
    }
    return cleaned
  }

  const getMessageForContact = (contact: Contact, isEncoded = true) => {
    let msg = template
    
    // Replace any {tag} with column value
    msg = msg.replace(/\{(.+?)\}/g, (match, tag) => {
      // Find the key in contact that matches the tag (case insensitive)
      const key = Object.keys(contact).find(k => k.toLowerCase() === tag.toLowerCase())
      return key ? contact[key] : match
    })

    return isEncoded ? encodeURIComponent(msg) : msg
  }

  const handleSend = (index: number, contact: Contact) => {
    const phone = formatWhatsAppNumber(contact[mapping.phone])
    const message = getMessageForContact(contact)
    
    if (!phone || phone.length < 10) {
      alert('Telefone inválido para este contato. Verifique o mapeamento das colunas.')
      return
    }

    const url = `https://wa.me/${phone}?text=${message}`
    window.open(url, '_blank')
    
    const newSent = new Set(sentIndices)
    newSent.add(index)
    setSentIndices(new Set(newSent))
  }

  const applyFormatting = (prefix: string, suffix: string) => {
    const textarea = textareaRef.current
    if (!textarea) return

    const start = textarea.selectionStart
    const end = textarea.selectionEnd
    const text = template
    
    const selectedText = text.substring(start, end)
    const newText = text.substring(0, start) + prefix + selectedText + suffix + text.substring(end)
    
    setTemplate(newText)
    
    // Restore focus and selection
    setTimeout(() => {
      textarea.focus()
      textarea.setSelectionRange(start + prefix.length, end + prefix.length)
    }, 0)
  }

  const filteredData = useMemo(() => {
    return data.map((item, index) => ({ ...item, originalIndex: index } as MappedContact))
      .filter(contact => {
        const name = String(contact[mapping.name] || '').toLowerCase()
        const phone = String(contact[mapping.phone] || '').toLowerCase()
        const term = searchTerm.toLowerCase()
        return name.includes(term) || phone.includes(term)
      })
  }, [data, searchTerm, mapping])

  const progressPercentage = data.length > 0 ? (sentIndices.size / data.length) * 100 : 0

  return (
    <div className="container">
      <header className="header">
        <h1>WhatsApp Automator</h1>
        <p>Gerenciamento inteligente de mensagens individuais via planilha.</p>
      </header>

      {error && (
        <div className="card" style={{ borderColor: 'var(--error-color)', display: 'flex', alignItems: 'center', gap: '1rem', color: 'var(--error-color)' }}>
          <AlertCircle size={24} />
          <span>{error}</span>
        </div>
      )}

      {data.length === 0 ? (
        <div className="card">
          <div className="upload-area" onClick={() => fileInputRef.current?.click()}>
            <Upload size={48} />
            <h2>Upload da Planilha</h2>
            <p>Clique ou arraste seu arquivo .xlsx ou .csv aqui</p>
            <input 
              type="file" 
              ref={fileInputRef} 
              style={{ display: 'none' }} 
              accept=".xlsx,.xls,.csv"
              onChange={handleFileUpload}
            />
          </div>
        </div>
      ) : (
        <>
          <div className="progress-wrapper">
            <div className="progress-header">
              <span>Progresso dos Envios</span>
              <span>{sentIndices.size} de {data.length} ({Math.round(progressPercentage)}%)</span>
            </div>
            <div className="progress-bg">
              <div 
                className="progress-fill" 
                style={{ width: `${progressPercentage}%` }}
              ></div>
            </div>
          </div>

          <div className="card">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem', flexWrap: 'wrap', gap: '1rem' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <FileSpreadsheet className="text-primary" />
                <h2 style={{ margin: 0 }}>Configuração</h2>
              </div>
              <div style={{ display: 'flex', gap: '0.5rem' }}>
                <button className="btn btn-danger" onClick={resetProgress} title="Resetar status de enviado">
                  <Trash2 size={18} /> Limpar Progresso
                </button>
                <button className="btn btn-secondary" onClick={reset}>
                  <RefreshCw size={18} /> Novo Arquivo
                </button>
              </div>
            </div>

            <div className="mapping-grid">
              <div className="form-group">
                <label>Coluna do Nome (Destinatário)</label>
                <select 
                  value={mapping.name} 
                  onChange={(e) => setMapping({ ...mapping, name: e.target.value })}
                >
                  <option value="">Selecione...</option>
                  {columns.map(col => <option key={col} value={col}>{col}</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Coluna do Telefone</label>
                <select 
                  value={mapping.phone} 
                  onChange={(e) => setMapping({ ...mapping, phone: e.target.value })}
                >
                  <option value="">Selecione...</option>
                  {columns.map(col => <option key={col} value={col}>{col}</option>)}
                </select>
              </div>
            </div>

            <div className="form-group" style={{ marginTop: '1.5rem' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                <label>Template da Mensagem</label>
                <div style={{ display: 'flex', gap: '0.5rem' }}>
                  <button 
                    className="btn btn-secondary" 
                    style={{ padding: '0.25rem 0.5rem' }} 
                    onClick={() => applyFormatting('*', '*')}
                    title="Negrito"
                  >
                    <Bold size={16} />
                  </button>
                  <button 
                    className="btn btn-secondary" 
                    style={{ padding: '0.25rem 0.5rem' }} 
                    onClick={() => applyFormatting('_', '_')}
                    title="Itálico"
                  >
                    <ItalicIcon size={16} />
                  </button>
                  <button 
                    className="btn btn-secondary" 
                    style={{ padding: '0.25rem 0.5rem' }} 
                    onClick={() => applyFormatting('~', '~')}
                    title="Riscado"
                  >
                    <Strikethrough size={16} />
                  </button>
                </div>
              </div>
              <textarea 
                ref={textareaRef}
                value={template}
                onChange={(e) => setTemplate(e.target.value)}
                placeholder="Ex: Olá {Nome}, vimos que você mora em {Cidade}..."
              />
              
              {filteredData.length > 0 && (
                <div className="preview-box">
                  <div className="preview-header">
                    <Eye size={14} /> Prévia da mensagem (Exemplo: {filteredData[0][mapping.name] || 'Contato'})
                  </div>
                  {getMessageForContact(filteredData[0], false)}
                </div>
              )}
            </div>
          </div>

          <div className="card" style={{ marginTop: '2rem' }}>
            <div className="search-container">
              <Search size={20} />
              <input 
                type="text" 
                className="search-input" 
                placeholder="Pesquisar por nome ou telefone..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>

            <div style={{ overflowX: 'auto' }}>
              <table className="contact-table">
                <thead>
                  <tr>
                    <th>Nome</th>
                    <th>Telefone</th>
                    <th style={{ textAlign: 'center' }}>Status</th>
                    <th style={{ textAlign: 'right' }}>Ação</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.length > 0 ? (
                    filteredData.map((contact) => (
                      <tr key={contact.originalIndex}>
                        <td>{contact[mapping.name] || <span style={{ color: 'var(--error-color)' }}>Nulo</span>}</td>
                        <td>{contact[mapping.phone] || '-'}</td>
                        <td style={{ textAlign: 'center' }}>
                          {sentIndices.has(contact.originalIndex) ? (
                            <span className="status-badge status-sent">Enviado</span>
                          ) : (
                            <span className="status-badge status-pending">Pendente</span>
                          )}
                        </td>
                        <td style={{ textAlign: 'right' }}>
                          <button 
                            className={`btn ${sentIndices.has(contact.originalIndex) ? 'btn-secondary' : ''}`}
                            onClick={() => handleSend(contact.originalIndex, contact)}
                            disabled={!mapping.name || !mapping.phone}
                          >
                            {sentIndices.has(contact.originalIndex) ? <RefreshCw size={16} /> : <Send size={16} />}
                            {sentIndices.has(contact.originalIndex) ? 'Reenviar' : 'Enviar'}
                          </button>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={4} style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-secondary)' }}>
                        Nenhum contato encontrado para "{searchTerm}"
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  )
}

export default App
