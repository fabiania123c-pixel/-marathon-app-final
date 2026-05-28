import React, { useState } from 'react'
import * as XLSX from 'xlsx'
import html2canvas from 'html2canvas'
import { jsPDF } from 'jspdf'
import './App.css'

function App() {
  const [people, setPeople] = useState([])
  const [unidades, setUnidades] = useState([])
  const [selectedUnidad, setSelectedUnidad] = useState('')
  const [editingId, setEditingId] = useState(null)
  const [editData, setEditData] = useState({})

  const handleFileUpload = async (e) => {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        const ws = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })

        const allPeople = []
        const uniqueUnidades = new Set()

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i]
          if (!row[0]) continue

          const nombre = String(row[0]).trim()
          const cargo = String(row[1] || '').trim()
          const jefe = String(row[2] || '').trim()
          const unidad = String(row[3] || 'SIN UNIDAD').trim()
          const area = String(row[4] || '').trim()

          allPeople.push({
            id: `${nombre}_${i}`,
            nombre,
            cargo,
            jefe,
            unidad,
            area
          })

          uniqueUnidades.add(unidad)
        }

        setPeople(allPeople)
        setUnidades(Array.from(uniqueUnidades).sort())
        setEditingId(null)
        if (uniqueUnidades.size > 0) {
          setSelectedUnidad(Array.from(uniqueUnidades).sort()[0])
        }
      } catch (err) {
        alert('Error al leer Excel: ' + err.message)
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const getUnidadPeople = () => {
    return people.filter(p => p.unidad === selectedUnidad)
  }

  const buildHierarchy = (unidadPeople) => {
    const hierarchy = new Map()
    
    for (const person of unidadPeople) {
      if (!hierarchy.has(person.nombre)) {
        hierarchy.set(person.nombre, [])
      }
    }

    for (const person of unidadPeople) {
      if (person.jefe) {
        const jefe = unidadPeople.find(p => p.nombre.toUpperCase() === person.jefe.toUpperCase())
        if (jefe) {
          const subs = hierarchy.get(jefe.nombre) || []
          subs.push(person.nombre)
          hierarchy.set(jefe.nombre, subs)
        }
      }
    }

    return hierarchy
  }

  const findRoot = (unidadPeople) => {
    for (const p of unidadPeople) {
      if (!unidadPeople.some(x => x.nombre === p.jefe)) {
        return p
      }
    }
    return unidadPeople[0] || null
  }

  const handleCellClick = (person) => {
    setEditingId(person.id)
    setEditData({ ...person })
  }

  const handleEditChange = (field, value) => {
    setEditData({ ...editData, [field]: value })
  }

  const handleSaveEdit = () => {
    const updatedPeople = people.map(p => p.id === editingId ? editData : p)
    setPeople(updatedPeople)
    setEditingId(null)
  }

  const handleCancel = () => {
    setEditingId(null)
  }

  const renderNode = (person, hierarchy, unidadPeople) => {
    const subordinados = hierarchy.get(person.nombre) || []
    const isEditing = editingId === person.id
    const currentData = isEditing ? editData : person

    return (
      <div className="org-node" key={person.id}>
        <div 
          className={`org-box ${isEditing ? 'editing' : ''}`}
          onClick={() => handleCellClick(person)}
        >
          {isEditing ? (
            <div className="edit-form" onClick={(e) => e.stopPropagation()}>
              <input
                type="text"
                value={currentData.cargo}
                onChange={(e) => handleEditChange('cargo', e.target.value)}
                placeholder="Cargo"
                autoFocus
              />
              <input
                type="text"
                value={currentData.nombre}
                onChange={(e) => handleEditChange('nombre', e.target.value)}
                placeholder="Nombre"
              />
              <input
                type="text"
                value={currentData.jefe}
                onChange={(e) => handleEditChange('jefe', e.target.value)}
                placeholder="Jefe"
              />
              <div className="edit-buttons">
                <button onClick={handleSaveEdit} className="save-btn">✓</button>
                <button onClick={handleCancel} className="cancel-btn">✕</button>
              </div>
            </div>
          ) : (
            <>
              <div className="org-box-cargo">{currentData.cargo}</div>
              <div className="org-box-nombre">{currentData.nombre}</div>
              {currentData.area && <div className="org-box-area">{currentData.area}</div>}
            </>
          )}
        </div>

        {subordinados.length > 0 && (
          <div className="org-children-wrapper">
            <svg className="org-connector" viewBox="0 0 1000 250" preserveAspectRatio="none">
              {/* Línea vertical desde padre - START AQUÍ en y=0 */}
              <line x1="500" y1="0" x2="500" y2="60" stroke="#2563eb" strokeWidth="4" strokeLinecap="round" />
              {/* Línea horizontal conectora */}
              <line x1="0" y1="60" x2="1000" y2="60" stroke="#2563eb" strokeWidth="4" strokeLinecap="round" />
              {/* Líneas verticales a cada subordinado */}
              {subordinados.map((_, idx) => {
                const x = (idx / (subordinados.length > 1 ? subordinados.length - 1 : 1)) * 1000
                return <line key={idx} x1={x} y1="60" x2={x} y2="250" stroke="#2563eb" strokeWidth="4" strokeLinecap="round" />
              })}
            </svg>
            
            <div className="org-children">
              {subordinados.map(subName => {
                const sub = unidadPeople.find(p => p.nombre === subName)
                return sub ? (
                  <div key={sub.id} className="org-child">
                    {renderNode(sub, hierarchy, unidadPeople)}
                  </div>
                ) : null
              })}
            </div>
          </div>
        )}
      </div>
    )
  }

  const exportPNG = async () => {
    const element = document.getElementById('chart')
    if (!element) return

    const canvas = await html2canvas(element, { backgroundColor: '#ffffff', scale: 2 })
    const link = document.createElement('a')
    link.href = canvas.toDataURL()
    link.download = `Marathon_${selectedUnidad}.png`
    link.click()
  }

  const exportPDF = async () => {
    const element = document.getElementById('chart')
    if (!element) return

    const canvas = await html2canvas(element, { backgroundColor: '#ffffff', scale: 2 })
    const imgData = canvas.toDataURL('image/png')
    const pdf = new jsPDF({
      orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
      unit: 'mm',
      format: [canvas.width / 2, canvas.height / 2]
    })

    pdf.addImage(imgData, 'PNG', 0, 0, canvas.width / 2, canvas.height / 2)
    pdf.save(`Marathon_${selectedUnidad}.pdf`)
  }

  const unidadPeople = getUnidadPeople()
  const hierarchy = buildHierarchy(unidadPeople)
  const root = findRoot(unidadPeople)

  return (
    <div className="app">
      <div className="header">
        <h1>🎯 Marathon Organigrama</h1>
        <p>Edita datos directamente en los cuadros</p>
      </div>

      <div className="card upload-card">
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          className="file-input"
        />
      </div>

      {people.length > 0 && (
        <>
          <div className="card controls-card">
            <select
              value={selectedUnidad}
              onChange={(e) => {
                setSelectedUnidad(e.target.value)
                setEditingId(null)
              }}
              className="select"
            >
              {unidades.map(u => (
                <option key={u} value={u}>
                  {u} ({people.filter(p => p.unidad === u).length})
                </option>
              ))}
            </select>
            <button onClick={exportPNG} className="btn">📥 PNG</button>
            <button onClick={exportPDF} className="btn">📄 PDF</button>
          </div>

          <div className="card chart-card">
            <div id="chart" className="chart-container">
              {root ? (
                <div className="org-tree">
                  {renderNode(root, hierarchy, unidadPeople)}
                </div>
              ) : (
                <p>Sin datos</p>
              )}
            </div>
          </div>
        </>
      )}
    </div>
  )
}

export default App
