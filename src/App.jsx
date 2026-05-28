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

  const renderNode = (person, hierarchy, unidadPeople, depth = 0) => {
    const subordinados = hierarchy.get(person.nombre) || []
    const isEditing = editingId === person.id
    const currentData = isEditing ? editData : person

    return (
      <div key={person.id} className="org-node">
        <div 
          className="org-box"
          onClick={() => handleCellClick(person)}
          style={{ cursor: 'pointer' }}
        >
          {isEditing ? (
            <div className="edit-form">
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
                <button onClick={handleSaveEdit} className="save-btn">✓ Guardar</button>
                <button onClick={handleCancel} className="cancel-btn">✕ Cancelar</button>
              </div>
            </div>
          ) : (
            <>
              <div className="org-box-title">{currentData.cargo}</div>
              <div className="org-box-name">{currentData.nombre}</div>
              {currentData.area && <div className="org-box-area">{currentData.area}</div>}
            </>
          )}
        </div>

        {subordinados.length > 0 && (
          <div className="org-children">
            {subordinados.length > 1 && (
              <svg className="org-lines" viewBox="0 0 200 60" preserveAspectRatio="none">
                <line x1="100" y1="0" x2="100" y2="20" stroke="#d1d5db" strokeWidth="2" />
                <line x1="0" y1="20" x2="200" y2="20" stroke="#d1d5db" strokeWidth="2" />
                {subordinados.map((_, i) => {
                  const x = (i / (subordinados.length - 1)) * 200
                  return <line key={i} x1={x} y1="20" x2={x} y2="60" stroke="#d1d5db" strokeWidth="2" />
                })}
              </svg>
            )}
            <div className="org-children-container">
              {subordinados.map((subName, idx) => {
                const sub = unidadPeople.find(p => p.nombre === subName)
                return sub ? (
                  <div key={sub.id} className="org-child-wrapper">
                    {renderNode(sub, hierarchy, unidadPeople, depth + 1)}
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
        <h1>🎯 Marathon Org Chart - Editable</h1>
        <p>Selecciona una Unidad Organizativa y edita los organigramas</p>
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
              onChange={(e) => setSelectedUnidad(e.target.value)}
              className="select"
            >
              {unidades.map(u => (
                <option key={u} value={u}>
                  {u} ({people.filter(p => p.unidad === u).length} personas)
                </option>
              ))}
            </select>
            <button onClick={exportPNG} className="btn">📥 PNG</button>
            <button onClick={exportPDF} className="btn">📄 PDF</button>
          </div>

          <div className="card chart-card">
            <div id="chart" className="chart-container">
              {root ? renderNode(root, hierarchy, unidadPeople) : <p>Sin datos</p>}
            </div>
          </div>
        </>
      )}
    </div>
  )
}

export default App
