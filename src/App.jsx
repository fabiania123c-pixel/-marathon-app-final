import React, { useState, useRef } from 'react'
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
  const [positions, setPositions] = useState({})
  const [draggingId, setDraggingId] = useState(null)
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 })
  const containerRef = useRef(null)

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
        setPositions({})
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

  const handleBoxMouseDown = (e, personId) => {
    if (editingId) return
    e.preventDefault()
    e.stopPropagation()
    setDraggingId(personId)
    setDragStart({
      x: e.clientX - (positions[personId]?.x || 0),
      y: e.clientY - (positions[personId]?.y || 0)
    })
  }

  const handleContainerMouseMove = (e) => {
    if (!draggingId || !containerRef.current) return
    
    const rect = containerRef.current.getBoundingClientRect()
    const newX = e.clientX - rect.left - dragStart.x
    const newY = e.clientY - rect.top - dragStart.y

    setPositions(prev => ({
      ...prev,
      [draggingId]: { x: newX, y: newY }
    }))
  }

  const handleContainerMouseUp = () => {
    setDraggingId(null)
  }

  const handleCellClick = (person, e) => {
    if (draggingId) return
    e.stopPropagation()
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
    setEditData({})
  }

  const handleCancel = () => {
    setEditingId(null)
    setEditData({})
  }

  const getBoxPosition = (personId) => {
    return positions[personId] || { x: 0, y: 0 }
  }

  const renderConnectorLine = (fromPos, toPos) => {
    const x1 = fromPos.x + 110
    const y1 = fromPos.y + 100
    const x2 = toPos.x + 110
    const y2 = toPos.y
    
    const midY = (y1 + y2) / 2

    return (
      <svg key={`line-${fromPos.x}-${fromPos.y}-${toPos.x}-${toPos.y}`} className="connector-line" style={{
        position: 'absolute',
        left: Math.min(x1, x2),
        top: y1,
        width: Math.abs(x2 - x1) + 10,
        height: Math.abs(y2 - y1) + 10,
        pointerEvents: 'none',
        overflow: 'visible'
      }}>
        <line
          x1={x1 > x2 ? Math.abs(x2 - x1) : 0}
          y1="0"
          x2={x1 > x2 ? 0 : Math.abs(x2 - x1)}
          y2={Math.abs(y2 - y1)}
          stroke="#2563eb"
          strokeWidth="3"
          fill="none"
          strokeLinecap="round"
        />
      </svg>
    )
  }

  const renderNode = (person, hierarchy, unidadPeople) => {
    const subordinados = hierarchy.get(person.nombre) || []
    const isEditing = editingId === person.id
    const currentData = isEditing ? editData : person
    const pos = getBoxPosition(person.id)

    return (
      <div 
        key={person.id} 
        className="org-node"
        style={{ position: 'relative' }}
      >
        <div
          className={`org-box ${draggingId === person.id ? 'dragging' : ''} ${isEditing ? 'editing' : ''}`}
          onMouseDown={(e) => handleBoxMouseDown(e, person.id)}
          onClick={(e) => handleCellClick(person, e)}
          style={{
            transform: `translate(${pos.x}px, ${pos.y}px)`,
            cursor: 'grab',
            position: 'relative',
            zIndex: draggingId === person.id ? 1000 : 10
          }}
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
              <div className="org-box-title">{currentData.cargo}</div>
              <div className="org-box-name">{currentData.nombre}</div>
              {currentData.area && <div className="org-box-area">{currentData.area}</div>}
            </>
          )}
        </div>

        {subordinados.length > 0 && (
          <div className="org-children" style={{ position: 'relative' }}>
            {subordinados.map(subName => {
              const sub = unidadPeople.find(p => p.nombre === subName)
              if (!sub) return null
              
              const subPos = getBoxPosition(sub.id)
              
              return (
                <div key={sub.id} style={{ position: 'relative' }}>
                  {renderConnectorLine(pos, subPos)}
                  {renderNode(sub, hierarchy, unidadPeople)}
                </div>
              )
            })}
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
        <h1>🎯 Marathon Org Chart</h1>
        <p>Organigrama interactivo • Arrastra y edita</p>
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
                setPositions({})
                setEditingId(null)
              }}
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
            <div 
              id="chart" 
              className="chart-container"
              ref={containerRef}
              onMouseMove={handleContainerMouseMove}
              onMouseUp={handleContainerMouseUp}
              onMouseLeave={handleContainerMouseUp}
            >
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
