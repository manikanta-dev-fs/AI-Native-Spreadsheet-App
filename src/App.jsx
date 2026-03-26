import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import './App.css'
import { createEngine } from './engine/core.js'

const TOTAL_ROWS = 50
const TOTAL_COLS = 50
const STORAGE_KEY = 'spreadsheet_v1'

// ─────────────────────────────────────────────────────────────
//  Task 1 — Filter Dropdown Component
//  Renders per-column filter UI with checkboxes for unique values
// ─────────────────────────────────────────────────────────────

function FilterDropdown({ colIndex, engine, version, filterConfig, setFilterConfig, onClose }) {
  // Collect all unique non-empty computed values in this column
  const allValues = useMemo(() => {
    const vals = new Set()
    for (let r = 0; r < engine.rows; r++) {
      const cell = engine.getCell(r, colIndex)
      const v = cell.error
        ? cell.error
        : (cell.computed !== null && cell.computed !== '' ? String(cell.computed) : cell.raw)
      if (v !== '') vals.add(v)
    }
    return [...vals].sort((a, b) => {
      const na = parseFloat(a), nb = parseFloat(b)
      return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b)
    })
  }, [colIndex, engine, version])

  // If no active filter, treat all values as selected
  const current = filterConfig[colIndex] || new Set(allValues)

  const toggleValue = (val) => {
    const next = new Set(current)
    next.has(val) ? next.delete(val) : next.add(val)
    setFilterConfig(prev => ({ ...prev, [colIndex]: next }))
  }

  const selectAll = () => setFilterConfig(prev => { const n = { ...prev }; delete n[colIndex]; return n })
  const clearAll = () => setFilterConfig(prev => ({ ...prev, [colIndex]: new Set() }))

  return (
    <div className="filter-dropdown" onClick={e => e.stopPropagation()}>
      <div className="filter-dropdown-header">
        <span>Filter</span>
        <button className="filter-close-btn" onClick={onClose}>✕</button>
      </div>
      <div className="filter-actions">
        <button className="filter-action-btn" onClick={selectAll}>Select All</button>
        <button className="filter-action-btn" onClick={clearAll}>Clear</button>
      </div>
      <div className="filter-list">
        {allValues.length === 0 && <div className="filter-empty">No values</div>}
        {allValues.map(val => (
          <label key={val} className="filter-item">
            <input
              type="checkbox"
              checked={current.has(val)}
              onChange={() => toggleValue(val)}
            />
            <span className="filter-item-label">{val}</span>
          </label>
        ))}
      </div>
    </div>
  )
}

// ─────────────────────────────────────────────────────────────
//  Main App
// ─────────────────────────────────────────────────────────────

export default function App() {
  const [engine] = useState(() => createEngine(TOTAL_ROWS, TOTAL_COLS))
  const [version, setVersion] = useState(0)
  const [selectedCell, setSelectedCell] = useState(null)
  const [editingCell, setEditingCell] = useState(null)
  const [editValue, setEditValue] = useState('')
  const [cellStyles, setCellStyles] = useState({})
  const cellInputRef = useRef(null)

  // ── Task 1: Sort & Filter state ──
  const [sortConfig, setSortConfig] = useState(null)       // { col, dir: 'asc' | 'desc' } | null
  const [filterConfig, setFilterConfig] = useState({})     // { colIndex: Set<string> }
  const [openFilter, setOpenFilter] = useState(null)       // colIndex of currently open filter dropdown

  // ── Task 2: Copy/paste state ──
  const [copySource, setCopySource] = useState(null)       // { r, c } of last copied cell

  // ── Task 3: localStorage debounce timer ──
  const saveTimerRef = useRef(null)
  // Guard: prevent auto-save from firing before the initial load completes
  const hasLoadedRef = useRef(false)

  // ── Task 3: Load from localStorage on first mount ──
  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY)
      if (saved) {
        const parsed = JSON.parse(saved)
        // Restore cell data via engine
        if (parsed.cells) engine.loadFromData(parsed)
        // Restore cell styles (formatting)
        if (parsed.styles) setCellStyles(parsed.styles)
        setVersion(v => v + 1)
      }
    } catch (e) {
      // Corrupted data — wipe and start fresh
      console.warn('localStorage data corrupted, clearing:', e)
      localStorage.removeItem(STORAGE_KEY)
    } finally {
      // Mark load as complete so the auto-save effect can proceed
      hasLoadedRef.current = true
    }
  }, [engine])

  // ── Task 3: Auto-save with 500ms debounce ──
  // Triggered on any version or style change (i.e. any cell edit or format change)
  // Undo/redo history is intentionally NOT persisted
  // Guard: skip save if we haven't loaded yet (avoids overwriting saved data on first render)
  useEffect(() => {
    if (!hasLoadedRef.current) return
    clearTimeout(saveTimerRef.current)
    saveTimerRef.current = setTimeout(() => {
      try {
        const data = engine.serialize()
        const toSave = { ...data, styles: cellStyles }
        localStorage.setItem(STORAGE_KEY, JSON.stringify(toSave))
      } catch (e) {
        // Handle storage quota exceeded gracefully
        if (e.name === 'QuotaExceededError') {
          console.warn('localStorage quota exceeded — save skipped')
        }
      }
    }, 500)
    return () => clearTimeout(saveTimerRef.current)
  }, [version, cellStyles, engine])

  // ── Task 2: Global keyboard listener for Ctrl+C / Ctrl+V ──
  useEffect(() => {
    const handleGlobalKeyDown = async (e) => {
      const isMac = navigator.platform.toUpperCase().includes('MAC')
      const ctrl = isMac ? e.metaKey : e.ctrlKey
      if (!ctrl) return

      // Ctrl+C — copy selected cell's computed value to system clipboard
      if (e.key === 'c' && selectedCell) {
        const cellData = engine.getCell(selectedCell.r, selectedCell.c)
        const val = cellData.error
          ? cellData.error
          : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)
        try {
          await navigator.clipboard.writeText(val)
          setCopySource({ r: selectedCell.r, c: selectedCell.c })
        } catch (err) {
          console.warn('Clipboard write failed:', err)
        }
      }

      // Ctrl+V — paste tab/newline separated data (Excel & Google Sheets format)
      // Each paste is a single undoable action (Ctrl+Z restores each changed cell)
      if (e.key === 'v' && selectedCell) {
        try {
          const text = await navigator.clipboard.readText()
          if (!text) return

          // Parse the pasted data: rows split by newline, columns by tab
          const pastedRows = text
            .replace(/\r\n/g, '\n')
            .replace(/\r/g, '\n')
            .trimEnd()
            .split('\n')
            .map(row => row.split('\t'))

          // Apply each pasted cell value via engine.setCell (which pushes to undo stack)
          pastedRows.forEach((row, dr) => {
            row.forEach((val, dc) => {
              const r = selectedCell.r + dr
              const c = selectedCell.c + dc
              if (r < engine.rows && c < engine.cols) {
                engine.setCell(r, c, val.trim())
              }
            })
          })

          forceRerender()
        } catch (err) {
          console.warn('Clipboard read failed:', err)
        }
      }
    }

    window.addEventListener('keydown', handleGlobalKeyDown)
    return () => window.removeEventListener('keydown', handleGlobalKeyDown)
  }, [selectedCell, engine])

  const forceRerender = useCallback(() => setVersion(v => v + 1), [])

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback((row, col) => {
    const key = `${row},${col}`
    return cellStyles[key] || {
      bold: false, italic: false, underline: false,
      bg: 'white', color: '#202124', align: 'left', fontSize: 13
    }
  }, [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    const key = `${row},${col}`
    setCellStyles(prev => ({
      ...prev,
      [key]: { ...getCellStyle(row, col), ...updates }
    }))
  }, [getCellStyle])

  // ────── Cell editing ──────

  const startEditing = useCallback((row, col) => {
    setSelectedCell({ r: row, c: col })
    setEditingCell({ r: row, c: col })
    const cellData = engine.getCell(row, col)
    setEditValue(cellData.raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine])

  const commitEdit = useCallback((row, col) => {
    const currentCell = engine.getCell(row, col)
    if (currentCell.raw !== editValue) {
      engine.setCell(row, col, editValue)
      forceRerender()
    }
    setEditingCell(null)
  }, [engine, editValue, forceRerender])

  const handleCellClick = useCallback((row, col) => {
    // Close any open filter dropdown when clicking a cell
    setOpenFilter(null)
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }
    if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
      startEditing(row, col)
    }
  }, [editingCell, commitEdit, startEditing])

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback((event, row, col) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'Tab') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    } else if (event.key === 'Escape') {
      setEditValue(engine.getCell(row, col).raw)
      setEditingCell(null)
    } else if (event.key === 'ArrowDown') {
      event.preventDefault(); commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault(); commitEdit(row, col)
      startEditing(Math.max(row - 1, 0), col)
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault(); commitEdit(row, col)
      if (col > 0) { startEditing(row, col - 1) }
      else if (row > 0) { startEditing(row - 1, engine.cols - 1) }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault(); commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    }
  }, [engine, commitEdit, startEditing])

  // ────── Formula bar ──────

  const handleFormulaBarKeyDown = useCallback((event) => {
    if (!editingCell) return
    handleKeyDown(event, editingCell.r, editingCell.c)
  }, [editingCell, handleKeyDown])

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell)
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw)
    }
  }, [selectedCell, editingCell, engine])

  const handleFormulaBarChange = useCallback((value) => {
    if (!editingCell && selectedCell) setEditingCell(selectedCell)
    setEditValue(value)
  }, [editingCell, selectedCell])

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => { if (engine.undo()) forceRerender() }, [engine, forceRerender])
  const handleRedo = useCallback(() => { if (engine.redo()) forceRerender() }, [engine, forceRerender])

  // ────── Formatting ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !getCellStyle(selectedCell.r, selectedCell.c).bold })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !getCellStyle(selectedCell.r, selectedCell.c).italic })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { underline: !getCellStyle(selectedCell.r, selectedCell.c).underline })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const changeFontSize = useCallback((size) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size })
  }, [selectedCell, updateCellStyle])

  const changeAlignment = useCallback((align) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { align })
  }, [selectedCell, updateCellStyle])

  const changeFontColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { color })
  }, [selectedCell, updateCellStyle])

  const changeBackgroundColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bg: color })
  }, [selectedCell, updateCellStyle])

  // ────── Clear ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    engine.setCell(selectedCell.r, selectedCell.c, '')
    forceRerender()
    const key = `${selectedCell.r},${selectedCell.c}`
    setCellStyles(prev => { const next = { ...prev }; delete next[key]; return next })
    setEditValue('')
  }, [selectedCell, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) { engine.setCell(r, c, '') }
    }
    forceRerender()
    setCellStyles({})
    setSelectedCell(null); setEditingCell(null); setEditValue('')
    // Task 3: Also clear localStorage when user clears everything
    localStorage.removeItem(STORAGE_KEY)
  }, [engine, forceRerender])

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return
    engine.insertRow(selectedCell.r); forceRerender()
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c })
  }, [selectedCell, engine, forceRerender])

  const deleteRow = useCallback(() => {
    if (!selectedCell) return
    engine.deleteRow(selectedCell.r); forceRerender()
    if (selectedCell.r >= engine.rows) setSelectedCell({ r: engine.rows - 1, c: selectedCell.c })
  }, [selectedCell, engine, forceRerender])

  const insertColumn = useCallback(() => {
    if (!selectedCell) return
    engine.insertColumn(selectedCell.c); forceRerender()
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 })
  }, [selectedCell, engine, forceRerender])

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return
    engine.deleteColumn(selectedCell.c); forceRerender()
    if (selectedCell.c >= engine.cols) setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 })
  }, [selectedCell, engine, forceRerender])

  // ────── Task 1: Sort handler ──────
  // Cycles: none → asc → desc → none (view-layer only, original data untouched)

  const handleColumnSort = useCallback((colIndex) => {
    setSortConfig(prev => {
      if (!prev || prev.col !== colIndex) return { col: colIndex, dir: 'asc' }
      if (prev.dir === 'asc') return { col: colIndex, dir: 'desc' }
      return null  // reset to original order
    })
  }, [])

  // ────── Task 1: Visible rows (sort + filter applied) ──────
  // This is purely a view transformation — engine data is never mutated
  // Formulas continue to reference their original cell positions

  const visibleRows = useMemo(() => {
    let rowIndices = Array.from({ length: engine.rows }, (_, i) => i)

    // Apply column filters — hide rows where the column value isn't in the allowed set
    Object.entries(filterConfig).forEach(([col, allowedSet]) => {
      if (!allowedSet) return
      const colIdx = parseInt(col)
      rowIndices = rowIndices.filter(r => {
        const cell = engine.getCell(r, colIdx)
        const val = cell.error
          ? cell.error
          : (cell.computed !== null && cell.computed !== '' ? String(cell.computed) : cell.raw)
        return allowedSet.has(val)
      })
    })

    // Apply sort on computed values (formula results, not raw formulas)
    if (sortConfig && sortConfig.dir) {
      rowIndices.sort((a, b) => {
        const ca = engine.getCell(a, sortConfig.col)
        const cb = engine.getCell(b, sortConfig.col)
        const va = ca.error ? ca.error : (ca.computed !== null ? ca.computed : ca.raw)
        const vb = cb.error ? cb.error : (cb.computed !== null ? cb.computed : cb.raw)
        const numA = parseFloat(va), numB = parseFloat(vb)
        const bothNumeric = !isNaN(numA) && !isNaN(numB)
        const cmp = bothNumeric ? numA - numB : String(va).localeCompare(String(vb))
        return sortConfig.dir === 'asc' ? cmp : -cmp
      })
    }

    return rowIndices
  }, [sortConfig, filterConfig, engine, version])

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null
  }, [selectedCell, getCellStyle])

  const getColumnLabel = useCallback((col) => {
    let label = '', num = col + 1
    while (num > 0) { num--; label = String.fromCharCode(65 + (num % 26)) + label; num = Math.floor(num / 26) }
    return label
  }, [])

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : 'No cell'

  const formulaBarValue = editingCell
    ? editValue
    : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  // ────── Render ──────

  return (
    <div className="app-wrapper" onClick={() => setOpenFilter(null)}>
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
        {/* Task 3: Visual indicator that data is being auto-saved */}
        <span className="autosave-hint">● Auto-saved</span>
      </div>

      <div className="main-content">

        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? 'active' : ''}`} onClick={toggleBold} title="Bold">B</button>
            <button className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? 'active' : ''}`} onClick={toggleItalic} title="Italic">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={(e) => changeFontSize(parseInt(e.target.value))}>
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left' ? 'active' : ''}`} onClick={() => changeAlignment('left')} title="Align Left">⬤←</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active' : ''}`} onClick={() => changeAlignment('center')} title="Align Center">⬤</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right' ? 'active' : ''}`} onClick={() => changeAlignment('right')} title="Align Right">⬤→</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input type="color" value={selectedCellStyle?.color || '#000000'} onChange={(e) => changeFontColor(e.target.value)} title="Font color"
              style={{ width: '32px', height: '32px', border: '1px solid #dadce0', cursor: 'pointer', borderRadius: '4px' }} />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select className="toolbar-select" value={selectedCellStyle?.bg || 'white'} onChange={(e) => changeBackgroundColor(e.target.value)}>
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo (Ctrl+Z)">↶ Undo</button>
            <button className="toolbar-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo (Ctrl+Y)">↷ Redo</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow}>+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow}>- Row</button>
            <button className="toolbar-btn" onClick={insertColumn}>+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn}>- Col</button>
          </div>

          <div className="toolbar-group">
            {/* Task 1: Clear active sort/filter */}
            <button
              id="clear-sort-filter-btn"
              className={`toolbar-btn ${sortConfig || Object.keys(filterConfig).length > 0 ? 'active' : ''}`}
              onClick={() => { setSortConfig(null); setFilterConfig({}) }}
              title="Clear all sort & filters"
            >
              ↕ Clear Sort/Filter
            </button>
          </div>

          <div className="toolbar-group">
            <button id="clear-cell-btn" className="toolbar-btn danger" onClick={clearSelectedCell} title="Clear selected cell">✕ Cell</button>
            <button id="clear-all-btn" className="toolbar-btn danger" onClick={clearAllCells} title="Clear all cells and reset storage">✕ All</button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll">
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => {
                  const isSorted = sortConfig?.col === colIndex
                  const isFiltered = !!filterConfig[colIndex]
                  return (
                    <th key={colIndex} className={`col-header ${isSorted ? 'col-sorted' : ''} ${isFiltered ? 'col-filtered' : ''}`}>
                      <div className="col-header-inner">
                        {/* Sort toggle — click label to cycle sort */}
                        <span
                          className="col-label"
                          onClick={() => handleColumnSort(colIndex)}
                          title="Click to sort"
                        >
                          {getColumnLabel(colIndex)}
                          {isSorted && (
                            <span className="sort-indicator">
                              {sortConfig.dir === 'asc' ? ' ↑' : ' ↓'}
                            </span>
                          )}
                        </span>
                        {/* Filter toggle — click ▼ to open dropdown */}
                        <span
                          className={`filter-toggle ${isFiltered ? 'filter-active' : ''}`}
                          onClick={(e) => {
                            e.stopPropagation()
                            setOpenFilter(openFilter === colIndex ? null : colIndex)
                          }}
                          title="Filter column"
                        >
                          ▼
                        </span>
                      </div>
                      {/* Task 1: Filter dropdown renders inline below the header */}
                      {openFilter === colIndex && (
                        <FilterDropdown
                          colIndex={colIndex}
                          engine={engine}
                          version={version}
                          filterConfig={filterConfig}
                          setFilterConfig={setFilterConfig}
                          onClose={() => setOpenFilter(null)}
                        />
                      )}
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {/* Task 1: Render visibleRows (sorted + filtered) instead of all rows */}
              {visibleRows.map((rowIndex) => (
                <tr key={rowIndex}>
                  <td className="row-header">{rowIndex + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const isSelected = selectedCell?.r === rowIndex && selectedCell?.c === colIndex
                    const isEditing = editingCell?.r === rowIndex && editingCell?.c === colIndex
                    // Task 2: Highlight the copy source cell
                    const isCopySource = copySource?.r === rowIndex && copySource?.c === colIndex
                    const cellData = engine.getCell(rowIndex, colIndex)
                    const style = cellStyles[`${rowIndex},${colIndex}`] || {}
                    const displayValue = cellData.error
                      ? cellData.error
                      : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected ? 'selected' : ''} ${isCopySource ? 'copy-source' : ''}`}
                        style={{ background: style.bg || 'white' }}
                        onMouseDown={(e) => { e.preventDefault(); handleCellClick(rowIndex, colIndex) }}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(rowIndex, colIndex)}
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: style.color || '#202124',
                              fontSize: (style.fontSize || 13) + 'px',
                              textAlign: style.align || 'left',
                              background: style.bg || 'white',
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: cellData.error ? '#d93025' : (style.color || '#202124'),
                              fontSize: (style.fontSize || 13) + 'px',
                            }}
                          >
                            {displayValue}
                          </div>
                        )}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click column header to sort · ▼ to filter · Ctrl+C/V for copy-paste (Excel/Sheets compatible) · Data auto-saved · Formulas: =SUM(A1:A5) · =AVG · =MAX · =MIN
        </p>
      </div>
    </div>
  )
}
