// ─────────────────────────────────────────────────────────────
//  Operator precedence for expression parsing
// ─────────────────────────────────────────────────────────────

const OPERATOR_PRECEDENCE = { '+': 1, '-': 1, '*': 2, '/': 2 }
const MAX_UNDO_HISTORY = 100

// ─────────────────────────────────────────────────────────────
//  Column / Cell coordinate utilities
// ─────────────────────────────────────────────────────────────

function columnLetterToIndex(columnLetter) {
    let index = 0
    for (let i = 0; i < columnLetter.length; i++) {
        index = index * 26 + (columnLetter.charCodeAt(i) - 64)
    }
    return index - 1
}

function indexToCellKey(row, col) {
    let columnLabel = ''
    let colNumber = col + 1
    while (colNumber > 0) {
        colNumber--
        columnLabel = String.fromCharCode(65 + (colNumber % 26)) + columnLabel
        colNumber = Math.floor(colNumber / 26)
    }
    return columnLabel + (row + 1)
}

function parseCellKey(cellKey) {
    const match = cellKey.match(/^([A-Z]+)(\d+)$/)
    if (!match) return null
    return { row: parseInt(match[2]) - 1, col: columnLetterToIndex(match[1]) }
}

// ─────────────────────────────────────────────────────────────
//  Dependency graph (tracks which cells depend on which)
// ─────────────────────────────────────────────────────────────

function createDependencyGraph() {
    const forwardEdges = new Map()
    const reverseEdges = new Map()

    function addDependency(cell, dependsOn) {
        if (!forwardEdges.has(cell)) forwardEdges.set(cell, new Set())
        if (!reverseEdges.has(dependsOn)) reverseEdges.set(dependsOn, new Set())
        forwardEdges.get(cell).add(dependsOn)
        reverseEdges.get(dependsOn).add(cell)
    }

    function removeAllDependencies(cell) {
        const deps = forwardEdges.get(cell)
        if (deps) {
            for (const dep of deps) {
                if (reverseEdges.has(dep)) {
                    reverseEdges.get(dep).delete(cell)
                    if (reverseEdges.get(dep).size === 0) reverseEdges.delete(dep)
                }
            }
            forwardEdges.delete(cell)
        }
        const dependents = reverseEdges.get(cell)
        if (dependents) {
            for (const dependent of dependents) {
                if (forwardEdges.has(dependent)) {
                    forwardEdges.get(dependent).delete(cell)
                    if (forwardEdges.get(dependent).size === 0) forwardEdges.delete(dependent)
                }
            }
            reverseEdges.delete(cell)
        }
    }

    function getDependencies(cell) {
        return forwardEdges.get(cell) || new Set()
    }

    function getDependents(cell) {
        return reverseEdges.get(cell) || new Set()
    }

    function getAllDependentsRecursive(cell, visited = new Set()) {
        if (visited.has(cell)) return new Set()
        visited.add(cell)
        const result = new Set()
        for (const dependent of getDependents(cell)) {
            result.add(dependent)
            for (const indirect of getAllDependentsRecursive(dependent, visited)) {
                result.add(indirect)
            }
        }
        return result
    }

    function hasCycle(startCell) {
        const visited = new Set()
        const recursionStack = new Set()
        const dfs = (node) => {
            if (recursionStack.has(node)) return true
            if (visited.has(node)) return false
            visited.add(node)
            recursionStack.add(node)
            for (const dep of getDependencies(node)) {
                if (dfs(dep)) return true
            }
            recursionStack.delete(node)
            return false
        }
        return dfs(startCell)
    }

    function topologicalSort(affectedCells) {
        const inDegree = new Map()
        const queue = []
        const sorted = []

        for (const cell of affectedCells) inDegree.set(cell, 0)
        for (const cell of affectedCells) {
            for (const dep of getDependencies(cell)) {
                if (affectedCells.has(dep)) {
                    inDegree.set(cell, (inDegree.get(cell) || 0) + 1)
                }
            }
        }
        for (const cell of affectedCells) {
            if (inDegree.get(cell) === 0) queue.push(cell)
        }
        while (queue.length > 0) {
            const cell = queue.shift()
            sorted.push(cell)
            for (const dependent of getDependents(cell)) {
                if (affectedCells.has(dependent)) {
                    inDegree.set(dependent, inDegree.get(dependent) - 1)
                    if (inDegree.get(dependent) === 0) queue.push(dependent)
                }
            }
        }
        return sorted
    }

    function clear() {
        forwardEdges.clear()
        reverseEdges.clear()
    }

    return {
        addDependency,
        removeAllDependencies,
        getDependencies,
        getDependents,
        getAllDependentsRecursive,
        hasCycle,
        topologicalSort,
        clear
    }
}

// ─────────────────────────────────────────────────────────────
//  Tokenizer — converts formula string into tokens
// ─────────────────────────────────────────────────────────────

function tokenize(expression) {
    const tokens = []
    let position = 0

    while (position < expression.length) {
        const char = expression[position]

        if (char === ' ') { position++; continue }
        if (char === '(' || char === ')') { tokens.push({ type: 'paren', value: char }); position++; continue }
        if ('+-*/,'.includes(char)) { tokens.push({ type: 'operator', value: char }); position++; continue }
        if (char === ':') { tokens.push({ type: 'range', value: ':' }); position++; continue }

        if ((char >= '0' && char <= '9') || char === '.') {
            let numberStr = ''
            while (position < expression.length && ((expression[position] >= '0' && expression[position] <= '9') || expression[position] === '.')) {
                numberStr += expression[position]; position++
            }
            tokens.push({ type: 'number', value: parseFloat(numberStr) }); continue
        }

        if ((char >= 'A' && char <= 'Z') || (char >= 'a' && char <= 'z')) {
            let identifier = ''
            while (position < expression.length && ((expression[position] >= 'A' && expression[position] <= 'Z') || (expression[position] >= 'a' && expression[position] <= 'z') || (expression[position] >= '0' && expression[position] <= '9'))) {
                identifier += expression[position].toUpperCase(); position++
            }
            if (position < expression.length && expression[position] === ':') {
                tokens.push({ type: 'cell', value: identifier })
            } else {
                const cellMatch = identifier.match(/^([A-Z]+)(\d+)$/)
                if (cellMatch) { tokens.push({ type: 'cell', value: identifier }) }
                else { tokens.push({ type: 'function', value: identifier }) }
            }
            continue
        }

        throw new Error(`Unexpected character: ${char}`)
    }

    return tokens
}

// ─────────────────────────────────────────────────────────────
//  Parser — converts tokens into an AST using shunting-yard
// ─────────────────────────────────────────────────────────────

function parseTokensToAST(tokens) {
    const outputQueue = []
    const operatorStack = []
    let position = 0

    while (position < tokens.length) {
        const token = tokens[position]

        if (token.type === 'number' || token.type === 'cell') {
            outputQueue.push(token); position++
        } else if (token.type === 'function') {
            operatorStack.push(token); position++
        } else if (token.type === 'operator') {
            while (operatorStack.length > 0 && operatorStack[operatorStack.length - 1].type === 'operator' && OPERATOR_PRECEDENCE[operatorStack[operatorStack.length - 1].value] >= OPERATOR_PRECEDENCE[token.value]) {
                outputQueue.push(operatorStack.pop())
            }
            operatorStack.push(token); position++
        } else if (token.type === 'paren' && token.value === '(') {
            operatorStack.push(token); position++
        } else if (token.type === 'paren' && token.value === ')') {
            while (operatorStack.length > 0 && operatorStack[operatorStack.length - 1].value !== '(') {
                outputQueue.push(operatorStack.pop())
            }
            if (operatorStack.length > 0 && operatorStack[operatorStack.length - 1].value === '(') operatorStack.pop()
            if (operatorStack.length > 0 && operatorStack[operatorStack.length - 1].type === 'function') {
                const funcToken = operatorStack.pop()
                if (outputQueue.length > 0 && outputQueue[outputQueue.length - 1].type === 'range') {
                    const rangeToken = outputQueue.pop()
                    outputQueue.push({ type: 'function', name: funcToken.value, range: rangeToken })
                } else { outputQueue.push(funcToken) }
            }
            position++
        } else if (token.type === 'range') {
            if (outputQueue.length < 2) throw new Error('Invalid range syntax')
            const endCell = outputQueue.pop()
            const startCell = outputQueue.pop()
            if (startCell.type !== 'cell' || endCell.type !== 'cell') throw new Error('Range must be between two cell references')
            outputQueue.push({ type: 'range', start: startCell.value, end: endCell.value }); position++
        } else { position++ }
    }

    while (operatorStack.length > 0) {
        if (operatorStack[operatorStack.length - 1].value === '(') throw new Error('Mismatched parentheses')
        outputQueue.push(operatorStack.pop())
    }

    return buildASTFromRPN(outputQueue)
}

function buildASTFromRPN(rpnTokens) {
    const stack = []
    for (const token of rpnTokens) {
        if (token.type === 'number' || token.type === 'cell' || token.type === 'range') {
            stack.push(token)
        } else if (token.type === 'operator') {
            if (stack.length < 2) throw new Error('Invalid expression')
            const right = stack.pop(); const left = stack.pop()
            stack.push({ type: 'binary', operator: token.value, left, right })
        } else if (token.type === 'function') {
            if (token.range) { stack.push(token) }
            else {
                if (stack.length === 0) throw new Error(`Function ${token.value} requires arguments`)
                const argument = stack.pop()
                if (argument.type === 'range') { stack.push({ type: 'function', name: token.value, range: argument }) }
                else { throw new Error(`Function ${token.value} requires a range`) }
            }
        }
    }
    if (stack.length !== 1) throw new Error('Invalid expression')
    return stack[0]
}

// ─────────────────────────────────────────────────────────────
//  Evaluator — walks the AST and computes values
// ─────────────────────────────────────────────────────────────

function expandCellRange(startKey, endKey) {
    const startMatch = startKey.match(/^([A-Z]+)(\d+)$/)
    const endMatch = endKey.match(/^([A-Z]+)(\d+)$/)
    if (!startMatch || !endMatch) throw new Error('REF')
    const startCol = columnLetterToIndex(startMatch[1]); const startRow = parseInt(startMatch[2]) - 1
    const endCol = columnLetterToIndex(endMatch[1]); const endRow = parseInt(endMatch[2]) - 1
    const cells = []
    for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
        for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
            cells.push(indexToCellKey(r, c))
        }
    }
    return cells
}

function evaluateAST(ast, getCellValue, visited = new Set()) {
    if (!ast) return 0
    if (ast.type === 'number') return ast.value
    if (ast.type === 'cell') {
        const key = ast.value
        if (visited.has(key)) throw new Error('CIRCULAR')
        visited.add(key)
        const value = getCellValue(key, visited)
        visited.delete(key)
        return value
    }
    if (ast.type === 'range') throw new Error('VALUE')
    if (ast.type === 'binary') {
        const leftVal = evaluateAST(ast.left, getCellValue, visited)
        const rightVal = evaluateAST(ast.right, getCellValue, visited)
        if (typeof leftVal !== 'number' || typeof rightVal !== 'number') throw new Error('VALUE')
        switch (ast.operator) {
            case '+': return leftVal + rightVal
            case '-': return leftVal - rightVal
            case '*': return leftVal * rightVal
            case '/': if (rightVal === 0) throw new Error('VALUE'); return leftVal / rightVal
            default: throw new Error('VALUE')
        }
    }
    if (ast.type === 'function') {
        const rangeCells = expandCellRange(ast.range.start, ast.range.end)
        const values = rangeCells.map(cellKey => {
            if (visited.has(cellKey)) throw new Error('CIRCULAR')
            visited.add(cellKey)
            const value = getCellValue(cellKey, visited)
            visited.delete(cellKey)
            const num = parseFloat(value)
            return isNaN(num) ? 0 : num
        })
        switch (ast.name) {
            case 'SUM': return values.reduce((sum, val) => sum + val, 0)
            case 'AVG': return values.length === 0 ? 0 : values.reduce((sum, val) => sum + val, 0) / values.length
            case 'MIN': return values.length === 0 ? 0 : Math.min(...values)
            case 'MAX': return values.length === 0 ? 0 : Math.max(...values)
            default: throw new Error('VALUE')
        }
    }
    throw new Error('VALUE')
}

function evaluateFormula(expression, getCellValue) {
    try {
        if (!expression || expression.trim() === '') return { value: '', error: null }
        if (!expression.startsWith('=')) return { value: expression, error: null }
        const formulaBody = expression.slice(1).trim()
        if (formulaBody === '') return { value: '', error: null }
        const tokens = tokenize(formulaBody)
        const ast = parseTokensToAST(tokens)
        const result = evaluateAST(ast, getCellValue)
        return { value: result, error: null }
    } catch (error) {
        if (error.message === 'CIRCULAR') return { value: null, error: '#CYCLE!' }
        if (error.message === 'REF') return { value: null, error: '#REF!' }
        if (error.message === 'VALUE') return { value: null, error: '#VALUE!' }
        return { value: null, error: '#PARSE!' }
    }
}

// ─────────────────────────────────────────────────────────────
//  Reference shifting (for insert/delete row/col)
// ─────────────────────────────────────────────────────────────

function shiftCellReferences(formula, rowShift, colShift, atIndex, isColumnOperation = false) {
    if (!formula || !formula.startsWith('=')) return formula
    return formula.replace(/([A-Z]+)(\d+)/g, (match, colLetter, rowStr) => {
        const colIdx = columnLetterToIndex(colLetter)
        const rowIdx = parseInt(rowStr) - 1
        let newRow = rowIdx; let newCol = colIdx
        if (isColumnOperation) { if (colIdx >= atIndex) newCol = colIdx + colShift }
        else { if (rowIdx >= atIndex) newRow = rowIdx + rowShift }
        if (newRow < 0 || newCol < 0) return match
        return indexToCellKey(newRow, newCol)
    })
}

function extractCellReferences(formula) {
    const references = new Set()
    const regex = /([A-Z]+\d+)/g
    let match
    while ((match = regex.exec(formula)) !== null) { references.add(match[1]) }
    return references
}

// ─────────────────────────────────────────────────────────────
//  Main engine factory
// ─────────────────────────────────────────────────────────────

export function createEngine(initialRows = 50, initialCols = 50) {
    let rows = initialRows
    let cols = initialCols
    const cells = new Map()
    const graph = createDependencyGraph()
    const dirtyCells = new Set()
    const computedCache = new Map()
    const undoStack = []
    const redoStack = []
    let _generation = 0
    let _batchMode = false
    let _batchCallbacks = []

    function cellKey(row, col) { return indexToCellKey(row, col) }
    function cellCoords(key) { return parseCellKey(key) }
    function isValidCellKey(key) {
        const coords = cellCoords(key)
        return coords && coords.row >= 0 && coords.row < rows && coords.col >= 0 && coords.col < cols
    }

    function updateDependencies(cell, rawFormula) {
        for (const ref of extractCellReferences(rawFormula)) {
            if (isValidCellKey(ref)) graph.addDependency(cell, ref)
        }
    }

    function markCellDirty(cell) {
        dirtyCells.add(cell)
        for (const dependent of graph.getAllDependentsRecursive(cell)) { dirtyCells.add(dependent) }
    }

    function markAllCellsDirty() { for (const key of cells.keys()) dirtyCells.add(key) }

    function resolveCellValue(cellKey, visited = new Set()) {
        if (visited.has(cellKey)) return '#CYCLE!'
        visited.add(cellKey)
        if (computedCache.has(cellKey)) {
            const cached = computedCache.get(cellKey)
            if (cached._gen === _generation && !dirtyCells.has(cellKey)) {
                visited.delete(cellKey)
                return cached.error || cached.computed
            }
            computedCache.delete(cellKey)
        }
        const cell = cells.get(cellKey)
        if (!cell || !cell.raw) { visited.delete(cellKey); return '' }
        if (!cell.raw.startsWith('=')) {
            const numericValue = parseFloat(cell.raw)
            visited.delete(cellKey)
            return isNaN(numericValue) ? cell.raw : numericValue
        }
        const result = evaluateFormula(cell.raw, (key) => resolveCellValue(key, visited))
        visited.delete(cellKey)
        if (result.error) {
            computedCache.set(cellKey, { computed: null, error: result.error, _gen: _generation })
            return result.error
        }
        computedCache.set(cellKey, { computed: result.value, error: null, _gen: _generation })
        return result.value
    }

    function recalculate() {
        if (dirtyCells.size === 0) return
        const affected = new Set(dirtyCells)
        const sorted = graph.topologicalSort(affected)
        for (const cell of sorted) {
            if (dirtyCells.has(cell)) { computedCache.delete(cell); resolveCellValue(cell) }
        }
        for (const cell of affected) {
            if (!sorted.includes(cell)) { computedCache.delete(cell); resolveCellValue(cell) }
        }
        dirtyCells.clear()
        if (_batchMode) { _batchCallbacks.forEach(cb => cb()) }
    }

    function setCellRaw(row, col, raw) {
        const key = cellKey(row, col)
        graph.removeAllDependencies(key)
        computedCache.delete(key)
        if (!raw || raw.trim() === '') { cells.delete(key); markCellDirty(key); return }
        if (raw.startsWith('=')) {
            cells.set(key, { raw, computed: null, error: null })
            updateDependencies(key, raw)
            if (graph.hasCycle(key)) {
                graph.removeAllDependencies(key)
                cells.set(key, { raw, computed: null, error: '#CYCLE!' })
                markCellDirty(key)
                return
            }
        } else { cells.set(key, { raw, computed: null, error: null }) }
        markCellDirty(key)
    }

    function getCellRaw(row, col) {
        const key = cellKey(row, col)
        const cell = cells.get(key)
        if (!cell) return { raw: '', computed: null, error: null }
        return { ...cell }
    }

    function takeSnapshot() {
        const snapshot = new Map()
        for (const [key, value] of cells.entries()) { snapshot.set(key, { ...value }) }
        return snapshot
    }

    function restoreSnapshot(snapshot) {
        cells.clear()
        for (const [key, value] of snapshot.entries()) { cells.set(key, { ...value }) }
        graph.clear(); computedCache.clear(); dirtyCells.clear(); _generation++
        for (const [key, value] of cells.entries()) {
            if (value.raw && value.raw.startsWith('=')) updateDependencies(key, value.raw)
        }
        markAllCellsDirty()
    }

    function performInsertRow(atIndex) {
        const snapshot = takeSnapshot()
        cells.clear(); graph.clear(); computedCache.clear(); dirtyCells.clear(); _generation++
        for (const [key, value] of snapshot.entries()) {
            const coords = cellCoords(key); if (!coords) continue
            if (coords.row >= atIndex) { const newKey = cellKey(coords.row + 1, coords.col); cells.set(newKey, { ...value, raw: shiftCellReferences(value.raw, 1, 0, atIndex) }) }
            else { cells.set(key, { ...value, raw: shiftCellReferences(value.raw, 1, 0, atIndex) }) }
        }
        rows++
        for (const [key, value] of cells.entries()) { if (value.raw.startsWith('=')) updateDependencies(key, value.raw) }
        markAllCellsDirty()
    }

    function performDeleteRow(atIndex) {
        const snapshot = takeSnapshot()
        cells.clear(); graph.clear(); computedCache.clear(); dirtyCells.clear(); _generation++
        for (const [key, value] of snapshot.entries()) {
            const coords = cellCoords(key); if (!coords) continue
            if (coords.row === atIndex) continue
            if (coords.row > atIndex) { const newKey = cellKey(coords.row - 1, coords.col); cells.set(newKey, { ...value, raw: shiftCellReferences(value.raw, -1, 0, atIndex) }) }
            else { cells.set(key, { ...value, raw: shiftCellReferences(value.raw, -1, 0, atIndex) }) }
        }
        rows--
        for (const [key, value] of cells.entries()) { if (value.raw.startsWith('=')) updateDependencies(key, value.raw) }
        markAllCellsDirty()
    }

    function performInsertColumn(atIndex) {
        const snapshot = takeSnapshot()
        cells.clear(); graph.clear(); computedCache.clear(); dirtyCells.clear(); _generation++
        for (const [key, value] of snapshot.entries()) {
            const coords = cellCoords(key); if (!coords) continue
            if (coords.col >= atIndex) { const newKey = cellKey(coords.row, coords.col + 1); cells.set(newKey, { ...value, raw: shiftCellReferences(value.raw, 0, 1, atIndex, true) }) }
            else { cells.set(key, { ...value, raw: shiftCellReferences(value.raw, 0, 1, atIndex, true) }) }
        }
        cols++
        for (const [key, value] of cells.entries()) { if (value.raw.startsWith('=')) updateDependencies(key, value.raw) }
        markAllCellsDirty()
    }

    function performDeleteColumn(atIndex) {
        const snapshot = takeSnapshot()
        cells.clear(); graph.clear(); computedCache.clear(); dirtyCells.clear(); _generation++
        for (const [key, value] of snapshot.entries()) {
            const coords = cellCoords(key); if (!coords) continue
            if (coords.col === atIndex) continue
            if (coords.col > atIndex) { const newKey = cellKey(coords.row, coords.col - 1); cells.set(newKey, { ...value, raw: shiftCellReferences(value.raw, 0, -1, atIndex, true) }) }
            else { cells.set(key, { ...value, raw: shiftCellReferences(value.raw, 0, -1, atIndex, true) }) }
        }
        cols--
        for (const [key, value] of cells.entries()) { if (value.raw.startsWith('=')) updateDependencies(key, value.raw) }
        markAllCellsDirty()
    }

    function pushToUndoStack(entry) {
        undoStack.push({ ...entry, _gen: _generation })
        if (undoStack.length > MAX_UNDO_HISTORY) undoStack.shift()
        redoStack.length = 0
    }

    function executeSetCell(row, col, value) {
        const previousValue = getCellRaw(row, col).raw
        if (previousValue !== value) { pushToUndoStack({ type: 'set', r: row, c: col, oldVal: previousValue, newVal: value }) }
        setCellRaw(row, col, value); _generation++; recalculate()
    }

    function executeInsertRow(atIndex) {
        const snapshot = takeSnapshot(); const previousRows = rows
        pushToUndoStack({ type: 'rowins', snap: snapshot, oldRows: previousRows })
        performInsertRow(atIndex); recalculate()
    }

    function executeDeleteRow(atIndex) {
        const snapshot = takeSnapshot(); const previousRows = rows
        pushToUndoStack({ type: 'rowdel', snap: snapshot, oldRows: previousRows })
        performDeleteRow(atIndex); recalculate()
    }

    function executeInsertColumn(atIndex) {
        const snapshot = takeSnapshot(); const previousCols = cols
        pushToUndoStack({ type: 'colins', snap: snapshot, oldCols: previousCols })
        performInsertColumn(atIndex); recalculate()
    }

    function executeDeleteColumn(atIndex) {
        const snapshot = takeSnapshot(); const previousCols = cols
        pushToUndoStack({ type: 'coldel', snap: snapshot, oldCols: previousCols })
        performDeleteColumn(atIndex); recalculate()
    }

    function undo() {
        if (undoStack.length === 0) return false
        const entry = undoStack.pop()
        if (entry.type === 'set') {
            const currentValue = getCellRaw(entry.r, entry.c).raw
            redoStack.push({ ...entry, newVal: currentValue })
            setCellRaw(entry.r, entry.c, entry.oldVal); _generation++; recalculate()
        } else {
            redoStack.push({ ...entry, restoreSnap: takeSnapshot(), restoreRows: rows, restoreCols: cols })
            restoreSnapshot(entry.snap)
            if (entry.type === 'rowins' || entry.type === 'rowdel') rows = entry.oldRows
            else if (entry.type === 'colins' || entry.type === 'coldel') cols = entry.oldCols
            recalculate()
        }
        return true
    }

    function redo() {
        if (redoStack.length === 0) return false
        const entry = redoStack.pop()
        if (entry.type === 'set') {
            const currentValue = getCellRaw(entry.r, entry.c).raw
            undoStack.push({ ...entry, oldVal: currentValue })
            setCellRaw(entry.r, entry.c, entry.newVal); _generation++; recalculate()
        } else {
            undoStack.push({ ...entry, snap: takeSnapshot(), oldRows: rows, oldCols: cols })
            restoreSnapshot(entry.restoreSnap); rows = entry.restoreRows; cols = entry.restoreCols; recalculate()
        }
        return true
    }

    function getCellForDisplay(row, col) {
        const cell = getCellRaw(row, col)
        const key = cellKey(row, col)
        const value = resolveCellValue(key)
        if (typeof value === 'string' && value.startsWith('#')) {
            return { raw: cell.raw, computed: null, error: value }
        }
        return { raw: cell.raw, computed: value, error: null }
    }

    // ── Task 3: Serialize all cell data for localStorage persistence ──
    // Returns a plain object safe to JSON.stringify
    // Note: Only raw formulas are serialized — computed values are recalculated on load
    function serialize() {
        const cellData = {}
        for (const [key, value] of cells.entries()) {
            cellData[key] = value.raw
        }
        return { cells: cellData, rows, cols }
    }

    // ── Task 3: Load cell data from a serialized object ──
    // Called on app startup to restore state from localStorage
    // Clears existing data first to avoid stale state
    function loadFromData(data) {
        if (!data || !data.cells) return
        // Clear existing state before loading
        cells.clear(); graph.clear(); computedCache.clear(); dirtyCells.clear(); _generation++
        for (const [key, raw] of Object.entries(data.cells)) {
            if (!raw || raw.trim() === '') continue
            const coords = parseCellKey(key)
            if (coords) setCellRaw(coords.row, coords.col, raw)
        }
        recalculate()
    }

    return {
        get rows() { return rows },
        get cols() { return cols },
        getCell: getCellForDisplay,
        setCell: executeSetCell,
        insertRow: executeInsertRow,
        deleteRow: executeDeleteRow,
        insertColumn: executeInsertColumn,
        deleteColumn: executeDeleteColumn,
        undo,
        redo,
        canUndo: () => undoStack.length > 0,
        canRedo: () => redoStack.length > 0,
        // Task 3 — persistence API
        serialize,
        loadFromData,
    }
}