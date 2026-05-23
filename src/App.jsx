import { useState, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from './supabase'

const ADMIN_PASSWORD = 'ReBella2026!'

const PRINT_HOUSES = [
  'OpenBooks',
  'Pagony',
  'Gabo',
  '10:40',
  'Líra Könyv',
  'Libri-Bookline',
  'Libri Könyvkiadó',
  'Alexandra Kiadó',
  'Kossuth Kiadó',
  'Magvető Kiadó',
  'Helikon Kiadó',
  'Európa Könyvkiadó',
  'Jelenkor Kiadó',
  'Park Könyvkiadó',
  'Scolar Kiadó',
  'Jaffa Kiadó',
  'Kalligram Kiadó',
  'Móra Kiadó',
  'Kolibri Kiadó',
  'Manó Könyvek',
  'Pozsonyi Pagony',
  'Csimota Könyvkiadó',
  'HVG Kiadó',
  'Medicina Könyvkiadó',
  'Typotex Kiadó',
  'Osiris Kiadó',
  'Balassi Kiadó',
  "L'Harmattan Kiadó",
  'Corvina Kiadó',
  'General Press',
  'Saxum Kiadó',
  'Gold Book Kiadó',
  'Könyvmolyképző Kiadó',
  'Napraforgó Kiadó',
  'Book24',
  'EGYÉB',
]

function makeCode(termekkkod, ean, termeknev, printHouse) {
  if (termekkkod != null && String(termekkkod).trim() !== '') return String(termekkkod).trim()
  if (ean != null && String(ean).trim() !== '') return String(ean).trim()
  // Fallback: stable synthetic key from title + print house
  const slug = String(termeknev ?? '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // strip accents
    .toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '').substring(0, 60)
  return `${printHouse}__${slug}`
}

// Returns list of sheet names in an xlsx buffer
function getExcelSheets(buffer) {
  const wb = XLSX.read(buffer, { type: 'array' })
  return wb.SheetNames
}

// Detect which column format a sheet uses based on its header row
// Returns 'ob' for OpenBooks format, 'idegen' for Book24/foreign format
function detectSheetFormat(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { range: 1, defval: null, header: 1 })
  if (!rows.length) return 'ob'
  const headers = rows[0].map(h => String(h ?? '').toLowerCase())
  if (headers.includes('cím') || headers.includes('isbn')) return 'idegen'
  return 'ob'
}

function parseExcelBooks(buffer, printHouse, sheetName) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true })

  // If no sheet specified, try 'Aktuális készlet' first, then first sheet
  const targetSheet = sheetName
    ?? (wb.SheetNames.includes('Aktuális készlet') ? 'Aktuális készlet' : wb.SheetNames[0])

  const ws = wb.Sheets[targetSheet]
  if (!ws) throw new Error(`Nem található "${targetSheet}" munkalap a fájlban.`)

  const format = detectSheetFormat(ws)

  if (format === 'idegen') {
    // Book24 / foreign format: szerző, cím, kiadó, ISBN, fogyasztói áfás ár, kedvezmény, kedvezményes ár, mennyiség
    const rows = XLSX.utils.sheet_to_json(ws, { range: 1, defval: null })
    return rows
      .filter(r => r['cím'] != null && String(r['cím']).trim() !== '')
      .map(r => {
        const ean = r['ISBN'] != null ? String(r['ISBN']).trim() : null
        const rawKedv = r['kedvezmény']
        return {
          termekkkod:          makeCode(null, ean, r['cím'], printHouse),
          ean,
          szerzo:              r['szerző'] != null ? String(r['szerző']).trim() : null,
          termeknev:           r['cím'] != null ? String(r['cím']).trim() : null,
          akcio_keszlet:       0,
          normal_keszlet:      r['mennyiség'] != null ? Number(r['mennyiség']) : 0,
          fogy_ar:             r['fogyasztói áfás ár'] != null ? Number(r['fogyasztói áfás ár']) : null,
          netto_ar:            null,
          ob_netto_ar:         null,
          kedvezmeny_szazalek: rawKedv != null && rawKedv !== '' ? Number(rawKedv) / 100 : null,
          kedvezmenyes_ar:     r['kedvezményes ár'] != null ? Number(r['kedvezményes ár']) : null,
          arkototteg_lejar:    null,
          print_house:         printHouse,
          stocked_at:          new Date().toISOString(),
        }
      })
  }

  // Default: OB format
  const now = new Date().toISOString()
  const rows = XLSX.utils.sheet_to_json(ws, { range: 1, defval: null })
  return rows
    .filter(r => r['Terméknév'] != null && String(r['Terméknév']).trim() !== '')
    .map(r => ({
      termekkkod:          makeCode(r['Termékkód'], r['EAN'], r['Terméknév'], printHouse),
      ean:                 r['EAN'] != null ? String(r['EAN']) : null,
      szerzo:              r['szerző'] ?? null,
      termeknev:           r['Terméknév'] ?? null,
      akcio_keszlet:       r['akciós készlet'] ?? 0,
      normal_keszlet:      r['normál készlet'] ?? 0,
      fogy_ar:             r['fogy.ár (5% áfás)'] ?? null,
      netto_ar:            r['nettó ár'] ?? null,
      ob_netto_ar:         r['árréssel csökentett nettó ár (OB-nak fizetendő)'] ?? null,
      kedvezmeny_szazalek: r['javasolt kedvezmény vevőknek'] ?? null,
      kedvezmenyes_ar:     r['javasolt kedvezményes, bolti áfás ár'] ?? null,
      arkototteg_lejar:    r['árkötöttség lejár'] instanceof Date
                             ? r['árkötöttség lejár'].toISOString().split('T')[0]
                             : null,
      print_house:         printHouse,
      stocked_at:          now,
    }))
}

function parseCsvBooks(buffer, printHouse) {
  // Decode buffer to string (handle UTF-8 BOM if present)
  let text = new TextDecoder('utf-8').decode(buffer)
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1) // strip BOM

  // Parse via SheetJS — row 0 is metadata junk, row 1 is the real header
  const wb = XLSX.read(text, { type: 'string' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(ws, { range: 1, defval: null })

  return rows
    .filter(r => r['cím'] != null && String(r['cím']).trim() !== '') // skip blank rows
    .map(r => {
      const ean = r['ISBN'] != null ? String(r['ISBN']).trim() : null
      const rawKedv = r['kedvezmény']
      const kedvezmeny = rawKedv != null && rawKedv !== '' ? Number(rawKedv) / 100 : null
      return {
        termekkkod:          makeCode(null, ean, r['cím'], printHouse),
        ean,
        szerzo:              r['szerző'] != null ? String(r['szerző']).trim() : null,
        termeknev:           r['cím'] != null ? String(r['cím']).trim() : null,
        akcio_keszlet:       0,
        normal_keszlet:      r['mennyiség'] != null ? Number(r['mennyiség']) : 0,
        fogy_ar:             r['fogyasztói áfás ár'] != null ? Number(r['fogyasztói áfás ár']) : null,
        netto_ar:            null,
        ob_netto_ar:         null,
        kedvezmeny_szazalek: kedvezmeny,
        kedvezmenyes_ar:     r['kedvezményes ár'] != null ? Number(r['kedvezményes ár']) : null,
        arkototteg_lejar:    null,
        print_house:         printHouse,
        stocked_at:          new Date().toISOString(),
      }
    })
}

function AdminPanel() {
  const [input, setInput] = useState('')
  const [unlocked, setUnlocked] = useState(false)
  const [error, setError] = useState(false)

  function handleSubmit(e) {
    e.preventDefault()
    if (input === ADMIN_PASSWORD) {
      setUnlocked(true)
      setError(false)
    } else {
      setError(true)
      setInput('')
    }
  }

  if (!unlocked) {
    return (
      <div className="flex flex-col items-center justify-center py-24">
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-8 w-full max-w-sm">
          <h2 className="text-lg font-bold text-gray-900 mb-1">Admin belépés</h2>
          <p className="text-sm text-gray-500 mb-6">Add meg a jelszót a folytatáshoz.</p>
          <form onSubmit={handleSubmit} className="flex flex-col gap-3">
            <input
              type="password"
              value={input}
              onChange={(e) => { setInput(e.target.value); setError(false) }}
              placeholder="Jelszó"
              autoFocus
              className={`w-full px-4 py-3 rounded-lg border bg-white text-gray-900
                placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent
                ${error ? 'border-red-400' : 'border-gray-300'}`}
            />
            {error && <p className="text-sm text-red-600">Hibás jelszó, próbáld újra.</p>}
            <button
              type="submit"
              className="w-full py-3 rounded-lg bg-[#C0392B] text-white font-medium hover:bg-[#A93226] transition-colors"
            >
              Belépés
            </button>
          </form>
        </div>
      </div>
    )
  }

  return (
    <div className="flex flex-col gap-6">
      <div className="flex items-center justify-between">
        <h2 className="text-lg font-bold text-gray-900">Admin</h2>
        <button onClick={() => setUnlocked(false)} className="text-sm text-gray-500 hover:text-gray-700">
          Kilépés
        </button>
      </div>
      <ExcelUpload />
      <ManualBookAdd />
      <InventoryExport />
      <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-6">
        <h3 className="font-bold text-gray-900 mb-1">Ötletdoboz / problémák</h3>
        <p className="text-sm text-gray-500 mb-5">Piros = sürgős, sárga = közepes, zöld = alacsony prioritás.</p>
        <IdeasAdmin />
      </div>
    </div>
  )
}

function ExcelUpload() {
  const [printHouse, setPrintHouse] = useState('OpenBooks')
  const [sheetNames, setSheetNames] = useState(null)   // null = not loaded yet
  const [selectedSheet, setSelectedSheet] = useState(null)
  const [fileBuffer, setFileBuffer] = useState(null)
  const [isCsv, setIsCsv] = useState(false)
  const [parsedBooks, setParsedBooks] = useState(null)
  const [parseError, setParseError] = useState(null)
  const [uploading, setUploading] = useState(false)
  const [uploadDone, setUploadDone] = useState(false)

  function resetState() {
    setParsedBooks(null); setParseError(null); setUploadDone(false)
    setSheetNames(null); setSelectedSheet(null); setFileBuffer(null); setIsCsv(false)
  }

  function parseSheet(buffer, sheet, ph, csv) {
    try {
      const books = csv
        ? parseCsvBooks(buffer, ph)
        : parseExcelBooks(buffer, ph, sheet)
      setParsedBooks(books)
      setParseError(null)
    } catch (err) {
      setParseError(err.message)
      setParsedBooks(null)
    }
  }

  function handleFile(e) {
    const file = e.target.files[0]
    if (!file) return
    resetState()
    const csv = file.name.toLowerCase().endsWith('.csv')
    setIsCsv(csv)
    const reader = new FileReader()
    reader.onload = (ev) => {
      const buf = ev.target.result
      setFileBuffer(buf)
      const ph = printHouse.trim() || 'Ismeretlen'
      if (csv) {
        parseSheet(buf, null, ph, true)
      } else {
        const sheets = getExcelSheets(buf)
        if (sheets.length === 1 || sheets.includes('Aktuális készlet')) {
          // Single sheet or old-format file — parse immediately
          const sheet = sheets.includes('Aktuális készlet') ? 'Aktuális készlet' : sheets[0]
          setSelectedSheet(sheet)
          parseSheet(buf, sheet, ph, false)
        } else {
          // Multiple sheets — show selector
          setSheetNames(sheets)
          setSelectedSheet(sheets[0])
          parseSheet(buf, sheets[0], ph, false)
        }
      }
    }
    reader.readAsArrayBuffer(file)
  }

  function handleSheetChange(sheet) {
    setSelectedSheet(sheet)
    setParsedBooks(null)
    const ph = printHouse.trim() || 'Ismeretlen'
    parseSheet(fileBuffer, sheet, ph, false)
  }

  async function handleUpload() {
    if (!parsedBooks || parsedBooks.length === 0) return
    setUploading(true)

    // Deduplicate within the file: if duplicate termekkkod rows, keep the last one
    const seen = new Map()
    for (const book of parsedBooks) seen.set(book.termekkkod, book)
    const deduped = Array.from(seen.values())

    // Fetch existing books for this print_house so we can ADD stock instead of replace
    const codes = deduped.map(b => b.termekkkod)
    const { data: existing } = await supabase
      .from('books')
      .select('termekkkod, normal_keszlet, akcio_keszlet')
      .eq('print_house', printHouse)
      .in('termekkkod', codes)
    const existingMap = new Map((existing ?? []).map(b => [b.termekkkod, b]))

    // Merge: add new stock on top of existing stock (each upload = new delivery)
    const merged = deduped.map(book => {
      const prev = existingMap.get(book.termekkkod)
      return {
        ...book,
        normal_keszlet: (prev?.normal_keszlet ?? 0) + (book.normal_keszlet ?? 0),
        akcio_keszlet:  (prev?.akcio_keszlet  ?? 0) + (book.akcio_keszlet  ?? 0),
      }
    })

    // Upsert in batches of 50
    const BATCH = 50
    for (let i = 0; i < merged.length; i += BATCH) {
      const { error } = await supabase
        .from('books')
        .upsert(merged.slice(i, i + BATCH), { onConflict: 'termekkkod,print_house' })
      if (error) { alert('Hiba feltöltéskor: ' + error.message); setUploading(false); return }
    }

    setUploading(false)
    setUploadDone(true)
    setParsedBooks(null)
    // Reset file input
    document.getElementById('excel-upload-input').value = ''
  }

  return (
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-6">
      <h3 className="font-bold text-gray-900 mb-1">Készlet feltöltése</h3>
      <p className="text-sm text-gray-500 mb-5">
        Tölts fel egy új készletlistát (.xlsx vagy .csv). Az adott kiadóhoz tartozó könyvek frissülnek, újak hozzáadódnak.
      </p>

      <div className="flex flex-col gap-4">
        {/* Print house */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Kiadó neve</label>
          <select
            value={printHouse}
            onChange={(e) => {
              const ph = e.target.value
              setPrintHouse(ph)
              setParsedBooks(null)
              setUploadDone(false)
              if (fileBuffer) parseSheet(fileBuffer, selectedSheet, ph.trim() || 'Ismeretlen', isCsv)
            }}
            className="w-full px-4 py-2.5 rounded-lg border border-gray-300 bg-white text-gray-900
              focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent"
          >
            {PRINT_HOUSES.map((ph) => (
              <option key={ph} value={ph}>{ph}</option>
            ))}
          </select>
        </div>

        {/* File picker */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Fájl feltöltése (.xlsx vagy .csv)</label>
          <input
            id="excel-upload-input"
            type="file"
            accept=".xlsx,.csv"
            onChange={handleFile}
            className="w-full text-sm text-gray-600 file:mr-3 file:py-2 file:px-4
              file:rounded-lg file:border-0 file:text-sm file:font-medium
              file:bg-[#C0392B] file:text-white hover:file:bg-[#A93226] cursor-pointer"
          />
        </div>

        {/* Sheet selector — only shown for multi-sheet xlsx files */}
        {sheetNames && sheetNames.length > 1 && (
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Munkalap kiválasztása</label>
            <select
              value={selectedSheet ?? ''}
              onChange={e => handleSheetChange(e.target.value)}
              className="w-full px-4 py-2.5 rounded-lg border border-gray-300 bg-white text-gray-900
                focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent"
            >
              {sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        )}

        {/* Parse error */}
        {parseError && (
          <p className="text-sm text-red-600 bg-red-50 rounded-lg px-4 py-3">{parseError}</p>
        )}

        {/* Preview */}
        {parsedBooks && (
          <div className="bg-amber-50 border border-amber-200 rounded-lg px-4 py-3">
            <p className="text-sm font-medium text-amber-800">
              {parsedBooks.length} könyvet találtunk a fájlban.
            </p>
            <p className="text-xs text-amber-700 mt-0.5">
              Új könyvek hozzáadódnak, meglévők frissülnek (<strong>{parsedBooks[0].print_house}</strong>).
            </p>
          </div>
        )}

        {/* Upload done */}
        {uploadDone && (
          <p className="text-sm text-green-700 bg-green-50 rounded-lg px-4 py-3">
            ✓ Sikeresen feltöltve!
          </p>
        )}

        {/* Upload button */}
        {parsedBooks && (
          <button
            onClick={handleUpload}
            disabled={uploading}
            className="w-full py-3 rounded-lg bg-[#C0392B] text-white font-medium
              hover:bg-[#A93226] disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
          >
            {uploading ? 'Feltöltés...' : `${parsedBooks.length} könyv feltöltése`}
          </button>
        )}
      </div>
    </div>
  )
}

const EMPTY_BOOK = {
  termekkkod: '', ean: '', szerzo: '', termeknev: '',
  normal_keszlet: '', akcio_keszlet: '',
  fogy_ar: '', kedvezmeny_szazalek: '', kedvezmenyes_ar: '',
  arkototteg_lejar: '', print_house: 'OpenBooks',
}

function ManualBookAdd() {
  const [form, setForm] = useState(EMPTY_BOOK)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [error, setError] = useState(null)

  function set(field, value) {
    setForm(f => ({ ...f, [field]: value }))
    setSaved(false)
    setError(null)
  }

  async function handleSubmit(e) {
    e.preventDefault()
    if (!form.termeknev.trim()) { setError('A könyv címe kötelező.'); return }
    setSaving(true)
    // If no termekkkod given, auto-generate from title (same as Excel import)
    const resolvedCode = form.termekkkod.trim()
      || makeCode(null, form.ean.trim() || null, form.termeknev.trim(), form.print_house)
    const { error } = await supabase.from('books').upsert({
      termekkkod:          resolvedCode,
      ean:                 form.ean.trim() || null,
      szerzo:              form.szerzo.trim() || null,
      termeknev:           form.termeknev.trim(),
      normal_keszlet:      Number(form.normal_keszlet) || 0,
      akcio_keszlet:       Number(form.akcio_keszlet) || 0,
      fogy_ar:             form.fogy_ar !== '' ? Number(form.fogy_ar) : null,
      kedvezmeny_szazalek: form.kedvezmeny_szazalek !== '' ? Number(form.kedvezmeny_szazalek) / 100 : null,
      kedvezmenyes_ar:     form.kedvezmenyes_ar !== '' ? Number(form.kedvezmenyes_ar) : null,
      arkototteg_lejar:    form.arkototteg_lejar || null,
      print_house:         form.print_house,
      stocked_at:          new Date().toISOString(),
    }, { onConflict: 'termekkkod,print_house' })
    setSaving(false)
    if (error) { setError(error.message); return }
    setSaved(true)
    setForm(EMPTY_BOOK)
  }

  const inputClass = "w-full px-4 py-2.5 rounded-lg border border-gray-300 bg-white text-gray-900 focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent"
  const labelClass = "block text-sm font-medium text-gray-700 mb-1"

  return (
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-6">
      <h3 className="font-bold text-gray-900 mb-1">Könyv manuális hozzáadása</h3>
      <p className="text-sm text-gray-500 mb-5">
        Tölts ki minden ismert mezőt. Ha a termékkód már létezik az adott kiadónál, frissül.
      </p>
      <form onSubmit={handleSubmit} className="flex flex-col gap-4">

        {/* Kiadó */}
        <div>
          <label className={labelClass}>Kiadó *</label>
          <select value={form.print_house} onChange={e => set('print_house', e.target.value)} className={inputClass}>
            {PRINT_HOUSES.map(ph => <option key={ph} value={ph}>{ph}</option>)}
          </select>
        </div>

        {/* Cím + Szerző */}
        <div>
          <label className={labelClass}>Cím *</label>
          <input type="text" value={form.termeknev} onChange={e => set('termeknev', e.target.value)} className={inputClass} placeholder="Könyv címe" />
        </div>
        <div>
          <label className={labelClass}>Szerző</label>
          <input type="text" value={form.szerzo} onChange={e => set('szerzo', e.target.value)} className={inputClass} placeholder="Pl. Kovács János; Kiss Éva" />
        </div>

        {/* Termékkód + EAN */}
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className={labelClass}>Termékkód <span className="font-normal text-gray-400">(ha nincs, auto-generálódik)</span></label>
            <input type="text" value={form.termekkkod} onChange={e => set('termekkkod', e.target.value)} className={inputClass} placeholder="pl. O15724734 — hagyd üresen ha nincs" />
          </div>
          <div>
            <label className={labelClass}>EAN</label>
            <input type="text" value={form.ean} onChange={e => set('ean', e.target.value)} className={inputClass} placeholder="9789..." />
          </div>
        </div>

        {/* Készlet */}
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className={labelClass}>Normál készlet</label>
            <input type="number" min="0" value={form.normal_keszlet} onChange={e => set('normal_keszlet', e.target.value)} className={inputClass} placeholder="0" />
          </div>
          <div>
            <label className={labelClass}>Akciós készlet</label>
            <input type="number" min="0" value={form.akcio_keszlet} onChange={e => set('akcio_keszlet', e.target.value)} className={inputClass} placeholder="0" />
          </div>
        </div>

        {/* Árak */}
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className={labelClass}>Eredeti ár (Ft)</label>
            <input type="number" min="0" value={form.fogy_ar} onChange={e => set('fogy_ar', e.target.value)} className={inputClass} placeholder="4999" />
          </div>
          <div>
            <label className={labelClass}>Kedvezményes ár (Ft)</label>
            <input type="number" min="0" value={form.kedvezmenyes_ar} onChange={e => set('kedvezmenyes_ar', e.target.value)} className={inputClass} placeholder="3999" />
          </div>
        </div>

        {/* Kedvezmény % + Árkötöttség */}
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className={labelClass}>Kedvezmény %</label>
            <input type="number" min="0" max="100" value={form.kedvezmeny_szazalek} onChange={e => set('kedvezmeny_szazalek', e.target.value)} className={inputClass} placeholder="20" />
          </div>
          <div>
            <label className={labelClass}>Árkötöttség lejár</label>
            <input type="date" value={form.arkototteg_lejar} onChange={e => set('arkototteg_lejar', e.target.value)} className={inputClass} />
          </div>
        </div>

        {error && <p className="text-sm text-red-600 bg-red-50 rounded-lg px-4 py-3">{error}</p>}
        {saved && <p className="text-sm text-green-700 bg-green-50 rounded-lg px-4 py-3">✓ Könyv sikeresen mentve!</p>}

        <button
          type="submit"
          disabled={saving}
          className="w-full py-3 rounded-lg bg-[#C0392B] text-white font-medium
            hover:bg-[#A93226] disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
        >
          {saving ? 'Mentés...' : 'Könyv mentése'}
        </button>
      </form>
    </div>
  )
}

function InventoryExport() {
  const [printHouse, setPrintHouse] = useState('OpenBooks')
  const [downloading, setDownloading] = useState(false)

  async function handleDownload() {
    setDownloading(true)
    const { data, error } = await supabase
      .from('books')
      .select('*')
      .eq('print_house', printHouse)
      .order('termeknev', { ascending: true })

    if (error) { alert('Hiba: ' + error.message); setDownloading(false); return }

    const headers = ['Termékkód', 'EAN', 'Szerző', 'Cím', 'Kiadó', 'Normál készlet', 'Akciós készlet', 'Eredeti ár (Ft)', 'Kedvezmény %', 'Kedvezményes ár (Ft)', 'Árkötöttség lejár']
    const rows = data.map(b => [
      b.termekkkod,
      b.ean ?? '',
      b.szerzo ?? '',
      b.termeknev ?? '',
      b.print_house ?? '',
      b.normal_keszlet ?? 0,
      b.akcio_keszlet ?? 0,
      b.fogy_ar != null ? Math.round(b.fogy_ar) : '',
      b.kedvezmeny_szazalek != null ? Math.round(b.kedvezmeny_szazalek * 100) + '%' : '',
      b.kedvezmenyes_ar != null ? Math.round(b.kedvezmenyes_ar) : '',
      b.arkototteg_lejar ?? '',
    ])
    const csv = [headers, ...rows]
      .map(row => row.map(cell => {
        const s = String(cell ?? '')
        return s.includes(',') || s.includes('"') || s.includes('\n')
          ? '"' + s.replace(/"/g, '""') + '"' : s
      }).join(',')).join('\n')

    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    const today = new Date().toISOString().split('T')[0]
    const slug = printHouse.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '')
    a.href = url
    a.download = `rebella_keszlet_${slug}_${today}.csv`
    a.click()
    URL.revokeObjectURL(url)
    setDownloading(false)
  }

  return (
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-6">
      <h3 className="font-bold text-gray-900 mb-1">Készlet letöltése</h3>
      <p className="text-sm text-gray-500 mb-5">
        Töltsd le egy kiadó teljes aktuális készletét CSV-ben.
      </p>
      <div className="flex flex-col gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Kiadó</label>
          <select
            value={printHouse}
            onChange={e => setPrintHouse(e.target.value)}
            className="w-full px-4 py-2.5 rounded-lg border border-gray-300 bg-white text-gray-900
              focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent"
          >
            {PRINT_HOUSES.map(ph => <option key={ph} value={ph}>{ph}</option>)}
          </select>
        </div>
        <button
          onClick={handleDownload}
          disabled={downloading}
          className="w-full py-3 rounded-lg bg-[#C0392B] text-white font-medium
            hover:bg-[#A93226] disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
        >
          {downloading ? 'Letöltés...' : 'Készlet CSV letöltése'}
        </button>
      </div>
    </div>
  )
}

function formatPrice(value) {
  if (value == null || isNaN(value)) return '—'
  return Math.round(value).toLocaleString('hu-HU') + ' Ft'
}

function formatDiscount(value) {
  if (value == null || isNaN(value)) return '—'
  return Math.round(value * 100) + '%'
}

const FUN_FACTS = [
  'A kenyér az egyik legrégebbi élelmiszer a világon',
  'Az egyiptomiak sütötték az első kovászos kenyeret',
  'A rozskenyér az északi hideg éghajlat miatt terjedt el Európában',
  'A baguette árát törvény szabályozza Franciaországban',
  'A ciabattát csak 1982-ben találták fel',
  'A toast kenyér az 1920-as években jelent meg',
  'A teljes kiőrlésű kenyér háromszor annyi rostot tartalmaz mint a fehér',
  'A kovászos kenyér alacsonyabb glikémiás indexű mint a fehér kenyér',
  'A kenyérpirítás csökkenti a glikémiás indexet',
  'A bagel főzve majd sütve készül',
  'A Dunakanyarnál a folyó közel 180 fokos kanyart tesz',
  'A Dunakanyart a Börzsöny és a Visegrádi-hegység fogja közre',
  'Visegrád a Dunakanyar egyik legikonikusabb városa',
  'A római korban limes-erődök sorakoztak a Duna mentén',
  'Zebegény közelében van a legszebb kilátópont a kanyarra',
  'A Dunakanyar a nemzeti romantika kedvelt festészeti témája volt',
  'A jégmadár a Duna egyik legszínpompásabb madara',
  'A fehér gólya rendszeres fészkelő a Dunakanyar falvaiban',
  'A fekete gólya zárt erdőkben költ a folyóvölgyek közelében',
  'A nagy kócsag az ártereken halász',
  'A kormorán kolóniái a Duna nagy fáin fészkelnek',
  'A dankasirály a Dunán is költ nem csak tengerparton',
  'A szürke gém magányosan és éjjel-nappal aktívan vadászik',
  'A halászsas ritka vendég de időnként feltűnik a Dunakanyarban',
  'A tőkés réce a leggyakoribb kacsafaj a Dunán',
  'Az espresso neve olaszul gyorsan préselt vizet jelent',
  'Az első kávéházat Konstantinápolyban nyitották 1475-ben',
  'A kávébab botanikailag bogyónak számít nem babnak',
  'Etiópia a kávé szülőhazája',
  'A finnek isszák a legtöbb kávét fejenként a világon',
  'A kávé a világ második legkereskedettebb árucikke az olaj után',
  'A koffein 45 percen belül teljesen felszívódik a véráramba',
  'A decaf kávé sem teljesen koffeintartalom-mentes',
  'A cold brew kávét 12-24 óráig áztatják hideg vízben',
  'A kapucsínó nevét a kapucinus szerzetesek barna csuhájáról kapta',
  'A mocha névadója egy jemeni kikötőváros',
  'A kávéfa 3-4 évig tart mire először terem',
  'Egy csésze espressohoz kb. 50 kávébogyó kell',
  'A robusta kétszer annyi koffeint tartalmaz mint az arabica',
  'A vietnámi kávé jellegzetessége a sűrített tej',
  'A bécsi kávéházi kultúra UNESCO örökség',
  'A latte art az 1980-as években Seattle-ben terjedt el',
  'A flat white Ausztráliából indult világhódító útjára',
  'A kávé szó az arab qahwa kifejezésből ered',
  'A réti sas a Duna mentén fészkel és halból él',
  'A búbos vöcsök látványos nászjelzéseiről ismert a dunai tavakban',
  'A nagy lilik és a vetési lúd télen tömegesen pihen a Duna homokpadjain',
  'A fekete rigó a Dunakanyar erdőinek legdalosabb madara',
]

function randomFact() {
  return FUN_FACTS[Math.floor(Math.random() * FUN_FACTS.length)]
}

function formatSoldDate(iso) {
  const d = new Date(iso)
  const pad = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`
}

function getStockStatus(normalRemaining, akcioRemaining) {
  if (normalRemaining > 0) return { label: `Készleten: ${normalRemaining} db`, color: 'bg-green-100 text-green-800' }
  if (akcioRemaining > 0) return { label: `Akciós készlet: ${akcioRemaining} db`, color: 'bg-orange-100 text-orange-800' }
  return { label: 'Nincs készleten', color: 'bg-red-100 text-red-800' }
}

function useDebounce(value, delay) {
  const [debounced, setDebounced] = useState(value)
  useEffect(() => {
    const timer = setTimeout(() => setDebounced(value), delay)
    return () => clearTimeout(timer)
  }, [value, delay])
  return debounced
}

function BookCard({ book, soldCount, onSell }) {
  const normalStock = book.normal_keszlet ?? 0
  const akcioStock = book.akcio_keszlet ?? 0
  const normalRemaining = Math.max(0, normalStock - soldCount)
  const akcioRemaining = soldCount > normalStock
    ? Math.max(0, akcioStock - (soldCount - normalStock))
    : akcioStock
  const totalRemaining = normalRemaining + akcioRemaining
  const stock = getStockStatus(normalRemaining, akcioRemaining)

  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-100 p-4 sm:p-5">
      <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-3">
        <div className="flex-1 min-w-0">
          <h3 className="text-lg font-bold text-gray-900 leading-tight">
            {book.termeknev || 'Ismeretlen cím'}
          </h3>
          <p className="text-gray-600 mt-1">{book.szerzo || 'Ismeretlen szerző'}</p>
          <p className="text-xs text-gray-400 mt-1">
            Kód: {book.termekkkod} | EAN: {book.ean ?? ''}
          </p>
        </div>
        <div className="flex flex-col items-start sm:items-end gap-2 shrink-0">
          <span className={`inline-block text-xs font-medium px-2.5 py-1 rounded-full ${stock.color}`}>
            {stock.label}
          </span>
          {book.kedvezmeny_szazalek != null && (
            <span className="inline-block text-xs font-medium px-2.5 py-1 rounded-full bg-amber-100 text-amber-800">
              {formatDiscount(book.kedvezmeny_szazalek)} kedvezmény
            </span>
          )}
        </div>
      </div>

      <div className="mt-4 flex flex-wrap items-end gap-x-6 gap-y-2">
        <div>
          <p className="text-xs text-gray-500">Eredeti ár</p>
          <p className="text-gray-700 line-through">{formatPrice(book.fogy_ar)}</p>
        </div>
        <div>
          <p className="text-xs text-gray-500">Kedvezményes ár</p>
          <p className="text-xl font-semibold text-[#C0392B]">{formatPrice(book.kedvezmenyes_ar)}</p>
        </div>
      </div>

      {book.arkototteg_lejar && (
        <p className="mt-3 text-sm text-amber-700">
          ⚠️ Árkötöttség lejár: {book.arkototteg_lejar}
        </p>
      )}

      <div className="mt-4 flex items-center gap-3">
        <button
          onClick={() => onSell(book)}
          disabled={totalRemaining === 0}
          className="text-sm font-medium px-4 py-2 rounded-lg transition-colors
            bg-[#C0392B] text-white hover:bg-[#A93226] disabled:bg-gray-300 disabled:text-gray-500 disabled:cursor-not-allowed"
        >
          Eladás
        </button>
        {soldCount > 0 && (
          <span className="text-sm text-gray-500">{soldCount} db eladva</span>
        )}
      </div>
    </div>
  )
}

function SoldTable({ soldList, onUndo, onExport, onClear }) {
  const [filterHouse, setFilterHouse] = useState('Összes')
  const [search, setSearch] = useState('')

  // Build list of print houses that actually appear in sold books
  const presentHouses = useMemo(() => {
    const set = new Set(soldList.map(i => i.print_house).filter(Boolean))
    return ['Összes', ...Array.from(set).sort()]
  }, [soldList])

  const filtered = useMemo(() => {
    let list = filterHouse === 'Összes' ? soldList : soldList.filter(i => i.print_house === filterHouse)
    if (search.trim()) {
      const q = search.toLowerCase().trim()
      list = list.filter(i =>
        (i.termeknev || '').toLowerCase().includes(q) ||
        (i.szerzo || '').toLowerCase().includes(q)
      )
    }
    return list
  }, [soldList, filterHouse, search])

  function handleUndo(id, title) {
    if (window.confirm(`Biztos hogy visszavonod?\n\n„${title}"`)) onUndo(id)
  }

  if (soldList.length === 0) {
    return (
      <div className="text-center py-16 text-gray-400">
        Még nem adtál el könyvet.
      </div>
    )
  }

  return (
    <div>
      {/* Search */}
      <input
        type="text"
        value={search}
        onChange={e => setSearch(e.target.value)}
        placeholder="Keresés cím vagy szerző alapján..."
        className="w-full px-4 py-2.5 mb-4 rounded-lg border border-gray-300 bg-white text-gray-900
          placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent text-sm"
      />

      {/* Filter + actions row */}
      <div className="flex flex-wrap items-center gap-3 mb-5">
        <select
          value={filterHouse}
          onChange={e => setFilterHouse(e.target.value)}
          className="px-3 py-2 text-sm rounded-lg border border-gray-300 bg-white text-gray-900
            focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent"
        >
          {presentHouses.map(ph => (
            <option key={ph} value={ph}>{ph}</option>
          ))}
        </select>
        <span className="text-sm text-gray-400">{filtered.length} tétel</span>
        <div className="flex gap-3 ml-auto">
          <button
            onClick={() => onExport(filtered, filterHouse)}
            className="text-sm font-medium px-4 py-2 rounded-lg bg-[#C0392B] text-white hover:bg-[#A93226] transition-colors"
          >
            CSV letöltése
          </button>
          <button
            onClick={onClear}
            className="text-sm font-medium px-4 py-2 rounded-lg border border-red-300 text-red-700 hover:bg-red-50 transition-colors"
          >
            Lista törlése
          </button>
        </div>
      </div>

      {filtered.length === 0 && (
        <p className="text-center py-10 text-gray-400">Nincs találat.</p>
      )}

      {/* Mobile cards */}
      <div className="sm:hidden flex flex-col gap-3">
        {filtered.map((item) => (
          <div key={item.id} className="bg-white rounded-lg shadow-sm border border-gray-100 p-4">
            <p className="font-bold text-gray-900">{item.termeknev}</p>
            <p className="text-sm text-gray-600">{item.szerzo}</p>
            <div className="mt-2 flex flex-wrap gap-x-4 gap-y-1 text-sm text-gray-700">
              <span>Eredeti: {formatPrice(item.fogy_ar)}</span>
              <span>Kedv.: {Math.round((item.kedvezmeny_szazalek ?? 0) * 100)}%</span>
              <span>Ár: {formatPrice(item.kedvezmenyes_ar)}</span>
            </div>
            <p className="text-xs text-gray-400 mt-1">{formatSoldDate(item.eladva_datum)}</p>
            <button onClick={() => handleUndo(item.id, item.termeknev)} className="mt-2 text-sm text-red-600 hover:text-red-800">
              Visszavonás
            </button>
          </div>
        ))}
      </div>

      {/* Desktop table */}
      <div className="hidden sm:block overflow-x-auto">
        <table className="w-full text-sm text-left">
          <thead className="text-xs text-gray-500 uppercase border-b border-gray-200">
            <tr>
              <th className="py-3 pr-4">Cím</th>
              <th className="py-3 pr-4">Szerző</th>
              <th className="py-3 pr-4 text-right">Eredeti ár</th>
              <th className="py-3 pr-4 text-right">Kedvezmény</th>
              <th className="py-3 pr-4 text-right">Kedv. ár</th>
              <th className="py-3 pr-4">Dátum</th>
              <th className="py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((item) => (
              <tr key={item.id} className="border-b border-gray-100 hover:bg-gray-50">
                <td className="py-3 pr-4 font-medium text-gray-900">{item.termeknev}</td>
                <td className="py-3 pr-4 text-gray-600">{item.szerzo}</td>
                <td className="py-3 pr-4 text-right text-gray-700">{formatPrice(item.fogy_ar)}</td>
                <td className="py-3 pr-4 text-right text-gray-700">{Math.round((item.kedvezmeny_szazalek ?? 0) * 100)}%</td>
                <td className="py-3 pr-4 text-right font-medium text-gray-900">{formatPrice(item.kedvezmenyes_ar)}</td>
                <td className="py-3 pr-4 text-gray-500">{formatSoldDate(item.eladva_datum)}</td>
                <td className="py-3">
                  <button onClick={() => handleUndo(item.id, item.termeknev)} className="text-red-600 hover:text-red-800">
                    Visszavonás
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

const PRIORITY_CONFIG = {
  red:   { label: 'Sürgős',    bg: 'bg-red-100',    text: 'text-red-700',    dot: 'bg-red-500',    border: 'border-red-300',    order: 0 },
  amber: { label: 'Közepes',   bg: 'bg-amber-100',  text: 'text-amber-700',  dot: 'bg-amber-400',  border: 'border-amber-300',  order: 1 },
  green: { label: 'Alacsony',  bg: 'bg-green-100',  text: 'text-green-700',  dot: 'bg-green-500',  border: 'border-green-300',  order: 2 },
}

function IdeasTab() {
  const [text, setText] = useState('')
  const [priority, setPriority] = useState('amber')
  const [submitting, setSubmitting] = useState(false)
  const [submitted, setSubmitted] = useState(false)

  async function handleSubmit(e) {
    e.preventDefault()
    if (!text.trim()) return
    setSubmitting(true)
    const { error } = await supabase.from('ideas').insert({ text: text.trim(), priority })
    setSubmitting(false)
    if (!error) { setText(''); setSubmitted(true); setTimeout(() => setSubmitted(false), 3000) }
  }

  return (
    <div className="flex flex-col gap-6">
      <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-6">
        <h2 className="font-bold text-gray-900 text-lg mb-1">Ötletdoboz / problémák</h2>
        <p className="text-sm text-gray-500 mb-5">Írj le egy ötletet vagy problémát, és jelöld meg a prioritását.</p>

        <form onSubmit={handleSubmit} className="flex flex-col gap-4">
          <textarea
            value={text}
            onChange={e => setText(e.target.value)}
            placeholder="Írd le az ötletet vagy problémát..."
            rows={4}
            className="w-full px-4 py-3 rounded-lg border border-gray-300 bg-white text-gray-900
              placeholder-gray-400 resize-none focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent"
          />

          {/* Priority picker */}
          <div>
            <p className="text-sm font-medium text-gray-700 mb-2">Prioritás</p>
            <div className="flex gap-3">
              {Object.entries(PRIORITY_CONFIG).map(([key, cfg]) => (
                <button
                  key={key}
                  type="button"
                  onClick={() => setPriority(key)}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg border-2 text-sm font-medium transition-all
                    ${priority === key ? `${cfg.bg} ${cfg.text} ${cfg.border}` : 'bg-white border-gray-200 text-gray-500 hover:border-gray-300'}`}
                >
                  <span className={`w-2.5 h-2.5 rounded-full ${cfg.dot}`} />
                  {cfg.label}
                </button>
              ))}
            </div>
          </div>

          {submitted && <p className="text-sm text-green-700 bg-green-50 rounded-lg px-4 py-3">✓ Beküldve, köszönjük!</p>}

          <button
            type="submit"
            disabled={submitting || !text.trim()}
            className="w-full py-3 rounded-lg bg-[#C0392B] text-white font-medium
              hover:bg-[#A93226] disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
          >
            {submitting ? 'Beküldés...' : 'Beküldés'}
          </button>
        </form>
      </div>
    </div>
  )
}

function IdeasAdmin() {
  const [ideas, setIdeas] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    supabase.from('ideas').select('*').order('created_at', { ascending: false })
      .then(({ data }) => { if (data) setIdeas(data); setLoading(false) })
  }, [])

  // Real-time updates
  useEffect(() => {
    const channel = supabase.channel('ideas_changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'ideas' }, () => {
        supabase.from('ideas').select('*').order('created_at', { ascending: false })
          .then(({ data }) => { if (data) setIdeas(data) })
      }).subscribe()
    return () => supabase.removeChannel(channel)
  }, [])

  async function toggleSolved(id, current) {
    const { error } = await supabase.from('ideas').update({ solved: !current }).eq('id', id)
    if (!error) setIdeas(prev => prev.map(i => i.id === id ? { ...i, solved: !current } : i))
  }

  async function removeIdea(id) {
    const { error } = await supabase.from('ideas').delete().eq('id', id)
    if (!error) setIdeas(prev => prev.filter(i => i.id !== id))
  }

  const sorted = [...ideas].sort((a, b) =>
    (PRIORITY_CONFIG[a.priority]?.order ?? 9) - (PRIORITY_CONFIG[b.priority]?.order ?? 9)
  )

  if (loading) return <p className="text-sm text-gray-400 text-center py-6">Betöltés...</p>
  if (sorted.length === 0) return <p className="text-sm text-gray-400 text-center py-6">Még nincs beküldött ötlet.</p>

  return (
    <div className="flex flex-col gap-3">
      {sorted.map(idea => {
        const cfg = PRIORITY_CONFIG[idea.priority] ?? PRIORITY_CONFIG.amber
        return (
          <div key={idea.id} className={`rounded-lg border p-4 ${idea.solved ? 'opacity-50' : ''} ${cfg.bg} ${cfg.border}`}>
            <div className="flex items-start gap-3">
              <span className={`mt-1.5 w-2.5 h-2.5 rounded-full shrink-0 ${cfg.dot}`} />
              <div className="flex-1 min-w-0">
                <p className={`text-sm ${idea.solved ? 'line-through text-gray-400' : 'text-gray-800'}`}>
                  {idea.text}
                </p>
                <div className="flex items-center gap-3 mt-2 flex-wrap">
                  <span className={`text-xs font-medium ${cfg.text}`}>{cfg.label}</span>
                  <span className="text-xs text-gray-400">
                    {new Date(idea.created_at).toLocaleDateString('hu-HU')}
                  </span>
                </div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <button
                  onClick={() => toggleSolved(idea.id, idea.solved)}
                  title={idea.solved ? 'Visszaállítás' : 'Megoldva'}
                  className={`text-xs px-3 py-1.5 rounded-lg border font-medium transition-colors
                    ${idea.solved
                      ? 'bg-white border-gray-300 text-gray-600 hover:bg-gray-50'
                      : 'bg-white border-green-300 text-green-700 hover:bg-green-50'}`}
                >
                  {idea.solved ? 'Visszaállítás' : '✓ Megoldva'}
                </button>
                <button
                  onClick={() => removeIdea(idea.id)}
                  title="Törlés"
                  className="text-xs px-3 py-1.5 rounded-lg border border-red-200 text-red-500 hover:bg-red-50 transition-colors"
                >
                  Törlés
                </button>
              </div>
            </div>
          </div>
        )
      })}
    </div>
  )
}

export default function App() {
  const [books, setBooks] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)
  const [query, setQuery] = useState('')
  const [tab, setTab] = useState('search')
  const [soldList, setSoldList] = useState([])

  const debouncedQuery = useDebounce(query, 300)

  // Load books from Supabase
  useEffect(() => {
    supabase
      .from('books')
      .select('*')
      .order('termeknev', { ascending: true })
      .then(({ data, error }) => {
        if (error) setError(error.message)
        else setBooks(data)
      })
  }, [])

  // Load sold list from Supabase
  useEffect(() => {
    supabase
      .from('sold_books')
      .select('*')
      .order('eladva_datum', { ascending: false })
      .then(({ data, error }) => {
        if (error) setError(error.message)
        else setSoldList(data)
        setLoading(false)
      })
  }, [])

  // Real-time sync for sold books across devices
  useEffect(() => {
    const channel = supabase
      .channel('sold_books_changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'sold_books' }, () => {
        supabase
          .from('sold_books')
          .select('*')
          .order('eladva_datum', { ascending: false })
          .then(({ data }) => { if (data) setSoldList(data) })
      })
      .subscribe()
    return () => supabase.removeChannel(channel)
  }, [])

  const soldCounts = useMemo(() => {
    const map = new Map()
    for (const item of soldList) {
      map.set(item.termekkkod, (map.get(item.termekkkod) ?? 0) + 1)
    }
    return map
  }, [soldList])

  const filtered = useMemo(() => {
    if (!debouncedQuery.trim()) return []
    const q = debouncedQuery.toLowerCase().trim()
    return books.filter((b) =>
      (b.termeknev || '').toLowerCase().includes(q) ||
      (b.szerzo || '').toLowerCase().includes(q) ||
      (b.termekkkod || '').toLowerCase().includes(q) ||
      (b.ean || '').toLowerCase().includes(q)
    )
  }, [debouncedQuery, books])

  const handleSell = useCallback(async (book) => {
    const { data, error } = await supabase.from('sold_books').insert({
      termekkkod: book.termekkkod,
      ean: book.ean,
      szerzo: book.szerzo,
      termeknev: book.termeknev,
      fogy_ar: book.fogy_ar,
      kedvezmeny_szazalek: book.kedvezmeny_szazalek,
      kedvezmenyes_ar: book.kedvezmenyes_ar,
      print_house: book.print_house,
      eladva_datum: new Date().toISOString(),
    }).select().single()
    if (!error && data) setSoldList((prev) => [data, ...prev])
  }, [])

  const handleUndo = useCallback(async (id) => {
    const { error } = await supabase.from('sold_books').delete().eq('id', id)
    if (!error) setSoldList((prev) => prev.filter((i) => i.id !== id))
  }, [])

  const handleExport = useCallback((list, filterHouse) => {
    const headers = ['Termékkód', 'EAN', 'Szerző', 'Cím', 'Kiadó', 'Eredeti ár (Ft)', 'Kedvezmény %', 'Kedvezményes ár (Ft)', 'Eladás dátuma']
    const rows = list.map((item) => [
      item.termekkkod,
      item.ean,
      item.szerzo,
      item.termeknev,
      item.print_house ?? '',
      item.fogy_ar != null ? Math.round(item.fogy_ar) : '',
      item.kedvezmeny_szazalek != null ? Math.round(item.kedvezmeny_szazalek * 100) + '%' : '',
      item.kedvezmenyes_ar != null ? Math.round(item.kedvezmenyes_ar) : '',
      formatSoldDate(item.eladva_datum),
    ])
    const csvContent = [headers, ...rows]
      .map((row) => row.map((cell) => {
        const str = String(cell ?? '')
        return str.includes(',') || str.includes('"') || str.includes('\n')
          ? '"' + str.replace(/"/g, '""') + '"' : str
      }).join(',')).join('\n')
    const blob = new Blob(['﻿' + csvContent], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    const today = new Date().toISOString().split('T')[0]
    const housePart = filterHouse && filterHouse !== 'Összes'
      ? '_' + filterHouse.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '')
      : ''
    a.href = url
    a.download = `rebella_eladott_konyvek${housePart}_${today}.csv`
    a.click()
    URL.revokeObjectURL(url)
  }, [])

  const handleClear = useCallback(async () => {
    if (!window.confirm('Biztosan törölni szeretnéd az összes eladott könyvet?')) return
    const { error } = await supabase.from('sold_books').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    if (!error) setSoldList([])
  }, [])

  if (loading) {
    return (
      <div className="min-h-screen bg-[#FAFAF7] flex items-center justify-center">
        <p className="text-gray-500 text-lg">Adatok betöltése...</p>
      </div>
    )
  }

  if (error) {
    return (
      <div className="min-h-screen bg-[#FAFAF7] flex items-center justify-center px-4">
        <p className="text-red-600 text-lg text-center">{error}</p>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-[#FAFAF7]">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-3xl mx-auto px-4 py-4">
          <h1 className="text-xl font-bold text-[#C0392B]">a Rebella-ba Kft legnagyszerűbb fanti kis applikációja</h1>
          <nav className="mt-3 flex gap-1">
            <button
              onClick={() => setTab('search')}
              className={`px-4 py-2 text-sm font-medium rounded-lg transition-colors ${
                tab === 'search' ? 'bg-[#C0392B] text-white' : 'text-gray-600 hover:bg-gray-100'
              }`}
            >
              Keresés
            </button>
            <button
              onClick={() => setTab('sold')}
              className={`px-4 py-2 text-sm font-medium rounded-lg transition-colors ${
                tab === 'sold' ? 'bg-[#C0392B] text-white' : 'text-gray-600 hover:bg-gray-100'
              }`}
            >
              Eladott könyvek ({soldList.length})
            </button>
            <button
              onClick={() => setTab('ideas')}
              className={`px-4 py-2 text-sm font-medium rounded-lg transition-colors ${
                tab === 'ideas' ? 'bg-[#C0392B] text-white' : 'text-gray-600 hover:bg-gray-100'
              }`}
            >
              Ötletdoboz
            </button>
            <button
              onClick={() => setTab('admin')}
              className={`px-4 py-2 text-sm font-medium rounded-lg transition-colors ${
                tab === 'admin' ? 'bg-[#C0392B] text-white' : 'text-gray-600 hover:bg-gray-100'
              }`}
            >
              Admin
            </button>
          </nav>
        </div>
      </header>

      <main className="max-w-3xl mx-auto px-4 py-6">
        {tab === 'search' && (
          <div>
            <input
              type="text"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Keresés cím, szerző, termékkód vagy EAN alapján..."
              className="w-full px-4 py-3 rounded-lg border border-gray-300 bg-white text-gray-900
                placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#C0392B] focus:border-transparent"
            />
            <div className="mt-5 flex flex-col gap-4">
              {!debouncedQuery.trim() && (
                <div className="text-center py-16 px-4">
                  <p className="text-gray-300 text-xs uppercase tracking-widest mb-3">Tudtad?</p>
                  <p className="text-gray-400 text-base italic">„{randomFact()}"</p>
                  <p className="text-gray-300 text-sm mt-6">Kezdj el gépelni a kereséshez...</p>
                </div>
              )}
              {debouncedQuery.trim() && filtered.length === 0 && (
                <p className="text-center py-16 text-gray-400">Nem található könyv.</p>
              )}
              {filtered.map((book) => (
                <BookCard
                  key={book.id}
                  book={book}
                  soldCount={soldCounts.get(book.termekkkod) ?? 0}
                  onSell={handleSell}
                />
              ))}
            </div>
          </div>
        )}
        {tab === 'sold' && (
          <SoldTable
            soldList={soldList}
            onUndo={handleUndo}
            onExport={handleExport}
            onClear={handleClear}
          />
        )}
        {tab === 'ideas' && <IdeasTab />}
        {tab === 'admin' && <AdminPanel />}
      </main>
    </div>
  )
}
