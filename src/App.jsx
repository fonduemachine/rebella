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
  'EGYÉB',
]

function parseExcelBooks(buffer, printHouse) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true })
  const ws = wb.Sheets['Aktuális készlet']
  if (!ws) throw new Error('Nem található "Aktuális készlet" munkalap a fájlban.')
  const rows = XLSX.utils.sheet_to_json(ws, { range: 1, defval: null })
  return rows.map(r => ({
    termekkkod:          String(r['Termékkód'] ?? ''),
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
  }))
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
    </div>
  )
}

function ExcelUpload() {
  const [printHouse, setPrintHouse] = useState('OpenBooks')
  const [parsedBooks, setParsedBooks] = useState(null)
  const [parseError, setParseError] = useState(null)
  const [uploading, setUploading] = useState(false)
  const [uploadDone, setUploadDone] = useState(false)

  function handleFile(e) {
    const file = e.target.files[0]
    if (!file) return
    setParsedBooks(null)
    setParseError(null)
    setUploadDone(false)
    const reader = new FileReader()
    reader.onload = (ev) => {
      try {
        const books = parseExcelBooks(ev.target.result, printHouse.trim() || 'Ismeretlen')
        setParsedBooks(books)
      } catch (err) {
        setParseError(err.message)
      }
    }
    reader.readAsArrayBuffer(file)
  }

  async function handleUpload() {
    if (!parsedBooks || parsedBooks.length === 0) return
    setUploading(true)

    // Upsert in batches of 50 — adds new books, updates existing ones
    const BATCH = 50
    for (let i = 0; i < parsedBooks.length; i += BATCH) {
      const { error } = await supabase
        .from('books')
        .upsert(parsedBooks.slice(i, i + BATCH), { onConflict: 'termekkkod,print_house' })
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
      <h3 className="font-bold text-gray-900 mb-1">Excel feltöltés</h3>
      <p className="text-sm text-gray-500 mb-5">
        Tölts fel egy új készletlistát. Az adott kiadóhoz tartozó összes könyv frissül.
      </p>

      <div className="flex flex-col gap-4">
        {/* Print house */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Kiadó neve</label>
          <select
            value={printHouse}
            onChange={(e) => { setPrintHouse(e.target.value); setParsedBooks(null); setUploadDone(false) }}
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
          <label className="block text-sm font-medium text-gray-700 mb-1">Excel fájl (.xlsx)</label>
          <input
            id="excel-upload-input"
            type="file"
            accept=".xlsx"
            onChange={handleFile}
            className="w-full text-sm text-gray-600 file:mr-3 file:py-2 file:px-4
              file:rounded-lg file:border-0 file:text-sm file:font-medium
              file:bg-[#C0392B] file:text-white hover:file:bg-[#A93226] cursor-pointer"
          />
        </div>

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
    if (!form.termekkkod.trim()) { setError('A termékkód kötelező.'); return }
    setSaving(true)
    const { error } = await supabase.from('books').upsert({
      termekkkod:          form.termekkkod.trim(),
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
            <label className={labelClass}>Termékkód *</label>
            <input type="text" value={form.termekkkod} onChange={e => set('termekkkod', e.target.value)} className={inputClass} placeholder="pl. O15724734" />
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

function formatPrice(value) {
  if (value == null || isNaN(value)) return '—'
  return Math.round(value).toLocaleString('hu-HU') + ' Ft'
}

function formatDiscount(value) {
  if (value == null || isNaN(value)) return '—'
  return Math.round(value * 100) + '%'
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
          Eladva
        </button>
        {soldCount > 0 && (
          <span className="text-sm text-gray-500">{soldCount} db eladva</span>
        )}
      </div>
    </div>
  )
}

function SoldTable({ soldList, onUndo, onExport, onClear }) {
  if (soldList.length === 0) {
    return (
      <div className="text-center py-16 text-gray-400">
        Még nem adtál el könyvet.
      </div>
    )
  }

  return (
    <div>
      <div className="flex flex-wrap gap-3 mb-5">
        <button
          onClick={onExport}
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

      {/* Mobile cards */}
      <div className="sm:hidden flex flex-col gap-3">
        {soldList.map((item) => (
          <div key={item.id} className="bg-white rounded-lg shadow-sm border border-gray-100 p-4">
            <p className="font-bold text-gray-900">{item.termeknev}</p>
            <p className="text-sm text-gray-600">{item.szerzo}</p>
            <div className="mt-2 flex flex-wrap gap-x-4 gap-y-1 text-sm text-gray-700">
              <span>Eredeti: {formatPrice(item.fogy_ar)}</span>
              <span>Kedv.: {Math.round((item.kedvezmeny_szazalek ?? 0) * 100)}%</span>
              <span>Ár: {formatPrice(item.kedvezmenyes_ar)}</span>
            </div>
            <p className="text-xs text-gray-400 mt-1">{formatSoldDate(item.eladva_datum)}</p>
            <button onClick={() => onUndo(item.id)} className="mt-2 text-sm text-red-600 hover:text-red-800">
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
            {soldList.map((item) => (
              <tr key={item.id} className="border-b border-gray-100 hover:bg-gray-50">
                <td className="py-3 pr-4 font-medium text-gray-900">{item.termeknev}</td>
                <td className="py-3 pr-4 text-gray-600">{item.szerzo}</td>
                <td className="py-3 pr-4 text-right text-gray-700">{formatPrice(item.fogy_ar)}</td>
                <td className="py-3 pr-4 text-right text-gray-700">{Math.round((item.kedvezmeny_szazalek ?? 0) * 100)}%</td>
                <td className="py-3 pr-4 text-right font-medium text-gray-900">{formatPrice(item.kedvezmenyes_ar)}</td>
                <td className="py-3 pr-4 text-gray-500">{formatSoldDate(item.eladva_datum)}</td>
                <td className="py-3">
                  <button onClick={() => onUndo(item.id)} className="text-red-600 hover:text-red-800">
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
      eladva_datum: new Date().toISOString(),
    }).select().single()
    if (!error && data) setSoldList((prev) => [data, ...prev])
  }, [])

  const handleUndo = useCallback(async (id) => {
    const { error } = await supabase.from('sold_books').delete().eq('id', id)
    if (!error) setSoldList((prev) => prev.filter((i) => i.id !== id))
  }, [])

  const handleExport = useCallback(() => {
    const headers = ['Termékkód', 'EAN', 'Szerző', 'Cím', 'Eredeti ár (Ft)', 'Kedvezmény %', 'Kedvezményes ár (Ft)', 'Eladás dátuma']
    const rows = soldList.map((item) => [
      item.termekkkod,
      item.ean,
      item.szerzo,
      item.termeknev,
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
    a.href = url
    a.download = `rebella_eladott_konyvek_${new Date().toISOString().split('T')[0]}.csv`
    a.click()
    URL.revokeObjectURL(url)
  }, [soldList])

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
          <h1 className="text-xl font-bold text-[#C0392B]">Rebella Könyvkészlet</h1>
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
                <p className="text-center py-16 text-gray-400">Kezdj el gépelni egy könyv címét...</p>
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
        {tab === 'admin' && <AdminPanel />}
      </main>
    </div>
  )
}
