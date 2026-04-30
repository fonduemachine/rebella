import { useState, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from './supabase'

function formatPrice(value) {
  if (value == null || isNaN(value)) return '—'
  return Math.round(value).toLocaleString('hu-HU') + ' Ft'
}

function formatDiscount(value) {
  if (value == null || isNaN(value)) return '—'
  return Math.round(value * 100) + '%'
}

function formatDate(d) {
  if (!d) return null
  const date = d instanceof Date ? d : new Date(d)
  if (isNaN(date.getTime())) return null
  return date.toISOString().split('T')[0]
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
  const normalStock = book['normál készlet'] ?? 0
  const akcioStock = book['akciós készlet'] ?? 0
  const normalRemaining = Math.max(0, normalStock - soldCount)
  const akcioRemaining = soldCount > normalStock
    ? Math.max(0, akcioStock - (soldCount - normalStock))
    : akcioStock
  const totalRemaining = normalRemaining + akcioRemaining
  const stock = getStockStatus(normalRemaining, akcioRemaining)
  const priceExpiry = formatDate(book['árkötöttség lejár'])

  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-100 p-4 sm:p-5">
      <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-3">
        <div className="flex-1 min-w-0">
          <h3 className="text-lg font-bold text-gray-900 leading-tight">
            {book['Terméknév'] || 'Ismeretlen cím'}
          </h3>
          <p className="text-gray-600 mt-1">{book['szerző'] || 'Ismeretlen szerző'}</p>
          <p className="text-xs text-gray-400 mt-1">
            Kód: {book['Termékkód']} | EAN: {String(book['EAN'] ?? '')}
          </p>
        </div>
        <div className="flex flex-col items-start sm:items-end gap-2 shrink-0">
          <span className={`inline-block text-xs font-medium px-2.5 py-1 rounded-full ${stock.color}`}>
            {stock.label}
          </span>
          {book['javasolt kedvezmény vevőknek'] != null && (
            <span className="inline-block text-xs font-medium px-2.5 py-1 rounded-full bg-amber-100 text-amber-800">
              {formatDiscount(book['javasolt kedvezmény vevőknek'])} kedvezmény
            </span>
          )}
        </div>
      </div>

      <div className="mt-4 flex flex-wrap items-end gap-x-6 gap-y-2">
        <div>
          <p className="text-xs text-gray-500">Eredeti ár</p>
          <p className="text-gray-700 line-through">
            {formatPrice(book['fogy.ár (5% áfás)'])}
          </p>
        </div>
        <div>
          <p className="text-xs text-gray-500">Kedvezményes ár</p>
          <p className="text-xl font-semibold text-[#C0392B]">
            {formatPrice(book['javasolt kedvezményes, bolti áfás ár'])}
          </p>
        </div>
      </div>

      {priceExpiry && (
        <p className="mt-3 text-sm text-amber-700">
          ⚠️ Árkötöttség lejár: {priceExpiry}
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
            <button
              onClick={() => onUndo(item.id)}
              className="mt-2 text-sm text-red-600 hover:text-red-800"
            >
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
                  <button
                    onClick={() => onUndo(item.id)}
                    className="text-red-600 hover:text-red-800"
                  >
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
  const [dbError, setDbError] = useState(null)
  const [error, setError] = useState(null)
  const [query, setQuery] = useState('')
  const [tab, setTab] = useState('search')
  const [soldList, setSoldList] = useState([])

  const debouncedQuery = useDebounce(query, 300)

  // Load Excel
  useEffect(() => {
    async function loadData() {
      try {
        const response = await fetch('/Rebella-Bakeszlet.xlsx')
        const buffer = await response.arrayBuffer()
        const wb = XLSX.read(buffer, { type: 'array', cellDates: true })
        const ws = wb.Sheets['Aktuális készlet']
        if (!ws) { setError('A munkalap ("Aktuális készlet") nem található.'); return }
        setBooks(XLSX.utils.sheet_to_json(ws, { range: 1, defval: null }))
      } catch (e) {
        setError('Hiba az adatok betöltésekor: ' + e.message)
      }
    }
    loadData()
  }, [])

  // Load sold list from Supabase
  useEffect(() => {
    async function fetchSold() {
      const { data, error } = await supabase
        .from('sold_books')
        .select('*')
        .order('eladva_datum', { ascending: false })
      if (error) {
        setDbError(error.message)
      } else {
        setSoldList(data)
      }
      setLoading(false)
    }
    fetchSold()
  }, [])

  // Real-time sync across devices
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
    return books.filter((b) => {
      const title = (b['Terméknév'] || '').toLowerCase()
      const author = (b['szerző'] || '').toLowerCase()
      const code = String(b['Termékkód'] || '').toLowerCase()
      const ean = String(b['EAN'] || '').toLowerCase()
      return title.includes(q) || author.includes(q) || code.includes(q) || ean.includes(q)
    })
  }, [debouncedQuery, books])

  const handleSell = useCallback(async (book) => {
    const { data, error } = await supabase.from('sold_books').insert({
      termekkkod: book['Termékkód'],
      ean: String(book['EAN'] ?? ''),
      szerzo: book['szerző'],
      termeknev: book['Terméknév'],
      fogy_ar: book['fogy.ár (5% áfás)'],
      kedvezmeny_szazalek: book['javasolt kedvezmény vevőknek'],
      kedvezmenyes_ar: book['javasolt kedvezményes, bolti áfás ár'],
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
          ? '"' + str.replace(/"/g, '""') + '"'
          : str
      }).join(','))
      .join('\n')
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

  if (error || dbError) {
    return (
      <div className="min-h-screen bg-[#FAFAF7] flex items-center justify-center px-4">
        <div className="text-center">
          <p className="text-red-600 text-lg">{error || dbError}</p>
          {dbError && (
            <p className="text-gray-500 text-sm mt-2">
              Ellenőrizd a Supabase beállításokat és futtasd le a tábla-létrehozó SQL-t.
            </p>
          )}
        </div>
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
              {filtered.map((book, i) => (
                <BookCard
                  key={book['Termékkód'] + '-' + i}
                  book={book}
                  soldCount={soldCounts.get(book['Termékkód']) ?? 0}
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
      </main>
    </div>
  )
}
