import * as XLSX from 'xlsx'
import { createClient } from '@supabase/supabase-js'
import { readFileSync } from 'fs'

const supabase = createClient(
  'https://ceyzptovpnhfpdmwxugw.supabase.co',
  'sb_publishable_cgLWw_JuaqByAOx4snaTTA_NQ5PQaxP'
)

const buffer = readFileSync('./public/Rebella-Bakeszlet.xlsx')
const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true })
const ws = wb.Sheets['Aktuális készlet']
const rows = XLSX.utils.sheet_to_json(ws, { range: 1, defval: null })

console.log(`Found ${rows.length} books in Excel, uploading...`)

const books = rows.map(r => ({
  termekkkod:         String(r['Termékkód'] ?? ''),
  ean:                r['EAN'] != null ? String(r['EAN']) : null,
  szerzo:             r['szerző'] ?? null,
  termeknev:          r['Terméknév'] ?? null,
  akcio_keszlet:      r['akciós készlet'] ?? 0,
  normal_keszlet:     r['normál készlet'] ?? 0,
  fogy_ar:            r['fogy.ár (5% áfás)'] ?? null,
  netto_ar:           r['nettó ár'] ?? null,
  ob_netto_ar:        r['árréssel csökentett nettó ár (OB-nak fizetendő)'] ?? null,
  kedvezmeny_szazalek: r['javasolt kedvezmény vevőknek'] ?? null,
  kedvezmenyes_ar:    r['javasolt kedvezményes, bolti áfás ár'] ?? null,
  arkototteg_lejar:   r['árkötöttség lejár'] instanceof Date
                        ? r['árkötöttség lejár'].toISOString().split('T')[0]
                        : null,
  print_house:        'OpenBooks',
}))

const BATCH = 50
for (let i = 0; i < books.length; i += BATCH) {
  const batch = books.slice(i, i + BATCH)
  const { error } = await supabase.from('books').insert(batch)
  if (error) {
    console.error(`Error at batch ${i}:`, error.message)
    process.exit(1)
  }
  console.log(`Uploaded ${Math.min(i + BATCH, books.length)} / ${books.length}`)
}

console.log('Done! All books uploaded to Supabase.')
