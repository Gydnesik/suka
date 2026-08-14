import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { unzipSync } from 'npm:fflate@0.8.2'

const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
}

const supabaseUrl = Deno.env.get('SUPABASE_URL')!
const supabaseAnonKey = Deno.env.get('SUPABASE_PUBLISHABLE_KEY') || Deno.env.get('SUPABASE_ANON_KEY')!
const geminiKey = Deno.env.get('GEMINI_API_KEY')!
const geminiModel = Deno.env.get('GEMINI_MODEL') || 'gemini-3.6-flash'

const DAY_ORDER = ['Понедельник','Вторник','Среда','Четверг','Пятница','Суббота','Воскресенье']
const LATIN_TO_CYR = new Map([
  ['A','А'],['B','В'],['C','С'],['E','Е'],['H','Н'],['K','К'],['M','М'],['O','О'],['P','Р'],['T','Т'],['X','Х'],['Y','У'],
])

function normalizeClassName(value: unknown): string | null {
  let s = String(value ?? '').trim().toUpperCase()
  s = s.replace(/\bКЛАСС\b/giu, '').replace(/[\s._\-–—:№]+/g, '')
  for (const [latin, cyr] of LATIN_TO_CYR) s = s.replaceAll(latin, cyr)
  const match = s.match(/^(5|6|7|8|9|10|11)([АБВГ])$/u)
  return match ? `${match[1]}${match[2]}` : null
}

function normalizeDay(value: unknown): string {
  const raw = String(value ?? '').trim().toLowerCase().replace(/ё/g, 'е')
  const aliases: Record<string, string> = {
    'пн':'Понедельник','пон':'Понедельник','понедельник':'Понедельник',
    'вт':'Вторник','вторник':'Вторник','ср':'Среда','среда':'Среда',
    'чт':'Четверг','четверг':'Четверг','пт':'Пятница','пятница':'Пятница',
    'сб':'Суббота','суббота':'Суббота','вс':'Воскресенье','воскресенье':'Воскресенье',
  }
  return aliases[raw] || String(value ?? '').trim() || 'Понедельник'
}

function normalizeLesson(raw: any, fallbackIndex: number): any | null {
  if (!raw || typeof raw !== 'object') return null
  const subject = String(raw.subject ?? raw.name ?? raw.title ?? '').trim()
  if (!subject) return null
  const lessonRaw = Number(raw.lesson ?? raw.number ?? fallbackIndex + 1)
  const lesson = Number.isFinite(lessonRaw) && lessonRaw > 0 ? Math.floor(lessonRaw) : fallbackIndex + 1
  return {
    day: normalizeDay(raw.day), lesson, subject,
    time: String(raw.time ?? '').trim(),
    room: String(raw.room ?? raw.cabinet ?? raw.classroom ?? '').trim(),
  }
}

function normalizeLessons(raw: unknown): any[] {
  if (!Array.isArray(raw)) return []
  const result: any[] = []
  raw.forEach((item, i) => {
    const lesson = normalizeLesson(item, i)
    if (lesson) result.push(lesson)
  })
  result.sort((a,b) => {
    const d = DAY_ORDER.indexOf(a.day) - DAY_ORDER.indexOf(b.day)
    return d || a.lesson - b.lesson || a.subject.localeCompare(b.subject, 'ru')
  })
  const seen = new Set<string>()
  return result.filter(x => {
    const key = `${x.day}|${x.lesson}|${x.subject}|${x.time}|${x.room}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

function normalizeSchedules(input: any) {
  const out: Record<string, any[]> = {}
  const ignoredClasses: string[] = []
  const source = Array.isArray(input?.schedules)
    ? input.schedules
    : input?.schedules && typeof input.schedules === 'object'
      ? Object.entries(input.schedules).map(([class_name, lessons]) => ({ class_name, lessons }))
      : Array.isArray(input) ? input : []

  for (const item of source) {
    const rawClass = item?.class_name ?? item?.className ?? item?.class ?? item?.name
    const className = normalizeClassName(rawClass)
    if (!className) {
      if (String(rawClass ?? '').trim()) ignoredClasses.push(String(rawClass).trim())
      continue
    }
    const lessons = normalizeLessons(item?.lessons)
    if (!out[className]) out[className] = []
    out[className].push(...lessons)
  }
  for (const className of Object.keys(out)) out[className] = normalizeLessons(out[className])
  return { schedules: out, ignoredClasses: [...new Set(ignoredClasses)] }
}

const responseSchema = {
  type: 'object',
  properties: {
    schedules: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          class_name: { type: 'string' },
          lessons: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                day: { type: 'string' }, lesson: { type: 'integer' },
                subject: { type: 'string' }, time: { type: 'string' }, room: { type: 'string' },
              },
              required: ['day','lesson','subject','time','room'],
            },
          },
        },
        required: ['class_name','lessons'],
      },
    },
  },
  required: ['schedules'],
}

const SUPPORTED_EXTENSIONS = new Set(['xlsx','ods','csv'])
const MAX_UPLOAD_BYTES = 12 * 1024 * 1024
const MAX_SHEET_ROWS = 2500
const MAX_SHEET_COLS = 80
const MAX_XML_BYTES = 30 * 1024 * 1024

async function readUpload(req: Request): Promise<{ bytes: Uint8Array, fileName: string }> {
  const contentType = req.headers.get('content-type') || ''
  if (contentType.toLowerCase().startsWith('multipart/form-data')) {
    const form = await req.formData()
    const file = form.get('file')
    if (!(file instanceof File)) throw new Error('Файл не передан.')
    if (file.size > MAX_UPLOAD_BYTES) throw new Error('Файл слишком большой. Максимальный размер — 12 МБ.')
    return { bytes: new Uint8Array(await file.arrayBuffer()), fileName: file.name || 'schedule.xlsx' }
  }
  throw new Error('Ожидался файл multipart/form-data. Обнови страницу и попробуй снова.')
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return ''
  return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim()
}

function scoreScheduleSheet(rows: string[][]): number {
  let classHits = 0, dayHits = 0, timeHits = 0, lessonHits = 0, nonEmpty = 0
  const classRe = /\b(?:5|6|7|8|9|10|11)\s*[АБВГAB]\b/giu
  const dayRe = /\b(?:понедельник|вторник|среда|четверг|пятница|суббота)\b/giu
  const timeRe = /\b\d{1,2}[:.]\d{2}\s*[-–—]\s*\d{1,2}[:.]\d{2}\b/g
  for (const row of rows) {
    for (const cell of row) {
      if (!cell) continue
      nonEmpty++
      classRe.lastIndex=0; dayRe.lastIndex=0; timeRe.lastIndex=0
      classHits += (cell.match(classRe)||[]).length
      dayHits += (cell.match(dayRe)||[]).length
      timeHits += (cell.match(timeRe)||[]).length
    }
    for (const cell of row.slice(0,4)) if (/^\d{1,2}$/.test(cell.trim())) lessonHits++
  }
  return classHits*20 + dayHits*8 + timeHits*6 + lessonHits*2 + Math.min(nonEmpty,500)/100
}

function csvEscape(value: string): string {
  const s = normalizeCell(value)
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s
}


function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => {
      try { return String.fromCodePoint(parseInt(hex, 16)) } catch { return '' }
    })
    .replace(/&#([0-9]+);/g, (_, num) => {
      try { return String.fromCodePoint(Number(num)) } catch { return '' }
    })
    .replace(/&amp;/g, '&')
}

function xmlAttr(tag: string, local: string, prefix?: string): string {
  const name = prefix ? `${prefix}:${local}` : local
  const re = new RegExp(`(?:^|\\s)${name.replace(':','\\:')}\\s*=\\s*([\"'])(.*?)\\1`, 'i')
  const m = tag.match(re)
  return m ? decodeXmlEntities(m[2]) : ''
}

function xmlTextFragment(fragment: string): string {
  let s = fragment
    .replace(/<(?:(?:text|office):)?tab\b[^>]*\/?>/gi, '\t')
    .replace(/<(?:(?:text|office):)?line-break\b[^>]*\/?>/gi, '\n')
    .replace(/<\/(?:(?:text|office):)?p\s*>/gi, '\n')
    .replace(/<\/(?:t|text:p|text:span)\s*>/gi, ' ')
  s = s.replace(/<[^>]*>/g, '')
  return decodeXmlEntities(s).replace(/[ \t\f\v]+/g, ' ').replace(/\s*\n\s*/g, '\n').trim()
}

function xmlBlocks(xml: string, localName: string): string[] {
  const escaped = localName.replace(':', '\\:')
  const re = new RegExp(`<${escaped}\\b[^>]*>[\\s\\S]*?<\\/${escaped}>`, 'gi')
  return xml.match(re) || []
}

function xmlOpenTags(xml: string, localName: string): string[] {
  const escaped = localName.replace(':', '\\:')
  const re = new RegExp(`<${escaped}\\b[^>]*\\/?>`, 'gi')
  return xml.match(re) || []
}

function zipPathNormalize(path: string): string {
  const parts: string[] = []
  for (const part of path.replace(/\\/g, '/').split('/')) {
    if (!part || part === '.') continue
    if (part === '..') parts.pop()
    else parts.push(part)
  }
  return parts.join('/')
}

function resolveZipTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return zipPathNormalize(target.slice(1))
  return zipPathNormalize(`${baseDir}/${target}`)
}

function columnIndex(ref: string): number {
  const m = ref.match(/^([A-Z]+)\d+$/i)
  if (!m) return -1
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return col - 1
}

function parseOdsSheets(bytes: Uint8Array): Array<{name:string, rows:string[][]}> {
  const files = unzipSync(bytes)
  const contentBytes = files['content.xml']
  if (!contentBytes) throw new Error('ODS-файл не содержит content.xml.')
  if (contentBytes.byteLength > MAX_XML_BYTES) throw new Error('ODS-файл после распаковки слишком большой.')

  const xml = new TextDecoder().decode(contentBytes)
  const result: Array<{name:string, rows:string[][]}> = []

  for (const tableXml of xmlBlocks(xml, 'table:table')) {
    const opening = tableXml.match(/^<table:table\b[^>]*>/i)?.[0] || ''
    const name = xmlAttr(opening, 'name', 'table') || 'Лист'
    const rows: string[][] = []
    const merges: Array<{sr:number,er:number,sc:number,ec:number,value:string}> = []
    let r = 0

    for (const rowXml of xmlBlocks(tableXml, 'table:table-row')) {
      let repeatRows = Math.max(1, Number(xmlAttr(rowXml.match(/^<table:table-row\b[^>]*>/i)?.[0] || '', 'number-rows-repeated', 'table')) || 1)
      repeatRows = Math.min(repeatRows, MAX_SHEET_ROWS - r)
      if (repeatRows <= 0) break

      const cellXmls = [
        ...xmlBlocks(rowXml, 'table:table-cell'),
        ...xmlOpenTags(rowXml, 'table:covered-table-cell'),
      ]
      // Keep the original XML order. The simple combined regexes above can reorder
      // covered cells, so use one order-preserving matcher instead.
      const orderedCellRe = /<(?:table:)?(?:table-cell|covered-table-cell)\b[^>]*(?:\/>|>[\s\S]*?<\/(?:table:)?(?:table-cell|covered-table-cell)>)/gi
      const ordered = rowXml.match(orderedCellRe) || []
      const cells = ordered.length ? ordered : cellXmls

      for (let rr = 0; rr < repeatRows; rr++) {
        const row = new Array<string>(MAX_SHEET_COLS).fill('')
        let c = 0

        for (const cellXml of cells) {
          if (c >= MAX_SHEET_COLS) break
          const openingCell = cellXml.match(/^<[^>]+>/)?.[0] || ''
          const covered = /^<(?:(?:table:)?covered-table-cell)\b/i.test(openingCell)
          const repeat = Math.min(MAX_SHEET_COLS - c, Math.max(1, Number(xmlAttr(openingCell, 'number-columns-repeated', 'table')) || 1))
          const colSpan = Math.min(MAX_SHEET_COLS - c, Math.max(1, Number(xmlAttr(openingCell, 'number-columns-spanned', 'table')) || 1))
          const rowSpan = Math.max(1, Number(xmlAttr(openingCell, 'number-rows-spanned', 'table')) || 1)

          let value = ''
          if (!covered) {
            value = xmlAttr(openingCell, 'string-value', 'office') ||
              xmlAttr(openingCell, 'value', 'office') ||
              xmlTextFragment(cellXml)
          }
          value = normalizeCell(value)

          for (let k = 0; k < repeat; k++) {
            row[c + k] = value
            if (value && (colSpan > 1 || rowSpan > 1)) {
              merges.push({
                sr: r,
                er: Math.min(MAX_SHEET_ROWS - 1, r + rowSpan - 1),
                sc: c + k,
                ec: Math.min(MAX_SHEET_COLS - 1, c + k + colSpan - 1),
                value,
              })
            }
          }
          c += repeat
        }

        rows[r] = row
        r++
      }
    }

    for (const m of merges) {
      for (let rr = m.sr; rr <= m.er && rr < rows.length; rr++) {
        for (let cc = m.sc; cc <= m.ec; cc++) rows[rr][cc] = m.value
      }
    }

    while (rows.length && rows[rows.length - 1].every(x => !x)) rows.pop()
    if (rows.length) result.push({ name, rows })
  }

  return result
}

function parseXlsxSheets(bytes: Uint8Array): Array<{name:string, rows:string[][]}> {
  const files = unzipSync(bytes)
  const wbBytes = files['xl/workbook.xml']
  if (!wbBytes) throw new Error('XLSX-файл не содержит xl/workbook.xml.')

  const shared: string[] = []
  const ss = files['xl/sharedStrings.xml']
  if (ss) {
    const ssXml = new TextDecoder().decode(ss)
    for (const si of xmlBlocks(ssXml, 'si')) shared.push(xmlTextFragment(si))
  }

  const wbXml = new TextDecoder().decode(wbBytes)
  const relBytes = files['xl/_rels/workbook.xml.rels']
  const rels: Record<string, string> = {}
  if (relBytes) {
    const relXml = new TextDecoder().decode(relBytes)
    for (const tag of xmlOpenTags(relXml, 'Relationship')) {
      const id = xmlAttr(tag, 'Id')
      const target = xmlAttr(tag, 'Target')
      if (id && target) rels[id] = resolveZipTarget('xl', target)
    }
  }

  const out: Array<{name:string, rows:string[][]}> = []
  for (const shTag of xmlOpenTags(wbXml, 'sheet')) {
    const rid = xmlAttr(shTag, 'id', 'r') || xmlAttr(shTag, 'id')
    const target = rels[rid]
    if (!target) continue
    const sheetBytes = files[target]
    if (!sheetBytes) continue

    const sheetXml = new TextDecoder().decode(sheetBytes)
    const rows: string[][] = []

    for (const rowXml of xmlBlocks(sheetXml, 'row')) {
      const rowOpen = rowXml.match(/^<row\b[^>]*>/i)?.[0] || ''
      const rrRaw = Number(xmlAttr(rowOpen, 'r'))
      const rr = Number.isFinite(rrRaw) && rrRaw > 0 ? rrRaw - 1 : rows.length
      if (rr < 0 || rr >= MAX_SHEET_ROWS) continue
      if (!rows[rr]) rows[rr] = new Array<string>(MAX_SHEET_COLS).fill('')

      const cells = rowXml.match(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/gi) || []
      for (const cellXml of cells) {
        const open = cellXml.match(/^<c\b[^>]*>/i)?.[0] || ''
        const ref = xmlAttr(open, 'r')
        const col = columnIndex(ref)
        if (col < 0 || col >= MAX_SHEET_COLS) continue

        const type = xmlAttr(open, 't')
        const vMatch = cellXml.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)
        const isMatch = cellXml.match(/<is\b[^>]*>([\s\S]*?)<\/is>/i)
        let value = ''
        if (type === 's' && vMatch) {
          value = shared[Number(decodeXmlEntities(vMatch[1]).trim())] || ''
        } else if (type === 'inlineStr' && isMatch) {
          value = xmlTextFragment(isMatch[1])
        } else if (vMatch) {
          value = decodeXmlEntities(vMatch[1]).trim()
        }
        rows[rr][col] = normalizeCell(value)
      }
    }

    while (rows.length && (!rows[rows.length - 1] || rows[rows.length - 1].every(x => !x))) rows.pop()
    if (rows.length) out.push({ name: xmlAttr(shTag, 'name') || 'Лист', rows })
  }
  return out
}

function workbookToCleanCsv(bytes: Uint8Array, extension: string): string {
  let sheets:Array<{name:string,rows:string[][]}>
  if(extension==='ods') sheets=parseOdsSheets(bytes)
  else if(extension==='xlsx') sheets=parseXlsxSheets(bytes)
  else if(extension==='csv') sheets=parseCsv(bytes)
  else throw new Error('Формат .xls не поддерживается в новой экономной обработке. Сохрани файл как .xlsx или .ods и загрузи его снова.')

  let bestName='', bestCsv='', bestScore=-1
  for(const sheet of sheets){
    const score=scoreScheduleSheet(sheet.rows)
    if(score<=bestScore) continue
    const csv=sheet.rows.map(row=>row.map(csvEscape).join(',')).join('\n')
    if(!csv.trim()) continue
    bestName=sheet.name; bestScore=score; bestCsv=csv
  }
  if(!bestCsv.trim()) throw new Error('В таблице не найдено ни одного непустого листа.')
  return `ЛИСТ: ${bestName}\n\n${bestCsv}`
}

const SCHEDULE_PROMPT = `Ты — строгий парсер школьного расписания. На входе НЕ фотография и НЕ PDF, а очищенный CSV-текст электронной таблицы. Твоя задача — восстановить только основную сетку расписания для классов 5А–11Б.

КЛЮЧЕВОЕ ПРАВИЛО: В таблице могут быть служебные блоки сверху и снизу: дежурство, ВПР, замены, объявления, примечания, подписи, даты, телефоны и прочая мета-информация. ПОЛНОСТЬЮ ИГНОРИРУЙ их. Ищи ТОЛЬКО НИЖНЮЮ ОСНОВНУЮ СЕТКУ РАСПИСАНИЯ. Она может начинаться не с первой строки — например, с 17-й, 20-й или другой строки. Начало сетки определяется по сочетанию заголовков классов, дней/номеров уроков, времени и предметов.

1. КЛАССЫ
- Ищи и обрабатывай классы от 5А до 11Б включительно. Если в таблице есть 5В, 10Г и т.п., НЕ добавляй их: в этом задании нужны только 5А–11Б.
- Нормализуй 8а/8А/8 а/8-А/8A в 8А; 8Б сохраняй как 8Б.
- Латинская B в заголовке может означать кириллическую В, но не превращай явную Б в В.
- Не придумывай класс, которого нет в таблице.

2. ГРАНИЦЫ ОСНОВНОЙ СЕТКИ
- Сначала найди строку/строки заголовков классов.
- Затем найди основную сетку под этими заголовками и читай её ДО КОНЦА ВНИЗ.
- Верхние служебные блоки, даже если там встречаются названия классов, не считать расписанием.
- Нижние примечания/дежурство/ВПР после основной сетки не считать расписанием.
- Особое внимание строкам 6, 7, 8 и 9 уроков и вообще последним строкам основной сетки: они не должны потеряться.

3. ЧТЕНИЕ СТРОГО ПО ВЕРТИКАЛИ
- Для КАЖДОГО найденного класса читай его столбец строго сверху вниз, строка за строкой.
- Не смешивай соседние столбцы классов.
- Если расписание разделено на дни, переноси день вниз на все строки соответствующего блока до следующего дня.
- Номер урока и время бери из ЛЕВОЙ части той же строки. Сохраняй фактический номер урока и фактическое время, не заменяй их индексом массива.
- Если номер урока или время видны, обязательно сохрани их.
- Если в ячейке стоит ровно прочерк "-", запиши subject как "-". Не превращай его в пустую строку и не удаляй такой урок.
- Пустая ячейка означает отсутствие урока, КРОМЕ СЛУЧАЯ ОБЪЕДИНЁННОЙ ЯЧЕЙКИ.

4. ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ И ПУСТОТЫ
- Excel/ODS может представить объединённую ячейку как значение только в верхней левой клетке, а остальные клетки той же объединённой области будут пустыми. При конвертации часть таких значений уже продублирована по диапазону объединения, но ты всё равно обязан учитывать исходную логику таблицы.
- Если видно, что один урок/предмет относится сразу к нескольким классам или группам, полностью продублируй этот текст в каждой соответствующей колонке класса.
- Если значение стоит в верхней клетке вертикально объединённого блока и ниже идут пустые клетки в том же логическом блоке, наследуй значение вниз ТОЛЬКО если структура таблицы явно показывает объединение/общий урок. Не заполняй обычные случайные пустоты догадками.
- Никогда не теряй урок из-за того, что его ячейка после конвертации CSV стала пустой вследствие объединения.

5. ПРЕДМЕТ И КАБИНЕТ
- subject должен содержать название предмета/текст урока так, как он указан в ячейке, включая сокращения.
- room должен содержать кабинет, если он указан отдельно или явно читается в той же ячейке. Не выдумывай кабинет.
- Если предмет и кабинет записаны вместе, раздели их разумно: название предмета в subject, номер/обозначение кабинета в room.
- Если предмет реально отсутствует, не выдумывай его.

6. JSON
- Верни ТОЛЬКО JSON по заданной схеме.
- В JSON не должно быть OCR-комментариев, мета-информации, названий файлов, confidence, source, notes и других полей.
- Каждый class_name содержит только один канонический класс.
- lessons содержит только day, lesson, subject, time, room. День нужен приложению для правильного показа расписания по дням.
- Не добавляй уроки из служебных блоков.
- Лучше оставить реально пустое поле пустым, чем придумать значение.

ПЕРЕД ОТВЕТОМ СДЕЛАЙ КОНТРОЛЬ:
а) проверь каждый класс 5А–11Б отдельно;
б) проверь основную сетку от её первой строки до последней;
в) отдельно перепроверь нижнюю часть таблицы;
г) убедись, что прочерки "-" сохранены;
д) убедись, что объединённые уроки продублированы для всех нужных классов;
е) убедись, что старые/служебные блоки не попали в JSON.`

async function callGemini(textPrompt: string) {
  const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(geminiModel)}:generateContent`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': geminiKey },
    body: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: textPrompt }] }],
      generationConfig: {
        responseMimeType: 'application/json',
        responseSchema,
        temperature: 0,
      },
    }),
  })
  const data = await r.json()
  if (!r.ok) throw new Error(data?.error?.message || `Gemini HTTP ${r.status}`)
  let raw = ''
  for (const p of data?.candidates?.[0]?.content?.parts || []) if (p.text) raw += p.text
  raw = raw.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim()
  if (!raw) throw new Error('Gemini не вернул JSON.')
  try { return JSON.parse(raw) } catch { throw new Error('Gemini вернул невалидный JSON.') }
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: cors })
  if (req.method !== 'POST') return json({ error: 'Method not allowed' }, 405)

  try {
    if (!geminiKey) return json({ error: 'На сервере не задан GEMINI_API_KEY.' }, 500)

    const auth = req.headers.get('Authorization') || ''
    const supabase = createClient(supabaseUrl, supabaseAnonKey, {
      global: { headers: { Authorization: auth } },
    })
    const { data: { user } } = await supabase.auth.getUser()
    if (!user) return json({ error: 'Необходим вход в аккаунт.' }, 401)

    const { data: profile, error: profileError } = await supabase
      .from('profiles').select('role').eq('id', user.id).single()
    if (profileError || profile?.role !== 'admin') return json({ error: 'Доступ только для администратора.' }, 403)

    let upload: { bytes: Uint8Array, fileName: string }
    try {
      upload = await readUpload(req)
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err)
      return json({ error: message }, 413)
    }

    const extension = (upload.fileName.split('.').pop() || '').toLowerCase()
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      return json({ error: 'Поддерживаются только .xlsx, .ods и .csv.' }, 400)
    }

    let csvText: string
    try {
      csvText = workbookToCleanCsv(upload.bytes, extension)
      // После преобразования таблицы исходный бинарный буфер больше не нужен.
      upload.bytes = new Uint8Array(0)
    } catch (err) {
      return json({ error: err instanceof Error ? err.message : String(err) }, 422)
    }

    if (csvText.length > 180_000) {
      return json({ error: 'После очистки таблица слишком большая для обработки одной операцией. Удали лишние листы/служебные данные и загрузи только лист с расписанием.' }, 413)
    }

    const prompt = `${SCHEDULE_PROMPT}

ОЧИЩЕННЫЙ CSV:
${csvText}`

    let firstPass: any
    try {
      firstPass = await callGemini(prompt)
    } catch (err) {
      return json({ error: err instanceof Error ? err.message : String(err) }, 502)
    }

    const normalized = normalizeSchedules(firstPass)
    const entries = Object.entries(normalized.schedules)
    if (!entries.length) {
      return json({
        error: 'Не удалось распознать ни одного поддерживаемого класса 5А–11Б.',
        ignoredClasses: normalized.ignoredClasses,
      }, 422)
    }

    // Важное изменение: сохранение выполняется одной атомарной SQL-операцией.
    // Она СНАЧАЛА удаляет ВСЕ старые строки schedules, затем вставляет только
    // результат текущего файла. Поэтому старые дни, классы и алиасы никогда
    // не смешиваются с новым расписанием.
    const { error: replaceError } = await supabase.rpc('replace_schedules', {
      p_schedules: normalized.schedules,
    })
    if (replaceError) {
      return json({ error: `Не удалось заменить расписание в базе: ${replaceError.message}` }, 500)
    }

    return json({
      schedules: normalized.schedules,
      ignoredClasses: normalized.ignoredClasses,
      model: geminiModel,
    })
  } catch (err) {
    return json({ error: err instanceof Error ? err.message : String(err) }, 500)
  }
})

function json(data: unknown, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...cors, 'Content-Type': 'application/json' },
  })
}
