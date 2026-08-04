import { useEffect, useMemo, useState } from 'react'
import { useRef } from 'react'
import {
  AlertTriangle,
  ArrowDown,
  ArrowUpDown,
  ArrowUp,
  BarChart3,
  Bell,
  Download,
  Edit3,
  FileSpreadsheet,
  Plus,
  RotateCcw,
  Save,
  Search,
  Trash2,
  Upload,
} from 'lucide-react'
import {
  Area,
  AreaChart,
  Bar,
  BarChart as RechartsBarChart,
  CartesianGrid,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import * as XLSX from 'xlsx'
import seedData from './data/seedData.json'
import './App.css'

type AnyRow = Record<string, unknown>

type MonitorRow = {
  owner: string
  group: string
  account: string
  sku: string
  bundleSku: string
  asinType: string
  asin: string
  note: string
}

type MappingRow = {
  sku: string
  systemSku: string
  owner: string
  group: string
  account: string
}

type KeepaRow = {
  asin: string
  title: string
  brand: string
  price: number | null
  newCurrent: number | null
  rank: number | null
  buyBox: string
  coupon: string
  primePrice: number | null
  image: string
}

type HistoryPoint = {
  date: string
  asin: string
  price: number | null
  rank: number | null
  buyBox: string
}

type UploadKind = 'keepa' | 'mapping' | 'monitor'
type EditMode = 'monitor' | 'mapping'
type MaintenancePanel = EditMode | null

type MetricChange = {
  direction: 'up' | 'down'
  tone: 'good' | 'bad'
  label: string
}

type RuleMatch = {
  owner: string
  group: string
  account: string
  source: string
}

type UploadSummary = {
  kind: UploadKind
  fileName: string
  imported: number
  notes: string[]
  errors: string[]
}

type AlertItem = {
  id: string
  message: string
  asin?: string
  sku?: string
  kmAsin?: string
  kmSku?: string
  kmPrice?: number | null
  kmRank?: number | null
  competitorPrice?: number | null
  previousRank?: number | null
  currentRank?: number | null
  changePct?: number | null
  owner?: string
  category?: string
  monitorIndexes?: number[]
}

type AlertGroup = {
  key: string
  title: string
  count: number
  items: AlertItem[]
}

type RankTrendRow = {
  asin: string
  sku: string
  owner: string
  asinType: string
  direction: 'up' | 'down'
  startRank: number
  endRank: number
  days: number
  changePct: number
}

type BuyBoxStatusItem = {
  id: string
  owner: string
  sku: string
  asin: string
  newCurrent: number | null
}

type PendingKeepaUpload = {
  file: File
  mode: 'initial' | 'from-yesterday' | 'daily'
}

type BuyBoxDayBoard = {
  date: string
  lost: BuyBoxStatusItem[]
  recovered: BuyBoxStatusItem[]
}

type StoredState = {
  keepaRows: KeepaRow[]
  mappingRows: MappingRow[]
  monitorRows: MonitorRow[]
  history: HistoryPoint[]
  todayBuyBox: BuyBoxDayBoard
  yesterdayBuyBox: BuyBoxDayBoard
}

const typedSeed = seedData as unknown as {
  keepaRows: KeepaRow[]
  mappingRows: MappingRow[]
  monitorRows: MonitorRow[]
}

const storageKey = 'buybox-monitor-state-v2'

const uploadLabels: Record<UploadKind, string> = {
  keepa: '每日 Keepa 数据源',
  mapping: '映射信息',
  monitor: 'SKU / ASIN 监控清单',
}

const monitorHeaders = ['运营', '组别', '账号', '平台SKU', 'Bundle主SKU', 'ASIN分类', 'ASIN', '备注']
const mappingHeaders = ['平台SKU', '系统SKU', '运营', '小组', '店铺别名']
const keepaHeaders = ['ASIN', 'Title', 'Brand', 'New: Current', 'Sales Rank: Current', 'Buy Box Seller', 'Coupon', 'Prime Price', 'Image']
const keepaFieldCandidates = {
  asin: ['ASIN', 'asin', 'Product Codes: ASIN', 'Product Codes ASIN'],
  title: ['Title', '标题', '商品标题', 'Parent Title'],
  brand: ['Brand', '品牌', 'Manufacturer'],
  price: ['Buy Box: Current', 'Amazon: Current', 'New: Current', 'Price', '今日价格', '标价'],
  newCurrent: ['New: Current'],
  rank: ['Sales Rank: Current', '大类排名', '今天大类排名'],
  buyBox: ['Buy Box: Buy Box Seller', 'Buy Box Seller', 'Buy Box: Seller', 'buybox', 'Buy Box'],
  coupon: ['Coupon', 'Coupon：金额', 'Coupon：百分比', 'One Time Coupon: Absolute', 'One Time Coupon: Percentage'],
  primePrice: ['Prime Price', 'Prime 价格', 'Prime价', 'New, Prime exclusive: Current'],
  image: ['Image', '图片'],
}

const normalized = (value: unknown) =>
  String(value ?? '')
    .trim()
    .toLowerCase()

const asNumber = (value: unknown) => {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const text = String(value ?? '').replace(/[$,%\s,]/g, '')
  if (!text || text === '-') return null
  const parsed = Number(text)
  return Number.isFinite(parsed) ? parsed : null
}

const pick = (row: AnyRow, candidates: string[]) => {
  const keys = Object.keys(row)
  for (const candidate of candidates) {
    const key = keys.find((item) => normalized(item) === normalized(candidate))
    if (key) return row[key]
  }
  for (const candidate of candidates) {
    const key = keys.find((item) => normalized(item).includes(normalized(candidate)))
    if (key) return row[key]
  }
  return ''
}

const getPrimaryImageUrl = (value: string | null | undefined) => {
  const raw = String(value ?? '').trim()
  if (!raw) return ''
  const first = raw
    .split(/[\n,;|]/)
    .map((item) => item.trim())
    .find(Boolean)

  if (!first) return ''
  if (/^https?:\/\//i.test(first)) return first
  if (/^\/\//.test(first)) return `https:${first}`
  if (/\.(jpg|jpeg|png|webp)/i.test(first)) return `https://m.media-amazon.com/images/I/${first.replace(/^\/+/, '')}`
  return ''
}

const readWorkbookRows = async (file: File): Promise<AnyRow[]> => {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer)
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
  return XLSX.utils.sheet_to_json<AnyRow>(firstSheet, { defval: '' })
}

const readKeepaRows = async (file: File): Promise<AnyRow[]> => {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { dense: true })
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
  const matrix = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(firstSheet, {
    header: 1,
    raw: false,
    defval: '',
  })
  const headers = (matrix[0] ?? []).map((cell) => String(cell ?? '').trim())
  const normalizedHeaders = headers.map((header) => normalized(header))
  const findIndex = (candidates: string[]) =>
    candidates.reduce<number>((found, candidate) => {
      if (found !== -1) return found
      const exact = normalizedHeaders.findIndex((header) => header === normalized(candidate))
      if (exact !== -1) return exact
      return normalizedHeaders.findIndex((header) => header.includes(normalized(candidate)))
    }, -1)

  const indices = {
    asin: findIndex(keepaFieldCandidates.asin),
    title: findIndex(keepaFieldCandidates.title),
    brand: findIndex(keepaFieldCandidates.brand),
    price: findIndex(keepaFieldCandidates.price),
    newCurrent: normalizedHeaders.findIndex((header) => header === normalized('New: Current')),
    rank: findIndex(keepaFieldCandidates.rank),
    buyBox: findIndex(keepaFieldCandidates.buyBox),
    coupon: findIndex(keepaFieldCandidates.coupon),
    primePrice: findIndex(keepaFieldCandidates.primePrice),
    image: findIndex(keepaFieldCandidates.image),
  }

  if (indices.asin === -1) {
    throw new Error('Keepa 文件中未找到 ASIN 列，请确认导出包含 ASIN 字段。')
  }
  if (indices.newCurrent === -1) {
    throw new Error('Keepa 文件中未找到 New: Current 列，请使用每日 Product Viewer 导出模板。')
  }

  return matrix.slice(1).map((row) => ({
    ASIN: indices.asin >= 0 ? row[indices.asin] : '',
    Title: indices.title >= 0 ? row[indices.title] : '',
    Brand: indices.brand >= 0 ? row[indices.brand] : '',
    Price: indices.price >= 0 ? row[indices.price] : '',
    'New: Current': row[indices.newCurrent],
    'Sales Rank: Current': indices.rank >= 0 ? row[indices.rank] : '',
    'Buy Box Seller': indices.buyBox >= 0 ? row[indices.buyBox] : '',
    Coupon: indices.coupon >= 0 ? row[indices.coupon] : '',
    'Prime Price': indices.primePrice >= 0 ? row[indices.primePrice] : '',
    Image: indices.image >= 0 ? row[indices.image] : '',
  }))
}

const parseKeepa = (rows: AnyRow[]): KeepaRow[] =>
  rows
    .map((row) => {
      const asin = pick(row, keepaFieldCandidates.asin)
      return {
        asin: String(asin || '').trim(),
        title: String(pick(row, keepaFieldCandidates.title) || '').trim(),
        brand: String(pick(row, keepaFieldCandidates.brand) || '').trim(),
        price: asNumber(pick(row, keepaFieldCandidates.price)),
        newCurrent: asNumber(pick(row, keepaFieldCandidates.newCurrent)),
        rank: asNumber(pick(row, keepaFieldCandidates.rank)),
        buyBox: String(pick(row, keepaFieldCandidates.buyBox) || '').trim(),
        coupon: String(pick(row, keepaFieldCandidates.coupon) || '').trim(),
        primePrice: asNumber(pick(row, keepaFieldCandidates.primePrice)),
        image: String(pick(row, keepaFieldCandidates.image) || '').trim(),
      }
    })
    .filter((row) => row.asin)

const parseMapping = (rows: AnyRow[]): MappingRow[] =>
  rows
    .map((row) => ({
      sku: String(pick(row, ['平台SKU', 'SKU', 'sku']) || '').trim(),
      systemSku: String(pick(row, ['系统SKU', 'System SKU']) || '').trim(),
      owner: String(pick(row, ['运营', '负责人', 'Owner']) || '').trim(),
      group: String(pick(row, ['小组', '组别', 'Group']) || '').trim(),
      account: String(pick(row, ['店铺别名', '账号', 'Account']) || '').trim(),
    }))
    .filter((row) => row.sku)

const parseMonitor = (rows: AnyRow[]): MonitorRow[] =>
  rows
    .map((row) => ({
      owner: String(pick(row, ['运营', 'Owner']) || '').trim(),
      group: String(pick(row, ['组别', '小组', 'Group']) || '').trim(),
      account: String(pick(row, ['账号', '店铺别名', 'Account']) || '').trim(),
      sku: String(pick(row, ['平台sku', '平台SKU', 'SKU']) || '').trim(),
      bundleSku: String(pick(row, ['Bundle主SKU', 'Bundle 主 SKU']) || '').trim(),
      asinType: String(pick(row, ['ASIN分类', 'ASIN 分类', '类型']) || '').trim(),
      asin: String(pick(row, ['ASIN', 'asin']) || '').trim(),
      note: String(pick(row, ['竞对备注列', '备注', '库存状态备注']) || '').trim(),
    }))
    .filter((row) => row.sku && row.asin)

const initialHistory = (keepaRows: KeepaRow[]): HistoryPoint[] => {
  const date = new Date().toLocaleDateString('zh-CN', { month: '2-digit', day: '2-digit' })
  return keepaRows.map((row) => ({
    date,
    asin: row.asin,
    price: row.price,
    rank: row.rank,
    buyBox: row.buyBox,
  }))
}

const keepRecentFiveDays = (points: HistoryPoint[]) => {
  const recentDates = [...new Set(points.map((point) => point.date))].slice(-5)
  return points.filter((point) => recentDates.includes(point.date))
}

const mergeRecentHistory = (current: HistoryPoint[], nextRows: KeepaRow[], date: string) => {
  const nextPoints = nextRows.map((row) => ({
    date,
    asin: row.asin,
    price: row.price,
    rank: row.rank,
    buyBox: row.buyBox,
  }))
  return keepRecentFiveDays([...current.filter((point) => point.date !== date), ...nextPoints])
}

const emptyBuyBoxBoard = (): BuyBoxDayBoard => ({ date: '', lost: [], recovered: [] })

const normalizeKeepaRows = (rows: KeepaRow[] | undefined) =>
  (rows ?? []).map((row) => ({
    ...row,
    newCurrent: typeof row.newCurrent === 'number' && Number.isFinite(row.newCurrent) ? row.newCurrent : null,
  }))

const normalizeBuyBoxBoard = (board: BuyBoxDayBoard | undefined): BuyBoxDayBoard => ({
  date: String(board?.date ?? ''),
  lost: Array.isArray(board?.lost) ? board.lost : [],
  recovered: Array.isArray(board?.recovered) ? board.recovered : [],
})

const getImportDate = (fileName: string) => {
  const fileDate = fileName.match(/\d{4}-\d{2}-\d{2}/)?.[0]
  return fileDate ?? new Date().toLocaleDateString('sv-SE')
}

const buildBuyBoxBoard = (
  monitorRows: MonitorRow[],
  mappingRows: MappingRow[],
  currentRows: KeepaRow[],
  previousRows: KeepaRow[],
  date: string,
): BuyBoxDayBoard => {
  const currentByAsin = new Map(currentRows.map((row) => [normalized(row.asin), row.newCurrent]))
  const previousByAsin = new Map(previousRows.map((row) => [normalized(row.asin), row.newCurrent]))
  const mappingBySku = new Map(mappingRows.map((row) => [normalized(row.sku), row]))
  const lost: BuyBoxStatusItem[] = []
  const recovered: BuyBoxStatusItem[] = []
  const seen = new Set<string>()

  for (const row of monitorRows) {
    if (!normalized(row.asinType).includes('kmasin')) continue
    const asinKey = normalized(row.asin)
    const itemKey = `${normalized(row.sku)}::${asinKey}`
    if (seen.has(itemKey)) continue
    seen.add(itemKey)

    if (!currentByAsin.has(asinKey)) continue
    const newCurrent = currentByAsin.get(asinKey)
    const item: BuyBoxStatusItem = {
      id: `buybox-${itemKey}`,
      owner: row.owner || mappingBySku.get(normalized(row.sku))?.owner || '',
      sku: row.sku,
      asin: row.asin,
      newCurrent: newCurrent ?? null,
    }

    const isCurrentLost = newCurrent === null || newCurrent === 0
    const previousExists = previousByAsin.has(asinKey)
    const previousNewCurrent = previousByAsin.get(asinKey)
    const wasPreviouslyLost = previousExists && (previousNewCurrent === null || previousNewCurrent === 0)
    if (isCurrentLost) lost.push(item)
    if (wasPreviouslyLost && typeof newCurrent === 'number' && newCurrent > 0) recovered.push(item)
  }

  return { date, lost, recovered }
}

const getMetricChange = (
  points: HistoryPoint[] | undefined,
  metric: 'price' | 'rank',
): MetricChange | null => {
  const validPoints = (points ?? []).filter((point) => typeof point[metric] === 'number')
  if (validPoints.length < 2) return null
  const previous = validPoints[validPoints.length - 2][metric]
  const current = validPoints[validPoints.length - 1][metric]
  if (typeof previous !== 'number' || typeof current !== 'number' || previous === current) return null
  const delta = current - previous
  if (metric === 'price') {
    return {
      direction: delta > 0 ? 'up' : 'down',
      tone: delta > 0 ? 'bad' : 'good',
      label: `${delta > 0 ? '+' : ''}${delta.toFixed(2)}`,
    }
  }
  return {
    direction: delta < 0 ? 'up' : 'down',
    tone: delta < 0 ? 'bad' : 'good',
    label: Math.abs(delta).toLocaleString(),
  }
}

const getRuleMatch = (row: MonitorRow, mapping?: MappingRow, keepa?: KeepaRow): RuleMatch => {
  if (row.owner || row.group || row.account) {
    return {
      owner: row.owner,
      group: row.group,
      account: row.account,
      source: '监控清单精确值',
    }
  }
  if (mapping) {
    return {
      owner: mapping.owner,
      group: mapping.group,
      account: mapping.account,
      source: '平台SKU映射',
    }
  }
  const title = normalized(keepa?.title)
  const brand = normalized(keepa?.brand)
  const asinType = normalized(row.asinType)
  if (title || brand) {
    if (asinType.includes('km')) {
      return {
        owner: '',
        group: '',
        account: '',
        source: 'Keepa标题/品牌识别(KM)',
      }
    }
    return {
      owner: '',
      group: '',
      account: '',
      source: 'Keepa标题/品牌识别',
    }
  }
  return {
    owner: '',
    group: '',
    account: '',
    source: '待补充规则',
  }
}

const loadInitialState = (): StoredState => {
  const seedKeepaRows = normalizeKeepaRows(typedSeed.keepaRows)
  const fallback = {
    keepaRows: seedKeepaRows,
    mappingRows: typedSeed.mappingRows,
    monitorRows: typedSeed.monitorRows,
    history: initialHistory(seedKeepaRows),
    todayBuyBox: emptyBuyBoxBoard(),
    yesterdayBuyBox: emptyBuyBoxBoard(),
  }
  try {
    const raw = localStorage.getItem(storageKey)
    if (!raw) return fallback
    const stored = JSON.parse(raw) as StoredState
    return {
      keepaRows: normalizeKeepaRows(stored.keepaRows),
      mappingRows: stored.mappingRows ?? fallback.mappingRows,
      monitorRows: stored.monitorRows ?? fallback.monitorRows,
      history: keepRecentFiveDays(stored.history ?? []),
      todayBuyBox: normalizeBuyBoxBoard(stored.todayBuyBox),
      yesterdayBuyBox: normalizeBuyBoxBoard(stored.yesterdayBuyBox),
    }
  } catch {
    return fallback
  }
}

const emptyMonitor: MonitorRow = {
  owner: '',
  group: '',
  account: '',
  sku: '',
  bundleSku: '',
  asinType: '竞对ASIN',
  asin: '',
  note: '',
}

const emptyMapping: MappingRow = {
  sku: '',
  systemSku: '',
  owner: '',
  group: '',
  account: '',
}

const emptyUploadSummary: UploadSummary = {
  kind: 'keepa',
  fileName: '',
  imported: 0,
  notes: [],
  errors: [],
}

const downloadTemplate = (kind: UploadKind) => {
  const headers = kind === 'monitor' ? monitorHeaders : kind === 'mapping' ? mappingHeaders : keepaHeaders
  const example =
    kind === 'monitor'
      ? ['SYZ1525【Zuri】', 'Polaris 1组', '老三-US', 'KM1B0901-06205-BK', '', '竞对ASIN', 'B08CF2V57W', '重点竞对']
      : kind === 'mapping'
        ? ['KM1B0901-06205-BK', 'B0901-06205-BK', 'SYZ1525【Zuri】', 'Polaris 1组', '老三-US']
        : ['B0CZSN3KZV', 'Product title', 'KEMIMOTO', 52.99, 265881, 'KEMIMOTO', '$5', 43.69, 'https://...']
  const sheet = XLSX.utils.aoa_to_sheet([headers, example])
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, '模板')
  XLSX.writeFile(workbook, `${uploadLabels[kind]}模板.xlsx`)
}

const exportRows = (mode: EditMode, rows: MonitorRow[] | MappingRow[]) => {
  const data =
    mode === 'monitor'
      ? (rows as MonitorRow[]).map((row) => ({
          运营: row.owner,
          组别: row.group,
          账号: row.account,
          平台SKU: row.sku,
          Bundle主SKU: row.bundleSku,
          ASIN分类: row.asinType,
          ASIN: row.asin,
          备注: row.note,
        }))
      : (rows as MappingRow[]).map((row) => ({
          平台SKU: row.sku,
          系统SKU: row.systemSku,
          运营: row.owner,
          小组: row.group,
          店铺别名: row.account,
        }))
  const sheet = XLSX.utils.json_to_sheet(data)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, mode === 'monitor' ? 'SKU监控清单' : '映射信息')
  XLSX.writeFile(workbook, `${mode === 'monitor' ? 'SKU监控清单' : '映射信息'}导出.xlsx`)
}

const exportAlertItems = (title: string, items: AlertItem[]) => {
  const data = items.map((item) => ({
    说明: item.message,
    运营: item.owner || '',
    SKU: item.sku || '',
    ASIN: item.asin || '',
    类型: item.category || '',
  }))
  const sheet = XLSX.utils.json_to_sheet(data)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, '预警清单')
  XLSX.writeFile(workbook, `${title}预警清单.xlsx`)
}

const exportBoardItems = (title: string, items: AlertItem[]) => {
  const data = items.map((item) => ({
    运营: item.owner || '',
    KMSKU: item.kmSku || item.sku || '',
    KMASIN: item.kmAsin || '',
    ASIN类型: item.category || '',
    竞对ASIN: item.asin || '',
    KM价格: item.kmPrice ?? '',
    当前价格: item.competitorPrice ?? '',
    价格变化: item.changePct ?? '',
    前次排名: item.previousRank ?? '',
    当前排名: item.currentRank ?? '',
  }))
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data), '看板导出')
  XLSX.writeFile(workbook, `${title}看板.xlsx`)
}

const exportMonitorWithDuplicateMarks = (rows: MonitorRow[]) => {
  const grouped = new Map<string, number[]>()
  rows.forEach((row, index) => {
    const key = `${normalized(row.owner)}::${normalized(row.sku)}::${normalized(row.asinType)}::${normalized(row.asin)}`
    grouped.set(key, [...(grouped.get(key) ?? []), index])
  })
  const duplicateIndexes = new Set<number>()
  for (const indexes of grouped.values()) {
    if (indexes.length > 1) indexes.forEach((index) => duplicateIndexes.add(index))
  }
  const markedData = rows.map((row, index) => ({
    重复标记: duplicateIndexes.has(index) ? '重复' : '',
    运营: row.owner,
    组别: row.group,
    账号: row.account,
    平台SKU: row.sku,
    Bundle主SKU: row.bundleSku,
    ASIN分类: row.asinType,
    ASIN: row.asin,
    备注: row.note,
  }))
  const dedupedRows = rows.filter((_, index) => !duplicateIndexes.has(index))
  const dedupedData = dedupedRows.map((row) => ({
    运营: row.owner,
    组别: row.group,
    账号: row.account,
    平台SKU: row.sku,
    Bundle主SKU: row.bundleSku,
    ASIN分类: row.asinType,
    ASIN: row.asin,
    备注: row.note,
  }))
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(markedData), '完整表_标重复')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dedupedData), '剔重后可重传')
  XLSX.writeFile(workbook, '重复监控关系检查表.xlsx')
}

const BuyBoxStatusSection = ({
  title,
  items,
  tone,
}: {
  title: string
  items: BuyBoxStatusItem[]
  tone: 'lost' | 'recovered'
}) => (
  <section className={`buybox-status-section buybox-status-${tone}`}>
    <div className="buybox-status-heading">
      <h3>{title}</h3>
      <span>{items.length} 条</span>
    </div>
    <div className="mini-table-wrap buybox-table-wrap">
      <table className="data-table mini-table">
        <thead><tr><th>运营</th><th>SKU</th><th>ASIN</th><th>New: Current</th></tr></thead>
        <tbody>
          {items.length
            ? items.map((item) => <tr key={item.id}><td>{item.owner || '-'}</td><td>{item.sku}</td><td>{item.asin}</td><td>{item.newCurrent === null ? '空值' : item.newCurrent === 0 ? '0' : item.newCurrent.toFixed(2)}</td></tr>)
            : <tr className="buybox-empty-row"><td colSpan={4}>暂无数据</td></tr>}
        </tbody>
      </table>
    </div>
  </section>
)

function App() {
  const initialState = useMemo(loadInitialState, [])
  const [keepaRows, setKeepaRows] = useState<KeepaRow[]>(initialState.keepaRows)
  const [mappingRows, setMappingRows] = useState<MappingRow[]>(initialState.mappingRows)
  const [monitorRows, setMonitorRows] = useState<MonitorRow[]>(initialState.monitorRows)
  const [history, setHistory] = useState<HistoryPoint[]>(initialState.history)
  const [todayBuyBox, setTodayBuyBox] = useState<BuyBoxDayBoard>(initialState.todayBuyBox)
  const [yesterdayBuyBox, setYesterdayBuyBox] = useState<BuyBoxDayBoard>(initialState.yesterdayBuyBox)
  const [pendingKeepaUpload, setPendingKeepaUpload] = useState<PendingKeepaUpload | null>(null)
  const [ownerQuery, setOwnerQuery] = useState('')
  const [skuQuery, setSkuQuery] = useState('')
  const [asinQuery, setAsinQuery] = useState('')
  const [keywordQuery, setKeywordQuery] = useState('')
  const [selectedAsin, setSelectedAsin] = useState(initialState.monitorRows[0]?.asin ?? '')
  const [priceAlert, setPriceAlert] = useState(12)
  const [rankAlert, setRankAlert] = useState(35)
  const [editMode, setEditMode] = useState<EditMode>('monitor')
  const [maintenancePanel, setMaintenancePanel] = useState<MaintenancePanel>(null)
  const [editingIndex, setEditingIndex] = useState<number | null>(null)
  const [monitorForm, setMonitorForm] = useState<MonitorRow>(emptyMonitor)
  const [mappingForm, setMappingForm] = useState<MappingRow>(emptyMapping)
  const [uploadSummary, setUploadSummary] = useState<UploadSummary>(emptyUploadSummary)
  const [collapsedAlerts, setCollapsedAlerts] = useState<Record<string, boolean>>({})
  const [priceViewOwner, setPriceViewOwner] = useState('')
  const [priceViewSku, setPriceViewSku] = useState('')
  const [priceViewCategory, setPriceViewCategory] = useState('')
  const [priceSortBy, setPriceSortBy] = useState<'price-diff-desc' | 'price-diff-asc' | 'rank-diff-desc' | 'rank-diff-asc'>('price-diff-desc')
  const [priceSortMenu, setPriceSortMenu] = useState<'price' | 'rank' | null>(null)
  const editorPanelRef = useRef<HTMLElement | null>(null)
  const [status, setStatus] = useState(
    `已自动载入原 Excel：${initialState.monitorRows.length} 条监控清单、${initialState.mappingRows.length} 条映射、${initialState.keepaRows.length} 条 Keepa。趋势仅保留最近 5 天。`,
  )

  useEffect(() => {
    const payload: StoredState = { keepaRows, mappingRows, monitorRows, history, todayBuyBox, yesterdayBuyBox }
    localStorage.setItem(storageKey, JSON.stringify(payload))
  }, [history, keepaRows, mappingRows, monitorRows, todayBuyBox, yesterdayBuyBox])

  useEffect(() => {
    if (!pendingKeepaUpload) return undefined
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') cancelKeepaUpload()
    }
    document.addEventListener('keydown', closeOnEscape)
    return () => document.removeEventListener('keydown', closeOnEscape)
  }, [pendingKeepaUpload])

  const keepaByAsin = useMemo(() => new Map(keepaRows.map((row) => [normalized(row.asin), row])), [keepaRows])
  const mappingBySku = useMemo(() => new Map(mappingRows.map((row) => [normalized(row.sku), row])), [mappingRows])

  const enrichedRows = useMemo(
    () =>
      monitorRows.map((row) => {
        const keepa = keepaByAsin.get(normalized(row.asin))
        const mapping = mappingBySku.get(normalized(row.sku))
        const ruleMatch = getRuleMatch(row, mapping, keepa)
        return {
          ...row,
          owner: ruleMatch.owner,
          group: ruleMatch.group,
          account: ruleMatch.account,
          ruleSource: ruleMatch.source,
          keepa,
          hasMapping: Boolean(mapping),
        }
      }),
    [keepaByAsin, mappingBySku, monitorRows],
  )

  const filteredRows = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const skuNeedle = normalized(skuQuery)
    const asinNeedle = normalized(asinQuery)
    const keywordNeedle = normalized(keywordQuery)
    return enrichedRows.filter((row) =>
      (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
      (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
      (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
      (!keywordNeedle ||
        [row.group, row.account, row.asinType, row.note, row.keepa?.brand, row.keepa?.title]
          .map(normalized)
          .some((value) => value.includes(keywordNeedle))),
    )
  }, [asinQuery, enrichedRows, keywordQuery, ownerQuery, skuQuery])

  const skuOptions = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const asinNeedle = normalized(asinQuery)
    const keywordNeedle = normalized(keywordQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!keywordNeedle ||
          [row.group, row.account, row.asinType, row.note, row.keepa?.brand, row.keepa?.title]
            .map(normalized)
            .some((value) => value.includes(keywordNeedle)))
      if (matchesContext && row.sku) options.add(row.sku)
      if (matchesContext && row.bundleSku) options.add(row.bundleSku)
      if (options.size >= 120) break
    }
    return [...options].sort()
  }, [asinQuery, enrichedRows, keywordQuery, ownerQuery])

  const ownerOptions = useMemo(() => {
    const skuNeedle = normalized(skuQuery)
    const asinNeedle = normalized(asinQuery)
    const keywordNeedle = normalized(keywordQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!keywordNeedle ||
          [row.group, row.account, row.asinType, row.note, row.keepa?.brand, row.keepa?.title]
            .map(normalized)
            .some((value) => value.includes(keywordNeedle)))
      if (matchesContext && row.owner) options.add(row.owner)
      if (options.size >= 120) break
    }
    return [...options].sort()
  }, [asinQuery, enrichedRows, keywordQuery, skuQuery])

  const asinOptions = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const skuNeedle = normalized(skuQuery)
    const keywordNeedle = normalized(keywordQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!keywordNeedle ||
          [row.group, row.account, row.asinType, row.note, row.keepa?.brand, row.keepa?.title]
            .map(normalized)
            .some((value) => value.includes(keywordNeedle)))
      if (matchesContext && row.asin) options.add(row.asin)
      if (options.size >= 160) break
    }
    return [...options].sort()
  }, [enrichedRows, keywordQuery, ownerQuery, skuQuery])

  const selectedRows = useMemo(
    () => enrichedRows.filter((row) => normalized(row.asin) === normalized(selectedAsin)),
    [enrichedRows, selectedAsin],
  )

  const kmAsinBySku = useMemo(() => {
    const grouped = new Map<string, typeof enrichedRows>()
    for (const row of enrichedRows) {
      const key = normalized(row.sku)
      grouped.set(key, [...(grouped.get(key) ?? []), row])
    }
    return grouped
  }, [enrichedRows])

  const selectedHistory = useMemo(
    () => history.filter((point) => normalized(point.asin) === normalized(selectedAsin)),
    [history, selectedAsin],
  )

  const historyByAsin = useMemo(() => {
    const grouped = new Map<string, HistoryPoint[]>()
    for (const point of history) {
      const key = normalized(point.asin)
      grouped.set(key, [...(grouped.get(key) ?? []), point])
    }
    return grouped
  }, [history])

  const editableRows = editMode === 'monitor' ? monitorRows : mappingRows
  const hasCurrentBuyBoxSnapshot = useMemo(
    () => Boolean(todayBuyBox.date || yesterdayBuyBox.date),
    [todayBuyBox.date, yesterdayBuyBox.date],
  )

  const alertGroups = useMemo(() => {
    const missingMapping: AlertItem[] = []
    const missingKeepa: AlertItem[] = []
    const unresolvedRules: AlertItem[] = []
    const duplicateRelations: AlertItem[] = []
    const priceChanges: AlertItem[] = []
    const rankDrops: AlertItem[] = []
    const grouped = new Map<string, HistoryPoint[]>()
    const duplicateKeys = new Set<string>()
    const seenMonitorKeys = new Set<string>()
    const duplicateIndexMap = new Map<string, number[]>()

    for (const [index, row] of enrichedRows.entries()) {
      const relationKey = `${normalized(row.owner)}::${normalized(row.sku)}::${normalized(row.asinType)}::${normalized(row.asin)}`
      if (seenMonitorKeys.has(relationKey)) duplicateKeys.add(relationKey)
      seenMonitorKeys.add(relationKey)
      duplicateIndexMap.set(relationKey, [...(duplicateIndexMap.get(relationKey) ?? []), index])

      const needsMappingSupport = !row.owner || !row.group || !row.account
      if (!row.hasMapping && needsMappingSupport) {
        missingMapping.push({
          id: `missing-mapping-${index}`,
          message: `${row.sku} 未匹配映射信息，需要补充运营/组别/账号映射`,
          asin: row.asin,
          sku: row.sku,
          owner: row.owner,
          category: row.asinType,
          monitorIndexes: [index],
        })
      }
      if (!row.keepa) {
        missingKeepa.push({
          id: `missing-keepa-${index}`,
          message: `${row.asin} 缺少 Keepa 数据，无法更新价格与排名`,
          asin: row.asin,
          sku: row.sku,
          owner: row.owner,
          category: row.asinType,
          monitorIndexes: [index],
        })
      }
      if (row.ruleSource === '待补充规则') {
        unresolvedRules.push({
          id: `rule-gap-${index}`,
          message: `${row.asin} 缺少运营归类规则，请补充监控清单或映射信息`,
          asin: row.asin,
          sku: row.sku,
          owner: row.owner,
          category: row.asinType,
          monitorIndexes: [index],
        })
      }
    }

    for (const key of duplicateKeys) {
      const [, sku, asinType, asin] = key.split('::')
      const duplicateIndexes = duplicateIndexMap.get(key) ?? []
      const rows = duplicateIndexes.map((index) => enrichedRows[index]).filter(Boolean)
      const owners = [...new Set(rows.map((row) => row.owner || '未填运营'))].join('、')
      duplicateRelations.push({
        id: `duplicate-${key}`,
        message: `${owners} 下 ${sku.toUpperCase()} / ${asinType} / ${asin.toUpperCase()} 重复监控`,
        asin: asin.toUpperCase(),
        sku: sku.toUpperCase(),
        owner: owners,
        category: asinType,
        monitorIndexes: duplicateIndexes,
      })
    }

    for (const point of history) grouped.set(point.asin, [...(grouped.get(point.asin) ?? []), point])
    for (const [asin, points] of grouped) {
      const ordered = [...points].slice(-2)
      if (ordered.length < 2) continue
      const [previous, current] = ordered
      const relatedRows = enrichedRows.filter((row) => normalized(row.asin) === normalized(asin))
      const row = relatedRows[0]
      const sameSkuRows = row?.sku ? kmAsinBySku.get(normalized(row.sku)) ?? [] : []
      const kmRow = sameSkuRows.find((candidate) => normalized(candidate.asinType).includes('kmasin'))
      const kmHistoryPoints = kmRow ? historyByAsin.get(normalized(kmRow.asin)) : undefined
      const kmPriceCurrent = kmRow?.keepa?.price ?? null
      const kmPricePrevious = (kmHistoryPoints ?? []).filter((point) => typeof point.price === 'number').slice(-2)[0]?.price ?? null
      if (previous.price && current.price) {
        const change = Math.abs((current.price - previous.price) / previous.price) * 100
        if (change >= priceAlert) {
          priceChanges.push({
            id: `price-${asin}`,
            message: `${asin} 价格变动 ${change.toFixed(1)}%，超过 ${priceAlert}%`,
            asin,
            sku: row?.sku,
            kmAsin: kmRow?.asin,
            kmSku: kmRow?.sku ?? row?.sku,
            kmPrice: kmPriceCurrent ?? kmPricePrevious,
            competitorPrice: current.price,
            owner: row?.owner,
            category: row?.asinType,
            previousRank: previous.rank,
            currentRank: current.rank,
            changePct: change,
          })
        }
      }
      if (previous.rank && current.rank) {
        const change = ((current.rank - previous.rank) / previous.rank) * 100
        if (change >= rankAlert) {
          rankDrops.push({
            id: `rank-${asin}`,
            message: `${asin} 排名下滑 ${change.toFixed(1)}%，超过 ${rankAlert}%`,
            asin,
            sku: row?.sku,
            owner: row?.owner,
            category: row?.asinType,
          })
        }
      }
    }

    const groups: AlertGroup[] = [
      { key: 'missing-keepa', title: '缺少 Keepa', count: missingKeepa.length, items: missingKeepa },
      { key: 'missing-mapping', title: '缺少映射', count: missingMapping.length, items: missingMapping },
      { key: 'rule-gap', title: '缺少归类规则', count: unresolvedRules.length, items: unresolvedRules },
      { key: 'duplicate', title: '重复监控关系', count: duplicateRelations.length, items: duplicateRelations },
    ]
    return groups.filter((group) => group.count > 0)
  }, [enrichedRows, history, historyByAsin, kmAsinBySku, priceAlert, rankAlert])

  const alerts = useMemo(
    () => alertGroups.flatMap((group) => group.items.map((item) => item.message)),
    [alertGroups],
  )

  const stats = [
    { label: '监控 ASIN', value: new Set(monitorRows.map((row) => row.asin)).size },
    { label: '平台 SKU', value: new Set(monitorRows.map((row) => row.sku)).size },
    { label: '映射 SKU', value: mappingRows.length },
    { label: 'Keepa 记录', value: keepaRows.length },
    { label: '当前预警', value: alerts.length },
  ]

  const missingKeepaItems = useMemo(
    () => alertGroups.find((group) => group.key === 'missing-keepa')?.items ?? [],
    [alertGroups],
  )

  const duplicateItems = useMemo(
    () => alertGroups.find((group) => group.key === 'duplicate')?.items ?? [],
    [alertGroups],
  )

  const ruleStats = useMemo(() => {
    const direct = enrichedRows.filter((row) => row.ruleSource === '监控清单精确值').length
    const mapped = enrichedRows.filter((row) => row.ruleSource === '平台SKU映射').length
    const keepaMatched = enrichedRows.filter((row) => row.ruleSource.includes('Keepa标题/品牌识别')).length
    const unresolved = enrichedRows.filter((row) => row.ruleSource === '待补充规则').length
    return { direct, mapped, keepaMatched, unresolved }
  }, [enrichedRows])

  const boardRows = useMemo(() => {
    const rows: AlertItem[] = []
    const ownerNeedle = normalized(priceViewOwner)
    const skuNeedle = normalized(priceViewSku)
    const categoryNeedle = normalized(priceViewCategory)

    for (const [asin, points] of historyByAsin) {
      const ordered = [...points].slice(-2)
      if (ordered.length < 2) continue
      const [previous, current] = ordered
      if (typeof previous.price !== 'number' || typeof current.price !== 'number') continue

      const relatedRows = enrichedRows.filter((row) => normalized(row.asin) === asin)
      const row = relatedRows[0]
      if (!row) continue

      const sameSkuRows = row.sku ? kmAsinBySku.get(normalized(row.sku)) ?? [] : []
      const kmRow = sameSkuRows.find((candidate) => normalized(candidate.asinType).includes('kmasin'))
      if (!kmRow) continue

      const pricePct = previous.price !== 0 ? Math.abs((current.price - previous.price) / previous.price) * 100 : 0
      const matchesFilter =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!skuNeedle || normalized(row.sku).includes(skuNeedle)) &&
        (!categoryNeedle || normalized(row.asinType).includes(categoryNeedle))

      if (!matchesFilter) continue

      rows.push({
        id: `board-${asin}`,
        message: `${asin} 价格 / 排名变化`,
        asin: row.asin,
        sku: row.sku,
        kmAsin: kmRow.asin,
        kmSku: kmRow.sku,
        kmPrice: kmRow.keepa?.price ?? null,
        kmRank: kmRow.keepa?.rank ?? null,
        competitorPrice: current.price,
        owner: row.owner,
        category: row.asinType,
        previousRank: previous.rank,
        currentRank: current.rank,
        changePct: pricePct,
      })
    }

    return rows
  }, [enrichedRows, historyByAsin, kmAsinBySku, priceViewCategory, priceViewOwner, priceViewSku])

  const priceAlertRows = useMemo(() => {
    return boardRows
      .filter((item) => typeof item.changePct === 'number' && item.changePct >= priceAlert)
      .map((item) => {
        const historyPoints = historyByAsin.get(normalized(item.asin))
        const validPoints = (historyPoints ?? []).filter((point) => typeof point.price === 'number')
        const previous = validPoints[validPoints.length - 2]?.price ?? null
        const current = validPoints[validPoints.length - 1]?.price ?? null
        const delta = typeof previous === 'number' && typeof current === 'number' ? current - previous : null
        const rankDelta =
          typeof item.previousRank === 'number' && typeof item.currentRank === 'number'
            ? item.currentRank - item.previousRank
            : null
        const changePct =
          typeof previous === 'number' && typeof current === 'number' && previous !== 0
            ? Math.abs((current - previous) / previous) * 100
            : null
        return {
          ...item,
          previous,
          current,
          delta,
          rankDelta,
          changePct,
        }
      })
      .sort((left, right) => {
        const leftPriceDiff = Math.abs(left.delta ?? 0)
        const rightPriceDiff = Math.abs(right.delta ?? 0)
        const leftRankDiff = Math.abs(left.rankDelta ?? 0)
        const rightRankDiff = Math.abs(right.rankDelta ?? 0)
        if (priceSortBy === 'price-diff-desc') return rightPriceDiff - leftPriceDiff
        if (priceSortBy === 'price-diff-asc') return leftPriceDiff - rightPriceDiff
        if (priceSortBy === 'rank-diff-desc') return rightRankDiff - leftRankDiff
        return leftRankDiff - rightRankDiff
      })
  }, [boardRows, historyByAsin, priceAlert, priceSortBy])

  const priceBoardGroups = useMemo(() => {
    const grouped = new Map<string, typeof priceAlertRows>()
    for (const row of priceAlertRows) {
      const key = `${row.owner || ''}::${row.kmSku || row.sku || ''}::${row.kmAsin || ''}`
      grouped.set(key, [...(grouped.get(key) ?? []), row])
    }
    return [...grouped.values()]
      .map((rows) => rows.sort((a, b) => (b.changePct ?? 0) - (a.changePct ?? 0)))
      .sort((a, b) => (b[0]?.changePct ?? 0) - (a[0]?.changePct ?? 0))
  }, [priceAlertRows])

  const rankTrendRows = useMemo(() => {
    const rows: RankTrendRow[] = []
    for (const [asin, points] of historyByAsin) {
      const valid = points.filter((point) => typeof point.rank === 'number')
      if (valid.length < 3) continue
      const ranks = valid.map((point) => point.rank as number)
      let downStreak = 1
      let upStreak = 1
      for (let index = 1; index < ranks.length; index += 1) {
        if (ranks[index] > ranks[index - 1]) downStreak += 1
        else downStreak = 1
        if (ranks[index] < ranks[index - 1]) upStreak += 1
        else upStreak = 1
      }
      const isDown = downStreak >= 3
      const isUp = upStreak >= 3
      if (!isDown && !isUp) continue
      const related = enrichedRows.find((row) => normalized(row.asin) === asin)
      if (!related) continue
      const startRank = ranks[0]
      const endRank = ranks[ranks.length - 1]
      const changePct = startRank !== 0 ? Math.abs((endRank - startRank) / startRank) * 100 : 0
      rows.push({
        asin: related.asin,
        sku: related.sku,
        owner: related.owner,
        asinType: related.asinType,
        direction: isDown ? 'down' : 'up',
        startRank,
        endRank,
        days: valid.length,
        changePct,
      })
    }
    return rows.sort((a, b) => b.changePct - a.changePct)
  }, [enrichedRows, historyByAsin])

  const visibleRows = filteredRows.slice(0, 180)

  const renderMetric = (value: number | null | undefined, change: MetricChange | null, format: 'price' | 'rank') => {
    if (typeof value !== 'number') return '-'
    const display = format === 'price' ? `$${value.toFixed(2)}` : value.toLocaleString()
    return (
      <span className="metric-cell">
        <span className="metric-value">{display}</span>
        {change ? (
          <span className={`metric-change metric-${change.tone}`} title={`较上次变化 ${change.label}`}>
            {change.direction === 'up' ? <ArrowUp size={13} /> : <ArrowDown size={13} />}
            <span>{change.label}</span>
          </span>
        ) : null}
      </span>
    )
  }

  const renderBoardPriceMetric = (
    current: number | null | undefined,
    delta: number | null | undefined,
    changePct: number | null | undefined,
  ) => {
    if (typeof current !== 'number') return '-'
    const toneClass = typeof delta === 'number' ? (delta > 0 ? 'delta-bad' : 'delta-good') : ''
    return (
      <span className="inline-metric">
        <span>{`$${current.toFixed(2)}`}</span>
        {typeof delta === 'number' && delta !== 0 ? (
          <span className={`${toneClass} plain-delta`}>
            {delta > 0 ? <ArrowUp size={13} /> : <ArrowDown size={13} />}
            <span>{`${delta > 0 ? '+' : ''}${delta.toFixed(2)}`}</span>
            {typeof changePct === 'number' ? <span>{`(${changePct.toFixed(1)}%)`}</span> : null}
          </span>
        ) : null}
      </span>
    )
  }

  const renderBoardRankMetric = (
    current: number | null | undefined,
    delta: number | null | undefined,
    startRank?: number | null,
  ) => {
    if (typeof current !== 'number') return '-'
    const isImproved = typeof delta === 'number' && delta < 0
    const toneClass = typeof delta === 'number' ? (isImproved ? 'delta-bad' : 'delta-good') : ''
    const pct =
      typeof delta === 'number' && typeof startRank === 'number' && startRank !== 0
        ? (Math.abs(delta) / startRank) * 100
        : null
    return (
      <span className="inline-metric">
        <span>{current.toLocaleString()}</span>
        {typeof delta === 'number' && delta !== 0 ? (
          <span className={`${toneClass} plain-delta`}>
            {isImproved ? <ArrowUp size={13} /> : <ArrowDown size={13} />}
            <span>{Math.abs(delta).toLocaleString()}</span>
            {typeof pct === 'number' ? <span>{`(${pct.toFixed(1)}%)`}</span> : null}
          </span>
        ) : null}
      </span>
    )
  }

  const importKeepaFile = async (
    file: File,
    target: 'today' | 'yesterday',
    archiveCurrent = false,
  ) => {
    setStatus(`正在解析 ${file.name} ...`)
    try {
      const rows = await readKeepaRows(file)
      const parsed = parseKeepa(rows)
      if (!parsed.length) throw new Error('Keepa 文件没有识别到有效 ASIN，请检查文件内容。')

      const date = getImportDate(file.name)
      const previousRows = hasCurrentBuyBoxSnapshot ? keepaRows : []
      const nextBoard = buildBuyBoxBoard(monitorRows, mappingRows, parsed, previousRows, date)
      if (target === 'yesterday') {
        const yesterdayBoard = buildBuyBoxBoard(monitorRows, mappingRows, parsed, [], date)
        setYesterdayBuyBox(yesterdayBoard)
        setTodayBuyBox(emptyBuyBoxBoard())
      } else {
        if (archiveCurrent && todayBuyBox.date) setYesterdayBuyBox(todayBuyBox)
        setTodayBuyBox(nextBoard)
      }
      setKeepaRows(parsed)
      setHistory((current) => mergeRecentHistory(current, parsed, date))

      const missingPrice = parsed.filter((row) => row.price === null).length
      const missingRank = parsed.filter((row) => row.rank === null).length
      const missingNewCurrent = parsed.filter((row) => row.newCurrent === null).length
      setUploadSummary({
        kind: 'keepa',
        fileName: file.name,
        imported: parsed.length,
        notes: [
          `识别到 ${parsed.length} 条 ASIN`,
          `New: Current 为空 ${missingNewCurrent} 条`,
          `${target === 'yesterday' ? '昨日基准' : '今日'} Buy Box 丢失 ${nextBoard.lost.length} 条`,
          `${target === 'yesterday' ? '昨日基准' : '今日'} Buy Box 恢复 ${target === 'yesterday' ? 0 : nextBoard.recovered.length} 条`,
          `缺少价格 ${missingPrice} 条、缺少排名 ${missingRank} 条`,
        ],
        errors: [],
      })
      setStatus(
        target === 'yesterday'
          ? `已将 ${parsed.length} 条 Keepa 数据保存为昨日基准：丢失 ${nextBoard.lost.length} 条。`
          : `已导入 ${parsed.length} 条 Keepa 数据：今日丢失 ${nextBoard.lost.length} 条，恢复 ${nextBoard.recovered.length} 条。`,
      )
    } catch (error) {
      setUploadSummary({
        kind: 'keepa',
        fileName: file.name,
        imported: 0,
        notes: [],
        errors: [error instanceof Error ? error.message : 'Keepa 文件解析失败，请检查导出字段。'],
      })
      setStatus(error instanceof Error ? error.message : 'Keepa 文件解析失败，请检查导出字段。')
    }
  }

  const handleUpload = async (kind: UploadKind, file: File | null) => {
    if (!file) return
    if (kind === 'keepa') {
      const mode = !hasCurrentBuyBoxSnapshot ? 'initial' : todayBuyBox.date ? 'daily' : 'from-yesterday'
      setPendingKeepaUpload({ file, mode })
      setStatus(
        mode === 'initial'
          ? '首次导入，请选择这份 Keepa 数据作为昨日基准或今日数据。'
          : mode === 'from-yesterday'
            ? '检测到已保存的昨日基准，请确认更新今日数据。'
            : '检测到已有今日快照，请选择是否保存为昨日数据。',
      )
      return
    }
    setStatus(`正在解析 ${file.name} ...`)
    if (kind === 'mapping') {
      const rows = await readWorkbookRows(file)
      const parsed = parseMapping(rows)
      setMappingRows(parsed)
      setUploadSummary({
        kind,
        fileName: file.name,
        imported: parsed.length,
        notes: [
          `识别到 ${parsed.length} 条映射信息`,
          `可用于补充平台 SKU 对应的运营、组别、账号`,
        ],
        errors: parsed.length ? [] : ['映射文件没有识别到有效的 平台SKU 行，请检查表头。'],
      })
      setStatus(`已导入 ${parsed.length} 条映射信息。`)
    }
    if (kind === 'monitor') {
      const rows = await readWorkbookRows(file)
      const parsed = parseMonitor(rows)
      setMonitorRows(parsed)
      setSelectedAsin(parsed[0]?.asin ?? '')
      setUploadSummary({
        kind,
        fileName: file.name,
        imported: parsed.length,
        notes: [
          `识别到 ${parsed.length} 条 SKU / ASIN 监控关系`,
          `监控清单中自带的运营信息会优先使用`,
        ],
        errors: parsed.length ? [] : ['监控清单没有识别到有效的 SKU 和 ASIN，请检查表头。'],
      })
      setStatus(`已导入 ${parsed.length} 条 SKU / ASIN 监控关系。`)
    }
  }

  const continueKeepaUpload = async (target: 'today' | 'yesterday', archiveCurrent = false) => {
    const pending = pendingKeepaUpload
    if (!pending) return
    setPendingKeepaUpload(null)
    await importKeepaFile(pending.file, target, archiveCurrent)
  }

  const cancelKeepaUpload = () => {
    setPendingKeepaUpload(null)
    setStatus('已取消本次 Keepa 上传，当前今日和昨日数据未改变。')
  }

  const resetToSeed = () => {
    const seedKeepaRows = normalizeKeepaRows(typedSeed.keepaRows)
    setKeepaRows(seedKeepaRows)
    setMappingRows(typedSeed.mappingRows)
    setMonitorRows(typedSeed.monitorRows)
    setHistory(keepRecentFiveDays(initialHistory(seedKeepaRows)))
    setTodayBuyBox(emptyBuyBoxBoard())
    setYesterdayBuyBox(emptyBuyBoxBoard())
    setPendingKeepaUpload(null)
    setSelectedAsin(typedSeed.monitorRows[0]?.asin ?? '')
    setEditingIndex(null)
    setStatus('已恢复为原 Excel 自动提取的数据。')
  }

  const startEdit = (index: number) => {
    setEditingIndex(index)
    if (editMode === 'monitor') setMonitorForm(monitorRows[index] ?? emptyMonitor)
    else setMappingForm(mappingRows[index] ?? emptyMapping)
  }

  const startAdd = () => {
    setEditingIndex(null)
    setMonitorForm(emptyMonitor)
    setMappingForm(emptyMapping)
  }

  const saveEdit = () => {
    if (editMode === 'monitor') {
      if (!monitorForm.sku || !monitorForm.asin) {
        setStatus('监控清单必须填写平台 SKU 和 ASIN。')
        return
      }
      setMonitorRows((rows) => {
        if (editingIndex === null) return [monitorForm, ...rows]
        return rows.map((row, index) => (index === editingIndex ? monitorForm : row))
      })
      setSelectedAsin(monitorForm.asin)
    } else {
      if (!mappingForm.sku) {
        setStatus('映射信息必须填写平台 SKU。')
        return
      }
      setMappingRows((rows) => {
        if (editingIndex === null) return [mappingForm, ...rows]
        return rows.map((row, index) => (index === editingIndex ? mappingForm : row))
      })
    }
    setEditingIndex(null)
    setStatus('已保存在线修改。')
  }

  const deleteEditRow = (index: number) => {
    if (editMode === 'monitor') setMonitorRows((rows) => rows.filter((_, rowIndex) => rowIndex !== index))
    else setMappingRows((rows) => rows.filter((_, rowIndex) => rowIndex !== index))
    setEditingIndex(null)
    setStatus('已删除记录。')
  }

  const toggleAlertGroup = (key: string) => {
    setCollapsedAlerts((current) => ({ ...current, [key]: !current[key] }))
  }

  const jumpToMapping = () => {
    setEditMode('mapping')
    setMaintenancePanel('mapping')
    setEditingIndex(null)
    setStatus('已打开映射信息维护区，可直接补充后保存。')
    window.setTimeout(() => editorPanelRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 80)
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <BarChart3 size={24} />
          <div>
            <h1>BuyBox Monitor</h1>
            <span>价格 · 排名 · 竞对预警</span>
          </div>
        </div>

        <div className="upload-stack">
          {(Object.keys(uploadLabels) as UploadKind[]).map((kind) => (
            <div className="upload-row" key={kind}>
              <label className="upload-button">
                <Upload size={16} />
                <span>{uploadLabels[kind]}</span>
                <input accept=".xlsx,.xls,.csv" type="file" onChange={(event) => { const file = event.currentTarget.files?.[0] ?? null; event.currentTarget.value = ''; void handleUpload(kind, file) }} />
              </label>
              <button className="icon-button" title="下载模板" type="button" onClick={() => downloadTemplate(kind)}>
                <Download size={16} />
              </button>
            </div>
          ))}
        </div>

        <section className="settings-panel">
          <div className="panel-title"><Bell size={17} />预警阈值</div>
          <label>价格变动 ≥ {priceAlert}%<input max="50" min="1" type="range" value={priceAlert} onChange={(event) => setPriceAlert(Number(event.target.value))} /></label>
          <label>排名下滑 ≥ {rankAlert}%<input max="200" min="5" step="5" type="range" value={rankAlert} onChange={(event) => setRankAlert(Number(event.target.value))} /></label>
        </section>

        <button className="reset-button" type="button" onClick={resetToSeed}><RotateCcw size={16} />恢复原表数据</button>
        <section className="nav-panel">
          <button className={maintenancePanel === 'monitor' ? 'active-nav' : ''} type="button" onClick={() => { setEditMode('monitor'); setMaintenancePanel((current) => current === 'monitor' ? null : 'monitor'); startAdd() }}>SKU 监控清单</button>
          <button className={maintenancePanel === 'mapping' ? 'active-nav' : ''} type="button" onClick={() => { setEditMode('mapping'); setMaintenancePanel((current) => current === 'mapping' ? null : 'mapping'); startAdd() }}>映射信息</button>
        </section>
        <p className="status-text">{status}</p>
        <section className="sidebar-note-panel">
          <div className="panel-title"><AlertTriangle size={16} />缺少 Keepa</div>
          <strong>{missingKeepaItems.length}</strong>
          <button className="reset-button" type="button" onClick={() => exportAlertItems('缺少Keepa', missingKeepaItems)}>下载缺失清单</button>
        </section>
        <section className="sidebar-info-panel">
          <div className="panel-title"><FileSpreadsheet size={16} />规则归类</div>
          <div className="sidebar-metric-list">
            <div className="sidebar-metric-item"><span>监控清单直连</span><strong>{ruleStats.direct}</strong></div>
            <div className="sidebar-metric-item"><span>映射补全</span><strong>{ruleStats.mapped}</strong></div>
            <div className="sidebar-metric-item"><span>标题/品牌识别</span><strong>{ruleStats.keepaMatched}</strong></div>
            <div className="sidebar-metric-item"><span>待补规则</span><strong>{ruleStats.unresolved}</strong></div>
          </div>
        </section>
        <section className="sidebar-info-panel">
          <div className="panel-title"><Upload size={16} />上传后说明</div>
          <div className="sidebar-upload-summary">
            <div className="sidebar-upload-count">
              <span>本次导入</span>
              <strong>{uploadSummary.imported}</strong>
              <p>{uploadSummary.fileName ? `${uploadLabels[uploadSummary.kind]} · ${uploadSummary.fileName}` : '等待上传文件'}</p>
            </div>
            <div className="sidebar-upload-block">
              <span>解析说明</span>
              {uploadSummary.notes.length ? <ul className="report-list compact-report-list">{uploadSummary.notes.map((note) => <li key={note}>{note}</li>)}</ul> : <p>上传后会在这里显示识别结果。</p>}
            </div>
            <div className="sidebar-upload-block">
              <span>报错说明</span>
              {uploadSummary.errors.length ? <ul className="report-list report-error compact-report-list">{uploadSummary.errors.map((error) => <li key={error}>{error}</li>)}</ul> : <p>当前没有上传报错。</p>}
            </div>
          </div>
        </section>
      </aside>

        <section className="workspace">
        <header className="topbar">
          <div className="topbar-intro">
            <span className="eyebrow">运营检索台</span>
            <h2>按运营、SKU、ASIN 快速定位监控关系</h2>
          </div>
          <div className="filter-grid">
            <div className="search-box"><Search size={18} /><input list="owner-options" placeholder="检索运营/人名" value={ownerQuery} onChange={(event) => setOwnerQuery(event.target.value)} /><datalist id="owner-options">{ownerOptions.map((owner) => <option key={owner} value={owner} />)}</datalist></div>
            <div className="search-box"><input list="sku-options" placeholder="检索 SKU" value={skuQuery} onChange={(event) => setSkuQuery(event.target.value)} /><datalist id="sku-options">{skuOptions.map((sku) => <option key={sku} value={sku} />)}</datalist></div>
            <div className="search-box"><input list="asin-options" placeholder="检索 ASIN" value={asinQuery} onChange={(event) => setAsinQuery(event.target.value)} /><datalist id="asin-options">{asinOptions.map((asin) => <option key={asin} value={asin} />)}</datalist></div>
            <div className="search-box"><input placeholder="品牌/标题/组别/账号" value={keywordQuery} onChange={(event) => setKeywordQuery(event.target.value)} /></div>
          </div>
          <div className="file-note"><FileSpreadsheet size={17} />纯规则检索，本地解析，不调用 token</div>
        </header>

        <section className="stat-grid">{stats.map((stat) => <div className="stat-card" key={stat.label}><span>{stat.label}</span><strong>{stat.value.toLocaleString()}</strong></div>)}</section>

        <section className="results-stack">
          <div className="table-panel">
            <div className="section-heading"><h2>检索结果</h2><span>{filteredRows.length} 条</span></div>
            <div className="data-table-wrap">
              <table className="data-table">
                <thead><tr><th>运营</th><th>SKU</th><th>类型</th><th>品牌</th><th>ASIN</th><th>归类规则</th><th>价格</th><th>排名</th></tr></thead>
                <tbody>{visibleRows.map((row, index) => {
                  const previous = visibleRows[index - 1]
                  const next = visibleRows[index + 1]
                  const isGroupStart = !previous || previous.sku !== row.sku
                  const isGroupEnd = !next || next.sku !== row.sku
                  const typeClass = normalized(row.asinType).includes('kmasin') ? 'type-km' : normalized(row.asinType).includes('竞对') ? 'type-competitor' : 'type-neutral'
                  const asinHistory = historyByAsin.get(normalized(row.asin))
                  const priceChange = getMetricChange(asinHistory, 'price')
                  const rankChange = getMetricChange(asinHistory, 'rank')
                  return <tr className={`${row.asin === selectedAsin ? 'selected-row' : ''} ${isGroupStart ? 'sku-group-start' : ''} ${isGroupEnd ? 'sku-group-end' : ''}`} key={`${row.sku}-${row.asin}`} onClick={() => setSelectedAsin(row.asin)}><td>{row.owner || '-'}</td><td className="sku-cell">{row.sku}</td><td><span className={`type-tag ${typeClass}`}>{row.asinType || '-'}</span></td><td>{row.keepa?.brand || '-'}</td><td className="asin-cell">{row.asin}</td><td><span className="rule-badge">{row.ruleSource}</span></td><td>{renderMetric(row.keepa?.price, priceChange, 'price')}</td><td>{renderMetric(row.keepa?.rank, rankChange, 'rank')}</td></tr>
                })}</tbody>
              </table>
            </div>
          </div>

          <div className="detail-panel">
            <div className="section-heading"><h2>ASIN 详情</h2><span>{selectedAsin || '未选择'}</span></div>
            {selectedRows[0] ? <div className="detail-stack"><div><span className="eyebrow">商品</span><h3>{selectedRows[0].keepa?.title || selectedRows[0].asin}</h3></div>{getPrimaryImageUrl(selectedRows[0].keepa?.image) ? <div className="detail-image-wrap"><img alt={selectedRows[0].asin} className="detail-image" src={getPrimaryImageUrl(selectedRows[0].keepa?.image)} /></div> : null}<dl className="detail-list"><div><dt>运营</dt><dd>{selectedRows[0].owner || '-'}</dd></div><div><dt>SKU</dt><dd>{selectedRows[0].sku}</dd></div><div><dt>品牌</dt><dd>{selectedRows[0].keepa?.brand || '-'}</dd></div><div><dt>归类规则</dt><dd>{selectedRows[0].ruleSource}</dd></div><div><dt>New: Current</dt><dd>{typeof selectedRows[0].keepa?.newCurrent === 'number' ? selectedRows[0].keepa.newCurrent.toFixed(2) : '-'}</dd></div><div><dt>Coupon</dt><dd>{selectedRows[0].keepa?.coupon || '-'}</dd></div></dl></div> : <p className="empty-state">点击左侧结果查看 ASIN。</p>}
          </div>
        </section>

        <section className="price-focus-panel">
          <div className="section-heading"><h2>价格异常看板</h2><span>{priceAlertRows.length} 条竞对变化</span><button className="action-pill" type="button" onClick={() => exportBoardItems('价格异常', priceAlertRows)}>下载表格</button></div>
          <div className="price-filter-grid">
            <div className="search-box"><input list="owner-options" placeholder="筛选运营" value={priceViewOwner} onChange={(event) => setPriceViewOwner(event.target.value)} /></div>
            <div className="search-box"><input list="sku-options" placeholder="筛选 SKU" value={priceViewSku} onChange={(event) => setPriceViewSku(event.target.value)} /></div>
            <div className="search-box"><input placeholder="筛选品类/类型" value={priceViewCategory} onChange={(event) => setPriceViewCategory(event.target.value)} /></div>
          </div>
          <div className="price-chart-wrap">
            <ResponsiveContainer height={260} width="100%">
              <RechartsBarChart data={priceAlertRows.slice().sort((a, b) => (b.changePct ?? 0) - (a.changePct ?? 0)).slice(0, 15)}>
                <CartesianGrid stroke="#e6ebf2" vertical={false} />
                <XAxis dataKey="asin" tickLine={false} interval={0} angle={-18} textAnchor="end" height={62} />
                <YAxis tickLine={false} />
                <Tooltip formatter={(value) => typeof value === 'number' ? `${value > 0 ? '+' : ''}${value.toFixed(2)}` : String(value ?? '')} labelFormatter={(label) => {
                  const row = priceAlertRows.find((item) => item.asin === label)
                  return row ? `竞对ASIN: ${label} | 运营: ${row.owner || '-'} | SKU: ${row.kmSku || row.sku || '-'} | 类型: ${row.category || '-'}` : `竞对ASIN: ${label}`
                }} />
                <Bar dataKey="delta" fill="#d97706" radius={[4, 4, 0, 0]} />
              </RechartsBarChart>
            </ResponsiveContainer>
          </div>
          <div className="mini-table-wrap">
            <table className="data-table mini-table">
              <thead><tr><th>运营</th><th>SKU</th><th>类型</th><th>ASIN</th><th><div className="sort-menu-wrap"><button className="table-sort" type="button" onClick={() => setPriceSortMenu((current) => current === 'price' ? null : 'price')}>价格 <ArrowUpDown size={13} /></button>{priceSortMenu === 'price' ? <div className="sort-menu"><button type="button" onClick={() => { setPriceSortBy('price-diff-desc'); setPriceSortMenu(null) }}>价格变动最大</button><button type="button" onClick={() => { setPriceSortBy('price-diff-asc'); setPriceSortMenu(null) }}>价格变动最小</button></div> : null}</div></th><th><div className="sort-menu-wrap"><button className="table-sort" type="button" onClick={() => setPriceSortMenu((current) => current === 'rank' ? null : 'rank')}>排名 <ArrowUpDown size={13} /></button>{priceSortMenu === 'rank' ? <div className="sort-menu"><button type="button" onClick={() => { setPriceSortBy('rank-diff-desc'); setPriceSortMenu(null) }}>排名差异最大</button><button type="button" onClick={() => { setPriceSortBy('rank-diff-asc'); setPriceSortMenu(null) }}>排名差异最小</button></div> : null}</div></th></tr></thead>
              <tbody>{priceBoardGroups.slice(0, 20).flatMap((group) => {
                const head = group[0]
                const kmRow = <tr className="price-parent-row" key={`km-${head.kmAsin}`}><td>{head.owner || '-'}</td><td>{head.kmSku || '-'}</td><td><span className="type-tag type-km">KMASIN</span></td><td>{head.kmAsin || '-'}</td><td>{typeof head.kmPrice === 'number' ? `$${head.kmPrice.toFixed(2)}` : '-'}</td><td>{typeof head.kmRank === 'number' ? head.kmRank.toLocaleString() : '-'}</td></tr>
                const children = group.map((item) => <tr key={item.id}><td>{item.owner || '-'}</td><td>{item.kmSku || item.sku || '-'}</td><td><span className="type-tag type-competitor">竞对ASIN</span></td><td>{item.asin || '-'}</td><td>{renderBoardPriceMetric(item.current, item.delta, item.changePct)}</td><td>{renderBoardRankMetric(item.currentRank, item.rankDelta, item.previousRank)}</td></tr>)
                return [kmRow, ...children]
              })}</tbody>
            </table>
          </div>
        </section>

        {maintenancePanel ? <section className="editor-panel">
          <div className="section-heading editor-heading">
            <div className="tabs"><button className={editMode === 'monitor' ? 'active-tab' : ''} type="button" onClick={() => { setEditMode('monitor'); startAdd() }}>SKU 监控清单</button><button className={editMode === 'mapping' ? 'active-tab' : ''} type="button" onClick={() => { setEditMode('mapping'); startAdd() }}>映射信息</button></div>
            <div className="editor-actions"><button type="button" onClick={startAdd}><Plus size={16} />新增</button><button type="button" onClick={saveEdit}><Save size={16} />保存</button><button type="button" onClick={() => exportRows(editMode, editableRows)}><Download size={16} />导出</button></div>
          </div>
          <div className="edit-form">
            {editMode === 'monitor' ? <>
              <input placeholder="运营" value={monitorForm.owner} onChange={(e) => setMonitorForm({ ...monitorForm, owner: e.target.value })} />
              <input placeholder="组别" value={monitorForm.group} onChange={(e) => setMonitorForm({ ...monitorForm, group: e.target.value })} />
              <input placeholder="账号" value={monitorForm.account} onChange={(e) => setMonitorForm({ ...monitorForm, account: e.target.value })} />
              <input placeholder="平台SKU" value={monitorForm.sku} onChange={(e) => setMonitorForm({ ...monitorForm, sku: e.target.value })} />
              <input placeholder="ASIN分类" value={monitorForm.asinType} onChange={(e) => setMonitorForm({ ...monitorForm, asinType: e.target.value })} />
              <input placeholder="ASIN" value={monitorForm.asin} onChange={(e) => setMonitorForm({ ...monitorForm, asin: e.target.value })} />
              <input placeholder="备注" value={monitorForm.note} onChange={(e) => setMonitorForm({ ...monitorForm, note: e.target.value })} />
            </> : <>
              <input placeholder="平台SKU" value={mappingForm.sku} onChange={(e) => setMappingForm({ ...mappingForm, sku: e.target.value })} />
              <input placeholder="系统SKU" value={mappingForm.systemSku} onChange={(e) => setMappingForm({ ...mappingForm, systemSku: e.target.value })} />
              <input placeholder="运营" value={mappingForm.owner} onChange={(e) => setMappingForm({ ...mappingForm, owner: e.target.value })} />
              <input placeholder="小组" value={mappingForm.group} onChange={(e) => setMappingForm({ ...mappingForm, group: e.target.value })} />
              <input placeholder="店铺别名" value={mappingForm.account} onChange={(e) => setMappingForm({ ...mappingForm, account: e.target.value })} />
            </>}
          </div>
          <div className="mini-table-wrap"><table className="data-table mini-table"><thead><tr>{editMode === 'monitor' ? <><th>运营</th><th>SKU</th><th>ASIN</th><th>类型</th><th>操作</th></> : <><th>平台SKU</th><th>系统SKU</th><th>运营</th><th>小组</th><th>操作</th></>}</tr></thead><tbody>{editableRows.slice(0, 80).map((row, index) => <tr key={editMode === 'monitor' ? `${(row as MonitorRow).sku}-${(row as MonitorRow).asin}-${index}` : `${(row as MappingRow).sku}-${index}`}><td>{editMode === 'monitor' ? (row as MonitorRow).owner : (row as MappingRow).sku}</td><td>{editMode === 'monitor' ? (row as MonitorRow).sku : (row as MappingRow).systemSku}</td><td>{editMode === 'monitor' ? (row as MonitorRow).asin : (row as MappingRow).owner}</td><td>{editMode === 'monitor' ? (row as MonitorRow).asinType : (row as MappingRow).group}</td><td><button className="row-icon" type="button" onClick={() => startEdit(index)}><Edit3 size={14} /></button><button className="row-icon danger" type="button" onClick={() => deleteEditRow(index)}><Trash2 size={14} /></button></td></tr>)}</tbody></table></div>
        </section> : null}

        <section className="chart-grid">
          <div className="chart-panel"><div className="section-heading"><h2>价格趋势</h2><span>Price</span></div><ResponsiveContainer height={230} width="100%"><LineChart data={selectedHistory}><CartesianGrid stroke="#e6ebf2" vertical={false} /><XAxis dataKey="date" tickLine={false} /><YAxis tickLine={false} width={54} /><Tooltip /><Line dataKey="price" stroke="#1f7a6d" strokeWidth={2.5} type="monotone" /></LineChart></ResponsiveContainer></div>
          <div className="chart-panel"><div className="section-heading"><h2>排名趋势</h2><span>Rank</span></div><ResponsiveContainer height={230} width="100%"><AreaChart data={selectedHistory}><CartesianGrid stroke="#e6ebf2" vertical={false} /><XAxis dataKey="date" tickLine={false} /><YAxis tickLine={false} width={66} /><Tooltip /><Area dataKey="rank" fill="#dfeeea" stroke="#5067a3" strokeWidth={2.5} type="monotone" /></AreaChart></ResponsiveContainer></div>
        </section>

        <section className="alerts-panel">
          <div className="section-heading"><h2>预警中心</h2><span>{alerts.length} 条</span></div>
          <div className="alert-groups">
            {alertGroups.filter((group) => !['missing-keepa', 'buybox', 'duplicate'].includes(group.key)).length ? alertGroups.filter((group) => !['missing-keepa', 'buybox', 'duplicate'].includes(group.key)).map((group) => <div className="alert-group" key={group.key}><div className="alert-group-head"><h3>{group.title}</h3><span>{group.count} 条</span><div className="alert-group-actions"><button className="action-pill" type="button" onClick={() => toggleAlertGroup(group.key)}>{collapsedAlerts[group.key] ? '展开明细' : '收起明细'}</button><button className="action-pill icon-pill" title="导出清单" type="button" onClick={() => exportAlertItems(group.title, group.items)}><Download size={14} /></button>{group.key === 'missing-mapping' ? <button className="action-pill" type="button" onClick={jumpToMapping}>去补映射</button> : null}</div></div>{collapsedAlerts[group.key] ? null : <div className="alert-list">{group.items.map((item) => <div className="alert-item" key={item.id}><AlertTriangle size={18} /><div className="alert-content"><span>{item.message}</span></div></div>)}</div>}</div>) : <p className="empty-state">当前没有超过阈值的变化，也没有数据缺口。</p>}
            <div className="alert-group">
              <div className="alert-group-head">
                <h3>重复监控</h3>
                <span>{duplicateItems.length} 条</span>
                <div className="alert-group-actions">
                  <button className="action-pill" type="button" onClick={() => exportMonitorWithDuplicateMarks(monitorRows)}>下载删减后重传表</button>
                </div>
              </div>
            </div>
            <div className="alert-group">
              <div className="alert-group-head">
                <h3>近5天排名趋势预警</h3>
                <span>{rankTrendRows.length} 条</span>
              </div>
              <div className="mini-table-wrap">
                <table className="data-table mini-table">
                  <thead><tr><th>运营</th><th>SKU</th><th>ASIN</th><th>类型</th><th>趋势</th><th>起始排名</th><th>当前排名</th><th>排名</th><th>连续天数</th></tr></thead>
                  <tbody>{rankTrendRows.map((row) => {
                    const improved = row.endRank < row.startRank
                    return <tr key={`${row.sku}-${row.asin}-${row.direction}`}><td>{row.owner || '-'}</td><td>{row.sku}</td><td>{row.asin}</td><td>{row.asinType}</td><td><span className={row.direction === 'down' ? 'delta-good' : 'delta-bad'}>{row.direction === 'down' ? '持续下滑' : '持续上升'}</span></td><td>{row.startRank.toLocaleString()}</td><td>{row.endRank.toLocaleString()}</td><td><span className={improved ? 'delta-bad plain-delta' : 'delta-good plain-delta'}>{improved ? <ArrowUp size={13} /> : <ArrowDown size={13} />}{`${Math.abs(row.endRank - row.startRank).toLocaleString()} (${row.changePct.toFixed(1)}%)`}</span></td><td>{row.days}</td></tr>
                  })}</tbody>
                </table>
              </div>
            </div>
          </div>
        </section>

        <section className="buybox-day-grid" aria-label="每日 Buy Box 监测">
          <section className="buybox-day-panel buybox-today-panel">
            <div className="section-heading buybox-day-heading"><div><span className="eyebrow">今日维度</span><h2>Buy Box 监测</h2></div><span>{todayBuyBox.date || '等待上传'}</span></div>
            <BuyBoxStatusSection title="Buy Box 丢失" items={todayBuyBox.lost} tone="lost" />
            <BuyBoxStatusSection title="Buy Box 恢复" items={todayBuyBox.recovered} tone="recovered" />
          </section>
          <section className="buybox-day-panel buybox-yesterday-panel">
            <div className="section-heading buybox-day-heading"><div><span className="eyebrow">昨日维度</span><h2>Buy Box 监测</h2></div><span>{yesterdayBuyBox.date || '尚未保存'}</span></div>
            <BuyBoxStatusSection title="Buy Box 丢失" items={yesterdayBuyBox.lost} tone="lost" />
            <BuyBoxStatusSection title="Buy Box 恢复" items={yesterdayBuyBox.recovered} tone="recovered" />
          </section>
        </section>
      </section>

      {pendingKeepaUpload ? (
        <div className="modal-backdrop">
          <section aria-describedby="keepa-save-description" aria-labelledby="keepa-save-title" aria-modal="true" className="save-snapshot-dialog" role="dialog">
            <div className="dialog-heading"><Save size={20} /><div><span className="eyebrow">上传每日 Keepa 数据</span><h2 id="keepa-save-title">{pendingKeepaUpload.mode === 'initial' ? '首次导入：这份数据属于哪一天？' : pendingKeepaUpload.mode === 'from-yesterday' ? '使用昨日基准更新今日数据？' : '是否保存昨日数据？'}</h2></div></div>
            <p id="keepa-save-description">
              {pendingKeepaUpload.mode === 'initial'
                ? <>当前没有可比较的历史快照。请选择将 <strong>{pendingKeepaUpload.file.name}</strong> 保存为昨日基准，或直接作为今日数据。</>
                : pendingKeepaUpload.mode === 'from-yesterday'
                  ? <>昨日基准 <strong>{yesterdayBuyBox.date}</strong> 已保存。导入 <strong>{pendingKeepaUpload.file.name}</strong> 后，系统将比较昨日与今日状态并生成恢复数据。</>
                  : <>保存后，当前今日的丢失和恢复数据会固定到右侧昨日栏，再使用 <strong>{pendingKeepaUpload.file.name}</strong> 更新左侧今日栏。</>}
            </p>
            <div className="dialog-actions">
              <button className="dialog-button dialog-button-cancel" type="button" onClick={cancelKeepaUpload}>取消上传</button>
              {pendingKeepaUpload.mode === 'initial' ? (
                <>
                  <button className="dialog-button" type="button" onClick={() => void continueKeepaUpload('today')}>作为今日数据</button>
                  <button autoFocus className="dialog-button dialog-button-primary" type="button" onClick={() => void continueKeepaUpload('yesterday')}><Save size={16} />作为昨日基准</button>
                </>
              ) : pendingKeepaUpload.mode === 'from-yesterday' ? (
                <button autoFocus className="dialog-button dialog-button-primary" type="button" onClick={() => void continueKeepaUpload('today')}><Upload size={16} />更新今日数据</button>
              ) : (
                <>
                  <button className="dialog-button" type="button" onClick={() => void continueKeepaUpload('today')}>不保存，继续</button>
                  <button autoFocus className="dialog-button dialog-button-primary" type="button" onClick={() => void continueKeepaUpload('today', true)}><Save size={16} />保存昨日数据并继续</button>
                </>
              )}
            </div>
          </section>
        </div>
      ) : null}
    </main>
  )
}

export default App
