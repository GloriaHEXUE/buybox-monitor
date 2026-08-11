import { useEffect, useMemo, useState } from 'react'
import {
  AlertTriangle,
  ArrowDown,
  ArrowUp,
  BarChart3,
  Bell,
  ChevronDown,
  CircleX,
  Database,
  Download,
  Edit3,
  FileSpreadsheet,
  Plus,
  RefreshCw,
  Save,
  Search,
  Trash2,
  Upload,
} from 'lucide-react'
import {
  Area,
  AreaChart,
  CartesianGrid,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import { useRef } from 'react'
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

type KeepaSnapshotRow = Pick<KeepaRow, 'asin' | 'newCurrent' | 'rank'>

type HistoryPoint = {
  date: string
  asin: string
  price: number | null
  rank: number | null
  buyBox: string
}

type UploadKind = 'keepa' | 'mapping' | 'monitor'
type EditMode = 'monitor' | 'mapping'
type WorkspacePage = 'dashboard' | 'data-update' | 'buybox' | 'add-monitor'

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

type KeepaUploadReport = {
  status: 'idle' | 'processing' | 'success' | 'error'
  fileName: string
  date: string
  imported: number
  notes: string[]
  errors: string[]
}

type KeepaUploadReports = Record<'yesterday' | 'today', KeepaUploadReport>

type ResultFilterKey = 'owner' | 'sku' | 'asinType' | 'brand' | 'asin' | 'price' | 'todayRank' | 'yesterdayRank'

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
  detail?: string
  group?: string
  account?: string
}

type AlertGroup = {
  key: string
  title: string
  count: number
  items: AlertItem[]
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
}

type BuyBoxDayBoard = {
  date: string
  lost: BuyBoxStatusItem[]
  recovered: BuyBoxStatusItem[]
}

type ResultColumnKey = ResultFilterKey

type SourceReport = {
  fileName: string
  imported: number
  updatedAt: string
  source: 'seed' | 'upload'
}

type SourceReports = Record<EditMode, SourceReport>

type StoredState = {
  keepaRows: KeepaRow[]
  yesterdayKeepaRows: KeepaSnapshotRow[]
  mappingRows: MappingRow[]
  monitorRows: MonitorRow[]
  onlineRows: MonitorRow[]
  sourceReports: SourceReports
  history: HistoryPoint[]
  todayBuyBox: BuyBoxDayBoard
  yesterdayBuyBox: BuyBoxDayBoard
  keepaUploadReports: KeepaUploadReports
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

const workspacePageLabels: Record<WorkspacePage, string> = {
  dashboard: '运营看板',
  'data-update': '数据更新',
  buybox: 'Buy Box 丢失|恢复',
  'add-monitor': '添加监控ASIN',
}

const downloadLabels: Record<UploadKind, string> = {
  keepa: '下载 Keepa 模板',
  mapping: '导出全部映射',
  monitor: '下载监控包',
}

const monitorHeaders = ['运营', '组别', '账号', '平台SKU', 'Bundle主SKU', 'ASIN分类', 'ASIN', '备注']
const mappingHeaders = ['平台SKU', '系统SKU', '运营', '小组', '店铺别名']
const keepaHeaders = ['ASIN', 'Title', 'Brand', 'New: Current', 'Sales Rank: Subcategory Sales Ranks', 'Buy Box Seller', 'Coupon', 'Prime Price', 'Image']
const keepaFieldCandidates = {
  asin: ['ASIN', 'asin', 'Product Codes: ASIN', 'Product Codes ASIN'],
  title: ['Title', '标题', '商品标题', 'Parent Title'],
  brand: ['Brand', '品牌', 'Manufacturer'],
  price: ['Buy Box: Current', 'Amazon: Current', 'New: Current', 'Price', '今日价格', '标价'],
  newCurrent: ['New: Current'],
  rank: ['Sales Rank: Subcategory Sales Ranks'],
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

const asSubcategoryRank = (value: unknown) => {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const text = String(value ?? '').trim()
  if (!text || text === '-') return null
  const rankedValue = text.match(/#\s*([\d,]+)/)?.[1]
  if (rankedValue) return asNumber(rankedValue)
  return asNumber(text)
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

const getImageCandidates = (value: string | null | undefined) => {
  const raw = String(value ?? '').trim()
  if (!raw) return []
  return raw
    .split(/[\n,;|]/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => {
      if (/^https?:\/\//i.test(item)) return item
      if (/^\/\//.test(item)) return `https:${item}`
      if (/\.(jpg|jpeg|png|webp)/i.test(item)) return `https://m.media-amazon.com/images/I/${item.replace(/^\/+/, '')}`
      return ''
    })
    .filter(Boolean)
}

const getPrimaryImageUrl = (value: string | null | undefined) => getImageCandidates(value)[0] || ''

const readWorkbookRows = async (file: File, kind: EditMode): Promise<AnyRow[]> => {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { dense: true })
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
  const matrix = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(firstSheet, {
    header: 1,
    raw: false,
    defval: '',
  })
  const candidateGroups =
    kind === 'mapping'
      ? [['平台SKU', 'sku'], ['系统SKU'], ['店铺别名', '账号'], ['运营', '负责人'], ['小组', '组别']]
      : [['平台SKU', 'sku'], ['ASIN'], ['ASIN分类', '类型'], ['运营', 'owner']]
  const normalizedRowHitCount = (row: (string | number | boolean | null)[]) => {
    const cells = row.map((cell) => normalized(cell))
    return candidateGroups.reduce((count, group) => {
      const matched = group.some((candidate) => cells.some((cell) => cell === normalized(candidate) || cell.includes(normalized(candidate))))
      return count + (matched ? 1 : 0)
    }, 0)
  }
  const headerRowIndex = matrix.findIndex((row) => normalizedRowHitCount(row) >= (kind === 'mapping' ? 3 : 2))
  if (headerRowIndex === -1) {
    throw new Error(kind === 'mapping' ? '映射文件未识别到“平台SKU/店铺别名/运营/小组”等表头。' : '监控清单未识别到“平台SKU/ASIN/ASIN分类”等表头。')
  }
  const headers = (matrix[headerRowIndex] ?? []).map((cell) => String(cell ?? '').trim())
  return matrix
    .slice(headerRowIndex + 1)
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])))
    .filter((row) => Object.values(row).some((value) => String(value ?? '').trim()))
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
  if (indices.rank === -1) {
    throw new Error('Keepa 文件中未找到 Sales Rank: Subcategory Sales Ranks 列，请检查导出字段。')
  }

  return matrix.slice(1).map((row) => ({
    ASIN: indices.asin >= 0 ? row[indices.asin] : '',
    Title: indices.title >= 0 ? row[indices.title] : '',
    Brand: indices.brand >= 0 ? row[indices.brand] : '',
    Price: indices.price >= 0 ? row[indices.price] : '',
    'New: Current': row[indices.newCurrent],
    'Sales Rank: Subcategory Sales Ranks': indices.rank >= 0 ? row[indices.rank] : '',
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
        rank: asSubcategoryRank(pick(row, keepaFieldCandidates.rank)),
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
      group: String(pick(row, ['小组', '组别', 'Group', '部门']) || '').trim(),
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

const monitorRowKey = (row: Pick<MonitorRow, 'sku' | 'asinType' | 'asin'>) =>
  `${normalized(row.sku)}::${normalized(row.asinType)}::${normalized(row.asin)}`

const mergeMonitorRows = (sourceRows: MonitorRow[], onlineRows: MonitorRow[]) => {
  const sourceKeys = new Set(sourceRows.map(monitorRowKey))
  return [...sourceRows, ...onlineRows.filter((row) => !sourceKeys.has(monitorRowKey(row)))]
}

const applyMappingsToOnlineRows = (rows: MonitorRow[], mappings: MappingRow[]) => {
  const bySku = new Map(mappings.map((row) => [normalized(row.sku), row]))
  return rows.map((row) => {
    const mapping = bySku.get(normalized(row.sku))
    return mapping
      ? { ...row, owner: mapping.owner, group: mapping.group, account: mapping.account }
      : row
  })
}

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

const toKeepaSnapshotRows = (rows: Array<KeepaRow | KeepaSnapshotRow> | undefined): KeepaSnapshotRow[] =>
  (rows ?? []).map((row) => ({
    asin: String(row.asin ?? '').trim(),
    newCurrent: typeof row.newCurrent === 'number' && Number.isFinite(row.newCurrent) ? row.newCurrent : null,
    rank: typeof row.rank === 'number' && Number.isFinite(row.rank) ? row.rank : null,
  })).filter((row) => row.asin)

const compactKeepaRowsForStorage = (rows: KeepaRow[]): KeepaRow[] =>
  rows.map((row) => ({ ...row, image: '' }))

const getImportDate = (fileName: string) => {
  const fileDate = fileName.match(/\d{4}-\d{2}-\d{2}/)?.[0]
  return fileDate ?? new Date().toLocaleDateString('sv-SE')
}

const buildBuyBoxBoard = (
  monitorRows: MonitorRow[],
  mappingRows: MappingRow[],
  currentRows: Array<KeepaRow | KeepaSnapshotRow>,
  previousRows: Array<KeepaRow | KeepaSnapshotRow>,
  date: string,
  mode: 'snapshot' | 'changes',
): BuyBoxDayBoard => {
  const currentByAsin = new Map(currentRows.map((row) => [normalized(row.asin), row.newCurrent]))
  const previousByAsin = new Map(previousRows.map((row) => [normalized(row.asin), row.newCurrent]))
  const mappingBySku = new Map(mappingRows.map((row) => [normalized(row.sku), row]))
  const lost: BuyBoxStatusItem[] = []
  const recovered: BuyBoxStatusItem[] = []
  const seen = new Set<string>()

  for (const row of monitorRows) {
    if (normalized(row.asinType) !== 'kmasin') continue
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
    if (mode === 'snapshot') {
      if (isCurrentLost) lost.push(item)
      continue
    }

    if (!previousExists) continue
    if (!wasPreviouslyLost && isCurrentLost) lost.push(item)
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
  const fallback: StoredState = {
    keepaRows: seedKeepaRows,
    yesterdayKeepaRows: [],
    mappingRows: typedSeed.mappingRows,
    monitorRows: typedSeed.monitorRows,
    onlineRows: [],
    sourceReports: {
      monitor: { fileName: 'seedData.json', imported: typedSeed.monitorRows.length, updatedAt: '内置数据', source: 'seed' },
      mapping: { fileName: 'seedData.json', imported: typedSeed.mappingRows.length, updatedAt: '内置数据', source: 'seed' },
    },
    history: initialHistory(seedKeepaRows),
    todayBuyBox: emptyBuyBoxBoard(),
    yesterdayBuyBox: emptyBuyBoxBoard(),
    keepaUploadReports: emptyKeepaUploadReports(),
  }
  try {
    const raw = localStorage.getItem(storageKey)
    if (!raw) return fallback
    const stored = JSON.parse(raw) as Partial<StoredState>
    const storedKeepaRows = normalizeKeepaRows(stored.keepaRows)
    const storedYesterdayRows = stored.yesterdayKeepaRows
      ? toKeepaSnapshotRows(stored.yesterdayKeepaRows)
      : stored.yesterdayBuyBox?.date && !stored.todayBuyBox?.date
        ? toKeepaSnapshotRows(stored.keepaRows)
        : []
    const mappingRows = stored.mappingRows ?? fallback.mappingRows
    const monitorRows = stored.monitorRows ?? fallback.monitorRows
    const onlineRows = stored.onlineRows ?? fallback.onlineRows
    const combinedMonitorRows = mergeMonitorRows(monitorRows, onlineRows)
    const todayDate = String(stored.todayBuyBox?.date ?? '')
    const yesterdayDate = String(stored.yesterdayBuyBox?.date ?? '')
    return {
      keepaRows: storedKeepaRows,
      yesterdayKeepaRows: storedYesterdayRows,
      mappingRows,
      monitorRows,
      onlineRows,
      sourceReports: stored.sourceReports ?? fallback.sourceReports,
      history: stored.history?.length ? keepRecentFiveDays(stored.history) : initialHistory(storedKeepaRows),
      todayBuyBox: todayDate
        ? buildBuyBoxBoard(combinedMonitorRows, mappingRows, storedKeepaRows, storedYesterdayRows, todayDate, 'changes')
        : emptyBuyBoxBoard(),
      yesterdayBuyBox: yesterdayDate
        ? buildBuyBoxBoard(combinedMonitorRows, mappingRows, storedYesterdayRows, [], yesterdayDate, 'snapshot')
        : emptyBuyBoxBoard(),
      keepaUploadReports: {
        yesterday: normalizeKeepaUploadReport(stored.keepaUploadReports?.yesterday),
        today: normalizeKeepaUploadReport(stored.keepaUploadReports?.today),
      },
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

const emptyOnlineMonitor: MonitorRow = {
  ...emptyMonitor,
  asinType: '',
}

const emptyUploadSummary: UploadSummary = {
  kind: 'keepa',
  fileName: '',
  imported: 0,
  notes: [],
  errors: [],
}

const emptyKeepaUploadReport = (): KeepaUploadReport => ({
  status: 'idle',
  fileName: '',
  date: '',
  imported: 0,
  notes: [],
  errors: [],
})

const emptyKeepaUploadReports = (): KeepaUploadReports => ({
  yesterday: emptyKeepaUploadReport(),
  today: emptyKeepaUploadReport(),
})

const normalizeKeepaUploadReport = (report: KeepaUploadReport | undefined): KeepaUploadReport => ({
  status: report?.status ?? 'idle',
  fileName: String(report?.fileName ?? ''),
  date: String(report?.date ?? ''),
  imported: Number.isFinite(report?.imported) ? Number(report?.imported) : 0,
  notes: Array.isArray(report?.notes) ? report.notes : [],
  errors: Array.isArray(report?.errors) ? report.errors : [],
})

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

const exportMonitorTemplatePackage = (monitorRows: MonitorRow[], mappingRows: MappingRow[]) => {
  const templateSheet = XLSX.utils.aoa_to_sheet([
    monitorHeaders,
    ['', '', '', 'KM1B0901-06205-BK', '', '竞对ASIN', 'B08CF2V57W', '填写平台SKU后可参考映射信息匹配运营/店铺/小组'],
  ])
  const monitorSheet = XLSX.utils.json_to_sheet(
    monitorRows.map((row) => ({
      运营: row.owner,
      组别: row.group,
      账号: row.account,
      平台SKU: row.sku,
      Bundle主SKU: row.bundleSku,
      ASIN分类: row.asinType,
      ASIN: row.asin,
      备注: row.note,
    })),
  )
  const mappingSheet = XLSX.utils.json_to_sheet(
    mappingRows.map((row) => ({
      平台SKU: row.sku,
      系统SKU: row.systemSku,
      运营: row.owner,
      小组: row.group,
      店铺别名: row.account,
    })),
  )
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, templateSheet, '监控模板')
  XLSX.utils.book_append_sheet(workbook, monitorSheet, '全部监控清单')
  XLSX.utils.book_append_sheet(workbook, mappingSheet, '映射信息')
  XLSX.writeFile(workbook, 'SKU监控清单模板与映射信息.xlsx')
}

const exportOnlineImportTemplate = (mappingRows: MappingRow[]) => {
  const workbook = XLSX.utils.book_new()
  const templateHeaders = ['平台SKU', 'ASIN分类', 'ASIN', '运营', '组别', '账号', '备注']
  const templateSheet = XLSX.utils.aoa_to_sheet([
    templateHeaders,
    ['KM1B0901-06205-BK', '竞对ASIN', 'B08CF2V57W', '', '', '', '填写平台SKU后自动匹配运营、组别、账号；无映射则显示 NA'],
    ['', '', '', '', '', '', ''],
    ['', '', '', '', '', '', ''],
    ['', '', '', '', '', '', ''],
  ])
  templateSheet['!dataValidation'] = [
    {
      sqref: 'B2:B2000',
      type: 'list',
      allowBlank: 1,
      formulas: ['"竞对ASIN,KMASIN"'],
    },
  ] as unknown as never
  ;['D2', 'E2', 'F2', 'D3', 'E3', 'F3', 'D4', 'E4', 'F4', 'D5', 'E5', 'F5'].forEach((cell) => {
    const row = Number(cell.slice(1))
    const formulaMap: Record<string, string> = {
      D: `IFERROR(XLOOKUP($A${row},映射信息!$A:$A,映射信息!$C:$C),"NA")`,
      E: `IFERROR(XLOOKUP($A${row},映射信息!$A:$A,映射信息!$D:$D),"NA")`,
      F: `IFERROR(XLOOKUP($A${row},映射信息!$A:$A,映射信息!$E:$E),"NA")`,
    }
    const column = cell[0]
    templateSheet[cell] = { t: 'str', f: formulaMap[column] }
  })
  templateSheet['!cols'] = [{ wch: 24 }, { wch: 16 }, { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 18 }, { wch: 48 }]
  const mappingSheet = XLSX.utils.json_to_sheet(
    mappingRows.map((row) => ({
      平台SKU: row.sku,
      系统SKU: row.systemSku,
      运营: row.owner,
      小组: row.group,
      店铺别名: row.account,
    })),
  )
  mappingSheet['!cols'] = [{ wch: 24 }, { wch: 22 }, { wch: 18 }, { wch: 16 }, { wch: 18 }]
  XLSX.utils.book_append_sheet(workbook, templateSheet, '批量导入模板')
  XLSX.utils.book_append_sheet(workbook, mappingSheet, '映射信息')
  XLSX.writeFile(workbook, '批量导入平台SKU与监控ASIN模板.xlsx')
}

const exportRows = (mode: EditMode, rows: MonitorRow[] | MappingRow[], fileName?: string) => {
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
  sheet['!cols'] = mode === 'monitor'
    ? [{ wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 24 }, { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 24 }]
    : [{ wch: 24 }, { wch: 22 }, { wch: 18 }, { wch: 16 }, { wch: 18 }]
  sheet['!autofilter'] = { ref: sheet['!ref'] ?? 'A1:A1' }
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, mode === 'monitor' ? 'SKU监控清单' : '映射信息')
  XLSX.writeFile(workbook, fileName ?? `${mode === 'monitor' ? 'SKU监控清单' : '映射信息'}导出.xlsx`)
}

const exportMissingKeepaReport = (items: AlertItem[]) => {
  const data = items.map((item) => ({
    缺失类型: item.detail || '未分类',
    处理建议: item.detail?.includes('未出现在本次上传')
      ? '检查 Keepa 导出范围是否覆盖该 ASIN，或确认该 ASIN 是否仍需监控'
      : item.detail?.includes('价格与排名都缺失')
        ? '优先检查链接是否下架、抑制，或 Keepa 是否抓取异常'
        : item.detail?.includes('价格缺失')
          ? '检查 Buy Box / New Current 是否为空'
          : '检查排名字段是否为空或类目异常',
    运营: item.owner || '',
    SKU: item.sku || '',
    ASIN分类: item.category || '',
    ASIN: item.asin || '',
    说明: item.message,
  }))
  const summary = XLSX.utils.json_to_sheet(data)
  const grouped = data.reduce<Record<string, number>>((acc, row) => {
    acc[row.缺失类型] = (acc[row.缺失类型] ?? 0) + 1
    return acc
  }, {})
  const summarySheet = XLSX.utils.json_to_sheet(
    Object.entries(grouped).map(([type, count]) => ({ 缺失类型: type, 数量: count })),
  )
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, summarySheet, '分类汇总')
  XLSX.utils.book_append_sheet(workbook, summary, '缺失明细')
  XLSX.writeFile(workbook, '缺少Keepa分类处理表.xlsx')
}

const exportUnmatchedOnlineRows = (rows: MonitorRow[]) => {
  const sheet = XLSX.utils.json_to_sheet(
    rows.map((row) => ({
      平台SKU: row.sku,
      ASIN分类: row.asinType,
      ASIN: row.asin,
      备注: row.note,
      当前运营: row.owner,
      当前组别: row.group,
      当前账号: row.account,
      处理建议: '请先补充映射总表，再重新上传映射信息以自动匹配',
    })),
  )
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, '待补映射平台SKU')
  XLSX.writeFile(workbook, '在线新增待补映射平台SKU.xlsx')
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
  const [yesterdayKeepaRows, setYesterdayKeepaRows] = useState<KeepaSnapshotRow[]>(initialState.yesterdayKeepaRows)
  const [mappingRows, setMappingRows] = useState<MappingRow[]>(initialState.mappingRows)
  const [monitorRows, setMonitorRows] = useState<MonitorRow[]>(initialState.monitorRows)
  const [onlineRows, setOnlineRows] = useState<MonitorRow[]>(initialState.onlineRows)
  const [sourceReports, setSourceReports] = useState<SourceReports>(initialState.sourceReports)
  const [history, setHistory] = useState<HistoryPoint[]>(initialState.history)
  const [todayBuyBox, setTodayBuyBox] = useState<BuyBoxDayBoard>(initialState.todayBuyBox)
  const [yesterdayBuyBox, setYesterdayBuyBox] = useState<BuyBoxDayBoard>(initialState.yesterdayBuyBox)
  const [pendingKeepaUpload, setPendingKeepaUpload] = useState<PendingKeepaUpload | null>(null)
  const [ownerQuery, setOwnerQuery] = useState('')
  const [skuQuery, setSkuQuery] = useState('')
  const [asinQuery, setAsinQuery] = useState('')
  const [groupQuery, setGroupQuery] = useState('')
  const [accountQuery, setAccountQuery] = useState('')
  const [brandQuery, setBrandQuery] = useState('')
  const [selectedAsin, setSelectedAsin] = useState('')
  const [batchOnlineRows, setBatchOnlineRows] = useState<MonitorRow[]>(Array.from({ length: 5 }, () => ({ ...emptyOnlineMonitor })))
  const [priceAlert] = useState(12)
  const [rankAlert] = useState(35)
  const [openPages, setOpenPages] = useState<WorkspacePage[]>(['dashboard'])
  const [activePage, setActivePage] = useState<WorkspacePage>('dashboard')
  const [isRefreshingOnline, setIsRefreshingOnline] = useState(false)
  const [uploadSummary, setUploadSummary] = useState<UploadSummary>(emptyUploadSummary)
  const [keepaUploadReports, setKeepaUploadReports] = useState<KeepaUploadReports>(initialState.keepaUploadReports)
  const [resultFilters, setResultFilters] = useState<Record<ResultFilterKey, string>>({
    owner: '',
    sku: '',
    asinType: '',
    brand: '',
    asin: '',
    price: '',
    todayRank: '',
    yesterdayRank: '',
  })
  const [openResultFilter, setOpenResultFilter] = useState<ResultFilterKey | null>(null)
  const [resultColumnWidths, setResultColumnWidths] = useState<Record<ResultColumnKey, number>>({
    owner: 132,
    sku: 168,
    asinType: 104,
    brand: 148,
    asin: 150,
    price: 148,
    todayRank: 130,
    yesterdayRank: 130,
  })
  const [status, setStatus] = useState(
    `已自动载入原 Excel：${initialState.monitorRows.length} 条监控清单、${initialState.mappingRows.length} 条映射、${initialState.keepaRows.length} 条 Keepa。趋势仅保留最近 5 天。`,
  )

  useEffect(() => {
    const payload: StoredState = {
      keepaRows: compactKeepaRowsForStorage(keepaRows),
      yesterdayKeepaRows,
      mappingRows,
      monitorRows,
      onlineRows,
      sourceReports,
      history: [],
      todayBuyBox,
      yesterdayBuyBox,
      keepaUploadReports,
    }
    try {
      localStorage.setItem(storageKey, JSON.stringify(payload))
    } catch {
      setStatus('数据已在当前页面完成解析，但浏览器存储空间不足，刷新后可能无法保留。请清理浏览器缓存或减少数据后重新上传。')
    }
  }, [keepaRows, keepaUploadReports, mappingRows, monitorRows, onlineRows, sourceReports, todayBuyBox, yesterdayBuyBox, yesterdayKeepaRows])

  useEffect(() => {
    if (!pendingKeepaUpload) return undefined
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') cancelKeepaUpload()
    }
    document.addEventListener('keydown', closeOnEscape)
    return () => document.removeEventListener('keydown', closeOnEscape)
  }, [pendingKeepaUpload])

  const keepaByAsin = useMemo(() => new Map(keepaRows.map((row) => [normalized(row.asin), row])), [keepaRows])
  const yesterdayKeepaByAsin = useMemo(
    () => new Map(yesterdayKeepaRows.map((row) => [normalized(row.asin), row])),
    [yesterdayKeepaRows],
  )
  const mappingBySku = useMemo(() => new Map(mappingRows.map((row) => [normalized(row.sku), row])), [mappingRows])
  const combinedMonitorRows = useMemo(
    () => mergeMonitorRows(monitorRows, onlineRows),
    [monitorRows, onlineRows],
  )
  const resizingColumnRef = useRef<{ key: ResultColumnKey; startX: number; startWidth: number } | null>(null)

  useEffect(() => {
    if (yesterdayBuyBox.date && yesterdayKeepaRows.length) {
      setYesterdayBuyBox(buildBuyBoxBoard(
        combinedMonitorRows,
        mappingRows,
        yesterdayKeepaRows,
        [],
        yesterdayBuyBox.date,
        'snapshot',
      ))
    }
    if (todayBuyBox.date && keepaRows.length) {
      setTodayBuyBox(buildBuyBoxBoard(
        combinedMonitorRows,
        mappingRows,
        keepaRows,
        yesterdayKeepaRows,
        todayBuyBox.date,
        'changes',
      ))
    }
  }, [combinedMonitorRows, keepaRows, mappingRows, todayBuyBox.date, yesterdayBuyBox.date, yesterdayKeepaRows])

  const enrichedRows = useMemo(
    () =>
      combinedMonitorRows.map((row) => {
        const keepa = keepaByAsin.get(normalized(row.asin))
        const yesterdayKeepa = yesterdayKeepaByAsin.get(normalized(row.asin))
        const mapping = mappingBySku.get(normalized(row.sku))
        const ruleMatch = getRuleMatch(row, mapping, keepa)
        return {
          ...row,
          owner: ruleMatch.owner,
          group: ruleMatch.group,
          account: ruleMatch.account,
          ruleSource: ruleMatch.source,
          keepa,
          yesterdayKeepa,
          hasMapping: Boolean(mapping),
        }
      }),
    [combinedMonitorRows, keepaByAsin, mappingBySku, yesterdayKeepaByAsin],
  )

  const filteredRows = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const skuNeedle = normalized(skuQuery)
    const asinNeedle = normalized(asinQuery)
    const groupNeedle = normalized(groupQuery)
    const accountNeedle = normalized(accountQuery)
    const brandNeedle = normalized(brandQuery)
    const columnNeedles = Object.fromEntries(
      Object.entries(resultFilters).map(([key, value]) => [key, normalized(value)]),
    ) as Record<ResultFilterKey, string>

    return enrichedRows.filter((row) => {
      const price = typeof row.keepa?.price === 'number' ? String(row.keepa.price) : ''
      const todayRank = typeof row.keepa?.rank === 'number' ? String(row.keepa.rank) : ''
      const yesterdayRank = typeof row.yesterdayKeepa?.rank === 'number' ? String(row.yesterdayKeepa.rank) : ''
      const matchesColumns =
        (!columnNeedles.owner || normalized(row.owner).includes(columnNeedles.owner)) &&
        (!columnNeedles.sku || normalized(row.sku).includes(columnNeedles.sku)) &&
        (!columnNeedles.asinType || normalized(row.asinType).includes(columnNeedles.asinType)) &&
        (!columnNeedles.brand || normalized(row.keepa?.brand).includes(columnNeedles.brand)) &&
        (!columnNeedles.asin || normalized(row.asin).includes(columnNeedles.asin)) &&
        (!columnNeedles.price || normalized(price).includes(columnNeedles.price)) &&
        (!columnNeedles.todayRank || normalized(todayRank).includes(columnNeedles.todayRank)) &&
        (!columnNeedles.yesterdayRank || normalized(yesterdayRank).includes(columnNeedles.yesterdayRank))

      return matchesColumns &&
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!groupNeedle || normalized(row.group).includes(groupNeedle)) &&
        (!accountNeedle || normalized(row.account).includes(accountNeedle)) &&
        (!brandNeedle || normalized(row.keepa?.brand).includes(brandNeedle))
    })
  }, [accountQuery, asinQuery, brandQuery, enrichedRows, groupQuery, ownerQuery, resultFilters, skuQuery])

  const skuOptions = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const asinNeedle = normalized(asinQuery)
    const groupNeedle = normalized(groupQuery)
    const accountNeedle = normalized(accountQuery)
    const brandNeedle = normalized(brandQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!groupNeedle || normalized(row.group).includes(groupNeedle)) &&
        (!accountNeedle || normalized(row.account).includes(accountNeedle)) &&
        (!brandNeedle || normalized(row.keepa?.brand).includes(brandNeedle))
      if (matchesContext && row.sku) options.add(row.sku)
      if (matchesContext && row.bundleSku) options.add(row.bundleSku)
      if (options.size >= 120) break
    }
    return [...options].sort()
  }, [accountQuery, asinQuery, brandQuery, enrichedRows, groupQuery, ownerQuery])

  const ownerOptions = useMemo(() => {
    const skuNeedle = normalized(skuQuery)
    const asinNeedle = normalized(asinQuery)
    const groupNeedle = normalized(groupQuery)
    const accountNeedle = normalized(accountQuery)
    const brandNeedle = normalized(brandQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!groupNeedle || normalized(row.group).includes(groupNeedle)) &&
        (!accountNeedle || normalized(row.account).includes(accountNeedle)) &&
        (!brandNeedle || normalized(row.keepa?.brand).includes(brandNeedle))
      if (matchesContext && row.owner) options.add(row.owner)
      if (options.size >= 120) break
    }
    return [...options].sort()
  }, [accountQuery, asinQuery, brandQuery, enrichedRows, groupQuery, skuQuery])

  const asinOptions = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const skuNeedle = normalized(skuQuery)
    const groupNeedle = normalized(groupQuery)
    const accountNeedle = normalized(accountQuery)
    const brandNeedle = normalized(brandQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!groupNeedle || normalized(row.group).includes(groupNeedle)) &&
        (!accountNeedle || normalized(row.account).includes(accountNeedle)) &&
        (!brandNeedle || normalized(row.keepa?.brand).includes(brandNeedle))
      if (matchesContext && row.asin) options.add(row.asin)
      if (options.size >= 160) break
    }
    return [...options].sort()
  }, [accountQuery, brandQuery, enrichedRows, groupQuery, ownerQuery, skuQuery])

  const groupOptions = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const skuNeedle = normalized(skuQuery)
    const asinNeedle = normalized(asinQuery)
    const accountNeedle = normalized(accountQuery)
    const brandNeedle = normalized(brandQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!accountNeedle || normalized(row.account).includes(accountNeedle)) &&
        (!brandNeedle || normalized(row.keepa?.brand).includes(brandNeedle))
      if (matchesContext && row.group) options.add(row.group)
      if (options.size >= 120) break
    }
    return [...options].sort()
  }, [accountQuery, asinQuery, brandQuery, enrichedRows, ownerQuery, skuQuery])

  const accountOptions = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const skuNeedle = normalized(skuQuery)
    const asinNeedle = normalized(asinQuery)
    const groupNeedle = normalized(groupQuery)
    const brandNeedle = normalized(brandQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!groupNeedle || normalized(row.group).includes(groupNeedle)) &&
        (!brandNeedle || normalized(row.keepa?.brand).includes(brandNeedle))
      if (matchesContext && row.account) options.add(row.account)
      if (options.size >= 120) break
    }
    return [...options].sort()
  }, [asinQuery, brandQuery, enrichedRows, groupQuery, ownerQuery, skuQuery])

  const brandOptions = useMemo(() => {
    const ownerNeedle = normalized(ownerQuery)
    const skuNeedle = normalized(skuQuery)
    const asinNeedle = normalized(asinQuery)
    const groupNeedle = normalized(groupQuery)
    const accountNeedle = normalized(accountQuery)
    const options = new Set<string>()
    for (const row of enrichedRows) {
      const matchesContext =
        (!ownerNeedle || normalized(row.owner).includes(ownerNeedle)) &&
        (!skuNeedle || [row.sku, row.bundleSku].map(normalized).some((value) => value.includes(skuNeedle))) &&
        (!asinNeedle || normalized(row.asin).includes(asinNeedle)) &&
        (!groupNeedle || normalized(row.group).includes(groupNeedle)) &&
        (!accountNeedle || normalized(row.account).includes(accountNeedle))
      if (matchesContext && row.keepa?.brand) options.add(row.keepa.brand)
      if (options.size >= 120) break
    }
    return [...options].sort()
  }, [accountQuery, asinQuery, enrichedRows, groupQuery, ownerQuery, skuQuery])

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

  useEffect(() => {
    const handleMouseMove = (event: MouseEvent) => {
      const resizing = resizingColumnRef.current
      if (!resizing) return
      const delta = event.clientX - resizing.startX
      setResultColumnWidths((current) => ({
        ...current,
        [resizing.key]: Math.max(88, resizing.startWidth + delta),
      }))
    }
    const handleMouseUp = () => {
      resizingColumnRef.current = null
    }
    document.addEventListener('mousemove', handleMouseMove)
    document.addEventListener('mouseup', handleMouseUp)
    return () => {
      document.removeEventListener('mousemove', handleMouseMove)
      document.removeEventListener('mouseup', handleMouseUp)
    }
  }, [])

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
          group: row.group,
          account: row.account,
          category: row.asinType,
          monitorIndexes: [index],
        })
      }
      const missingKeepaReason = !row.keepa
        ? 'Keepa 未出现在本次上传'
        : row.keepa.price === null && row.keepa.rank === null
          ? 'Keepa 已读取该 ASIN，但价格与排名都缺失'
          : row.keepa.price === null
            ? 'Keepa 已读取该 ASIN，但价格缺失'
            : row.keepa.rank === null
              ? 'Keepa 已读取该 ASIN，但排名缺失'
              : ''
      if (missingKeepaReason) {
        missingKeepa.push({
          id: `missing-keepa-${index}`,
          message: `${row.asin} ${missingKeepaReason}`,
          asin: row.asin,
          sku: row.sku,
          owner: row.owner,
          group: row.group,
          account: row.account,
          category: row.asinType,
          monitorIndexes: [index],
          detail: missingKeepaReason,
        })
      }
      if (row.ruleSource === '待补充规则') {
        unresolvedRules.push({
          id: `rule-gap-${index}`,
          message: `${row.asin} 缺少运营归类规则，请补充监控清单或映射信息`,
          asin: row.asin,
          sku: row.sku,
          owner: row.owner,
          group: row.group,
          account: row.account,
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
        group: rows[0]?.group,
        account: rows[0]?.account,
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
            group: row?.group,
            account: row?.account,
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
            group: row?.group,
            account: row?.account,
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

  const actionAlertGroups = useMemo(
    () => alertGroups.filter((group) => !['buybox', 'missing-keepa', 'rule-gap', 'duplicate'].includes(group.key)),
    [alertGroups],
  )

  const alerts = useMemo(
    () => actionAlertGroups.flatMap((group) => group.items.map((item) => item.message)),
    [actionAlertGroups],
  )

  const updateStats = [
    { label: '监控 ASIN', value: new Set(combinedMonitorRows.map((row) => row.asin)).size },
    { label: '平台 SKU', value: new Set(combinedMonitorRows.map((row) => row.sku)).size },
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

  const unmatchedOnlineRows = useMemo(
    () => onlineRows.filter((row) => !mappingBySku.has(normalized(row.sku))),
    [mappingBySku, onlineRows],
  )

  const openPage = (page: WorkspacePage) => {
    setOpenPages((current) => (current.includes(page) ? current : [...current, page]))
    setActivePage(page)
    if (page === 'add-monitor') {
      setBatchOnlineRows(Array.from({ length: 8 }, () => ({ ...emptyOnlineMonitor })))
      setStatus('已打开添加监控ASIN，可在线新增或导出模板批量维护。')
    } else if (page === 'data-update') {
      setStatus('已打开数据更新页，可上传 Keepa、映射信息和 SKU / ASIN 监控清单。')
    } else if (page === 'buybox') {
      setStatus('已打开 Buy Box 丢失/恢复页。')
    } else {
      setStatus('已打开运营看板。')
    }
  }

  const closePage = (page: WorkspacePage) => {
    setOpenPages((current) => {
      const next = current.filter((item) => item !== page) as WorkspacePage[]
      const fallback: WorkspacePage[] = next.length ? next : ['dashboard']
      if (activePage === page) {
        setActivePage(fallback[fallback.length - 1])
      }
      return fallback
    })
  }

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

  const renderRankComparison = (
    current: number | null | undefined,
    previous: number | null | undefined,
  ) => {
    if (typeof current !== 'number') return '-'
    const change = typeof previous === 'number' ? current - previous : 0
    return (
      <span className="rank-comparison">
        <span>{current.toLocaleString()}</span>
        {change < 0 ? <span title="较昨日排名上升"><ArrowUp aria-label="排名上升" className="rank-direction rank-up" size={15} /></span> : null}
        {change > 0 ? <span title="较昨日排名下跌"><ArrowDown aria-label="排名下跌" className="rank-direction rank-down" size={15} /></span> : null}
      </span>
    )
  }

  const renderResultHeader = (key: ResultFilterKey, label: string) => {
    const active = Boolean(resultFilters[key])
    const open = openResultFilter === key
    return (
      <th style={{ width: `${resultColumnWidths[key]}px` }}>
        <div className="result-filter-wrap">
          <span>{label}</span>
          <button
            aria-expanded={open}
            aria-label={`筛选${label}`}
            className={`result-filter-button ${active ? 'active' : ''}`}
            onClick={() => setOpenResultFilter((current) => current === key ? null : key)}
            title={`筛选${label}`}
            type="button"
          >
            <ChevronDown size={14} />
          </button>
          {open ? (
            <div className="result-filter-menu">
              <input
                aria-label={`${label}筛选条件`}
                autoFocus
                onChange={(event) => setResultFilters((current) => ({ ...current, [key]: event.target.value }))}
                placeholder={`筛选${label}`}
                value={resultFilters[key]}
              />
              {active ? <button onClick={() => setResultFilters((current) => ({ ...current, [key]: '' }))} type="button">清除</button> : null}
            </div>
          ) : null}
          <button
            aria-label={`调整${label}列宽`}
            className="column-resize-handle"
            onMouseDown={(event) => {
              event.preventDefault()
              event.stopPropagation()
              resizingColumnRef.current = {
                key,
                startX: event.clientX,
                startWidth: resultColumnWidths[key],
              }
            }}
            type="button"
          />
        </div>
      </th>
    )
  }

  const importKeepaFile = async (
    file: File,
    target: 'today' | 'yesterday',
  ) => {
    setStatus(`正在解析 ${file.name} ...`)
    setKeepaUploadReports((current) => ({
      ...current,
      [target]: {
        status: 'processing',
        fileName: file.name,
        date: getImportDate(file.name),
        imported: 0,
        notes: ['正在读取并校验 ASIN、New: Current、Sales Rank: Current 字段。'],
        errors: [],
      },
    }))
    try {
      const rows = await readKeepaRows(file)
      const parsed = parseKeepa(rows)
      if (!parsed.length) throw new Error('Keepa 文件没有识别到有效 ASIN，请检查文件内容。')

      const date = getImportDate(file.name)
      const previousRows = target === 'today' ? yesterdayKeepaRows : []
      const nextBoard = buildBuyBoxBoard(
        combinedMonitorRows,
        mappingRows,
        parsed,
        previousRows,
        date,
        target === 'today' ? 'changes' : 'snapshot',
      )
      if (target === 'yesterday') {
        setYesterdayKeepaRows(toKeepaSnapshotRows(parsed))
        setYesterdayBuyBox(nextBoard)
        setTodayBuyBox(emptyBuyBoxBoard())
        setKeepaRows([])
        setHistory(mergeRecentHistory([], parsed, date))
      } else {
        setTodayBuyBox(nextBoard)
        setKeepaRows(parsed)
        setHistory((current) => mergeRecentHistory(current, parsed, date))
      }

      const missingPrice = parsed.filter((row) => row.price === null).length
      const missingRank = parsed.filter((row) => row.rank === null).length
      const missingNewCurrent = parsed.filter((row) => row.newCurrent === null).length
      const successReport: KeepaUploadReport = {
        status: 'success',
        fileName: file.name,
        date,
        imported: parsed.length,
        notes: [
          `识别到 ${parsed.length} 条 ASIN`,
          `Sales Rank: Subcategory Sales Ranks 有效 ${parsed.length - missingRank} 条，缺失 ${missingRank} 条`,
          `New: Current 为空 ${missingNewCurrent} 条，${target === 'today' ? '今日新增丢失' : '昨日丢失'} ${nextBoard.lost.length} 条`,
          `缺少价格 ${missingPrice} 条${target === 'today' ? `，今日恢复 ${nextBoard.recovered.length} 条` : ''}`,
        ],
        errors: [],
      }
      setKeepaUploadReports((current) => ({
        ...current,
        ...(target === 'yesterday' ? { today: emptyKeepaUploadReport() } : {}),
        [target]: successReport,
      }))
      setStatus(
        target === 'yesterday'
          ? `已将 ${parsed.length} 条 Keepa 数据保存为昨日基准：丢失 ${nextBoard.lost.length} 条。`
          : `已导入 ${parsed.length} 条 Keepa 数据：今日新增丢失 ${nextBoard.lost.length} 条，恢复 ${nextBoard.recovered.length} 条。`,
      )
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Keepa 文件解析失败，请检查导出字段。'
      setKeepaUploadReports((current) => ({
        ...current,
        [target]: {
          status: 'error',
          fileName: file.name,
          date: getImportDate(file.name),
          imported: 0,
          notes: ['本次上传未写入，原有数据未改变。'],
          errors: [message],
        },
      }))
      setStatus(message)
    }
  }

  const handleSidebarDownload = (kind: UploadKind) => {
    if (kind === 'keepa') {
      downloadTemplate('keepa')
      return
    }
    if (kind === 'mapping') {
      exportRows('mapping', mappingRows, '映射信息完整版.xlsx')
      return
    }
    exportMonitorTemplatePackage(combinedMonitorRows, mappingRows)
  }

  const handleUpload = async (kind: UploadKind, file: File | null) => {
    if (!file) return
    if (kind === 'keepa') {
      setPendingKeepaUpload({ file })
      setStatus('请选择将这份 Keepa 数据作为昨日基准或今日数据。')
      return
    }
    setStatus(`正在解析 ${file.name} ...`)
    try {
      const rows = await readWorkbookRows(file, kind)
      const updatedAt = new Date().toLocaleString('zh-CN', { hour12: false })
      if (kind === 'mapping') {
        const parsed = parseMapping(rows)
        if (!parsed.length) throw new Error('映射文件没有识别到有效的平台 SKU 行，现有映射信息未改变。')
        setMappingRows(parsed)
        setOnlineRows((current) => applyMappingsToOnlineRows(current, parsed))
        setSourceReports((current) => ({
          ...current,
          mapping: { fileName: file.name, imported: parsed.length, updatedAt, source: 'upload' },
        }))
        setUploadSummary({
          kind,
          fileName: file.name,
          imported: parsed.length,
          notes: [
            `固定映射源已替换为 ${parsed.length} 条记录`,
            '在线添加数据已自动重新匹配运营、组别、账号',
            '当前按“平台SKU -> 运营 / 小组 / 店铺别名”自动匹配在线新增记录',
          ],
          errors: [],
        })
        setStatus(`已更新固定映射源：${parsed.length} 条；在线数据已同步重新匹配。`)
      }
      if (kind === 'monitor') {
        const parsed = parseMonitor(rows)
        if (!parsed.length) throw new Error('监控清单没有识别到有效的 SKU 和 ASIN，现有固定监控源未改变。')
        setMonitorRows(parsed)
        setSelectedAsin(parsed[0]?.asin ?? onlineRows[0]?.asin ?? '')
        setSourceReports((current) => ({
          ...current,
          monitor: { fileName: file.name, imported: parsed.length, updatedAt, source: 'upload' },
        }))
        setUploadSummary({
          kind,
          fileName: file.name,
          imported: parsed.length,
          notes: [
            `固定监控源已替换为 ${parsed.length} 条关系`,
            `保留在线添加 ${onlineRows.length} 条，汇总时自动合并`,
          ],
          errors: [],
        })
        setStatus(`已更新固定监控源：${parsed.length} 条；在线添加 ${onlineRows.length} 条保持不变。`)
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : `${uploadLabels[kind]}解析失败，原有数据未改变。`
      setUploadSummary({ kind, fileName: file.name, imported: 0, notes: ['本次上传未写入，固定数据源保持不变。'], errors: [message] })
      setStatus(message)
    }
  }

  const continueKeepaUpload = async (target: 'today' | 'yesterday') => {
    const pending = pendingKeepaUpload
    if (!pending) return
    setPendingKeepaUpload(null)
    await importKeepaFile(pending.file, target)
  }

  const cancelKeepaUpload = () => {
    setPendingKeepaUpload(null)
    setStatus('已取消本次 Keepa 上传，当前今日和昨日数据未改变。')
  }

  const startOnlineEdit = (index: number) => {
    const target = onlineRows[index] ?? emptyOnlineMonitor
    setBatchOnlineRows([
      { ...target },
      ...Array.from({ length: 7 }, () => ({ ...emptyOnlineMonitor })),
    ])
  }

  const startBatchOnlineAdd = () => {
    setBatchOnlineRows(Array.from({ length: 8 }, () => ({ ...emptyOnlineMonitor })))
    setStatus('已切换到在线批量添加，可按表格方式录入多条监控。')
  }

  const updateBatchOnlineCell = (index: number, key: keyof MonitorRow, value: string) => {
    setBatchOnlineRows((rows) => rows.map((row, rowIndex) => {
      if (rowIndex !== index) return row
      if (key === 'sku') {
        const mapping = mappingBySku.get(normalized(value))
        return {
          ...row,
          sku: value,
          owner: mapping?.owner ?? row.owner,
          group: mapping?.group ?? row.group,
          account: mapping?.account ?? row.account,
        }
      }
      return { ...row, [key]: value }
    }))
  }

  const addBatchOnlineLine = () => {
    setBatchOnlineRows((rows) => [...rows, { ...emptyOnlineMonitor }])
  }

  const removeBatchOnlineLine = (index: number) => {
    setBatchOnlineRows((rows) => rows.length > 1 ? rows.filter((_, rowIndex) => rowIndex !== index) : [{ ...emptyOnlineMonitor }])
  }

  const batchColumnKeys: Array<keyof MonitorRow> = ['sku', 'asinType', 'asin', 'owner', 'group', 'account']

  const applyBatchValue = (rows: MonitorRow[], rowIndex: number, key: keyof MonitorRow, value: string) => {
    while (rows.length <= rowIndex) rows.push({ ...emptyOnlineMonitor })
    const next = rows[rowIndex]
    if (key === 'sku') {
      const mapping = mappingBySku.get(normalized(value))
      rows[rowIndex] = {
        ...next,
        sku: value,
        owner: mapping?.owner ?? next.owner,
        group: mapping?.group ?? next.group,
        account: mapping?.account ?? next.account,
      }
      return
    }
    rows[rowIndex] = { ...next, [key]: value }
  }

  const handleBatchPaste = (rowIndex: number, columnKey: keyof MonitorRow, text: string) => {
    if (!/[\n\t]/.test(text)) return false
    const startColumn = batchColumnKeys.indexOf(columnKey)
    if (startColumn < 0) return false
    const lines = text.replace(/\r/g, '').split('\n').filter((line) => line.length > 0)
    if (!lines.length) return false
    setBatchOnlineRows((current) => {
      const next = current.map((row) => ({ ...row }))
      lines.forEach((line, lineIndex) => {
        const cells = line.split('\t')
        cells.forEach((cell, cellIndex) => {
          const key = batchColumnKeys[startColumn + cellIndex]
          if (!key) return
          applyBatchValue(next, rowIndex + lineIndex, key, cell.trim())
        })
      })
      return next
    })
    return true
  }

  const saveBatchOnlineRows = () => {
    const prepared = batchOnlineRows
      .map((row) => {
        const mapping = mappingBySku.get(normalized(row.sku))
        return {
          ...row,
          sku: row.sku.trim(),
          asinType: row.asinType.trim(),
          asin: row.asin.trim(),
          owner: (mapping?.owner ?? row.owner).trim(),
          group: (mapping?.group ?? row.group).trim(),
          account: (mapping?.account ?? row.account).trim(),
          note: row.note || '在线批量添加',
        }
      })
      .filter((row) => row.sku || row.asin || row.asinType)

    if (!prepared.length) {
      setStatus('请先填写至少一条批量在线记录。')
      return
    }

    const invalid = prepared.find((row) => !row.sku || !row.asinType || !row.asin || !row.owner || !row.group || !row.account)
    if (invalid) {
      setStatus('批量在线添加中存在未补齐的记录，请检查平台 SKU、ASIN 分类、ASIN 及映射信息。')
      return
    }

    const seen = new Set<string>()
    for (const row of prepared) {
      const key = monitorRowKey(row)
      if (seen.has(key)) {
        setStatus('批量在线添加中存在重复的平台 SKU + ASIN 分类 + ASIN，请先处理后再保存。')
        return
      }
      seen.add(key)
    }

    setOnlineRows((rows) => {
      let next = [...rows]
      for (const row of prepared) {
        const duplicateIndex = next.findIndex((item) => monitorRowKey(item) === monitorRowKey(row))
        if (duplicateIndex >= 0) next[duplicateIndex] = row
        else next = [row, ...next]
      }
      return next
    })
    setSelectedAsin(prepared[0]?.asin ?? '')
    setBatchOnlineRows(Array.from({ length: 8 }, () => ({ ...emptyOnlineMonitor })))
    setStatus(`已批量写入 ${prepared.length} 条在线监控，并同步到监控汇总。`)
  }

  const deleteOnlineRow = (index: number) => {
    setOnlineRows((rows) => rows.filter((_, rowIndex) => rowIndex !== index))
    setStatus('已删除在线记录，固定监控源未改变。')
  }

  const refreshOnlineMappings = () => {
    setIsRefreshingOnline(true)
    window.setTimeout(() => {
      setOnlineRows((rows) => applyMappingsToOnlineRows(rows, mappingRows))
      setBatchOnlineRows((rows) => rows.map((row) => {
        const mapping = mappingBySku.get(normalized(row.sku))
        return mapping ? { ...row, owner: mapping.owner, group: mapping.group, account: mapping.account } : row
      }))
      setIsRefreshingOnline(false)
      setStatus(`已重新匹配 ${onlineRows.length} 条在线记录。`)
    }, 450)
  }

  return (
    <main className="app-shell">
      <aside className="sidebar erp-sidebar">
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

        <button className="reset-button" title="清空昨日和今日 Keepa 数据" type="button" onClick={clearKeepaData}><Trash2 size={16} />一键清空</button>
        <section className="nav-panel">
          <button className={maintenancePanel === 'online' ? 'active-nav' : ''} type="button" onClick={() => openMaintenancePanel('online')}><Plus size={15} />在线添加 <span>{onlineRows.length}</span></button>
          <button className={maintenancePanel === 'monitor' ? 'active-nav' : ''} type="button" onClick={() => openMaintenancePanel('monitor')}><Eye size={15} />SKU / ASIN 监控清单</button>
          <button className={maintenancePanel === 'mapping' ? 'active-nav' : ''} type="button" onClick={() => openMaintenancePanel('mapping')}><Eye size={15} />映射信息</button>
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
            {([
              ['yesterday', '昨日数据源'],
              ['today', '今日数据源'],
            ] as const).map(([key, label]) => {
              const report = keepaUploadReports[key]
              const statusLabel = report.status === 'success' ? '上传成功' : report.status === 'error' ? '上传失败' : report.status === 'processing' ? '解析中' : '等待上传'
              return (
                <section className={`sidebar-upload-role upload-role-${report.status}`} key={key}>
                  <div className="sidebar-upload-role-heading"><strong>{label}</strong><span>{statusLabel}</span></div>
                  <p>{report.fileName || '尚未选择文件'}</p>
                  {report.status === 'success' ? <p>{report.date} · 已解析 {report.imported.toLocaleString()} 条</p> : null}
                  {report.notes.length ? <ul className="report-list compact-report-list">{report.notes.map((note) => <li key={note}>{note}</li>)}</ul> : null}
                  {report.errors.length ? <ul className="report-list report-error compact-report-list">{report.errors.map((error) => <li key={error}>{error}</li>)}</ul> : null}
                </section>
              )
            })}
            {uploadSummary.fileName && uploadSummary.kind !== 'keepa' ? (
              <div className="sidebar-upload-block">
                <span>其他文件</span>
                <p>{uploadLabels[uploadSummary.kind]} · {uploadSummary.fileName} · {uploadSummary.imported.toLocaleString()} 条</p>
                {uploadSummary.errors.length ? <ul className="report-list report-error compact-report-list">{uploadSummary.errors.map((error) => <li key={error}>{error}</li>)}</ul> : null}
              </div>
            ) : null}
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
                <thead><tr>{renderResultHeader('owner', '运营')}{renderResultHeader('sku', 'SKU')}{renderResultHeader('asinType', '类型')}{renderResultHeader('brand', '品牌')}{renderResultHeader('asin', 'ASIN')}{renderResultHeader('price', '价格')}{renderResultHeader('todayRank', '今日排名')}{renderResultHeader('yesterdayRank', '昨日排名')}</tr></thead>
                <tbody>{visibleRows.map((row, index) => {
                  const previous = visibleRows[index - 1]
                  const next = visibleRows[index + 1]
                  const isGroupStart = !previous || previous.sku !== row.sku
                  const isGroupEnd = !next || next.sku !== row.sku
                  const typeClass = normalized(row.asinType).includes('kmasin') ? 'type-km' : normalized(row.asinType).includes('竞对') ? 'type-competitor' : 'type-neutral'
                  const asinHistory = historyByAsin.get(normalized(row.asin))
                  const priceChange = getMetricChange(asinHistory, 'price')
                  return <tr className={`${row.asin === selectedAsin ? 'selected-row' : ''} ${isGroupStart ? 'sku-group-start' : ''} ${isGroupEnd ? 'sku-group-end' : ''}`} key={`${row.sku}-${row.asin}-${index}`} onClick={() => setSelectedAsin(row.asin)}><td>{row.owner || '-'}</td><td className="sku-cell">{row.sku}</td><td><span className={`type-tag ${typeClass}`}>{row.asinType || '-'}</span></td><td>{row.keepa?.brand || '-'}</td><td className="asin-cell">{row.asin}</td><td>{renderMetric(row.keepa?.price, priceChange, 'price')}</td><td>{renderRankComparison(row.keepa?.rank, row.yesterdayKeepa?.rank)}</td><td>{typeof row.yesterdayKeepa?.rank === 'number' ? row.yesterdayKeepa.rank.toLocaleString() : '-'}</td></tr>
                })}</tbody>
              </table>
            </div>
          </div>

          <div className="detail-panel">
            <div className="section-heading"><h2>ASIN 详情</h2><span>{selectedAsin || '未选择'}</span></div>
            {selectedRows[0] ? <div className="detail-stack"><div><span className="eyebrow">商品</span><h3>{selectedRows[0].keepa?.title || selectedRows[0].asin}</h3></div>{getPrimaryImageUrl(selectedRows[0].keepa?.image) ? <div className="detail-image-wrap"><img alt={selectedRows[0].asin} className="detail-image" src={getPrimaryImageUrl(selectedRows[0].keepa?.image)} /></div> : null}<dl className="detail-list"><div><dt>运营</dt><dd>{selectedRows[0].owner || '-'}</dd></div><div><dt>SKU</dt><dd>{selectedRows[0].sku}</dd></div><div><dt>品牌</dt><dd>{selectedRows[0].keepa?.brand || '-'}</dd></div><div><dt>归类规则</dt><dd>{selectedRows[0].ruleSource}</dd></div><div><dt>New: Current</dt><dd>{typeof selectedRows[0].keepa?.newCurrent === 'number' ? selectedRows[0].keepa.newCurrent.toFixed(2) : '-'}</dd></div><div><dt>Coupon</dt><dd>{selectedRows[0].keepa?.coupon || '-'}</dd></div></dl></div> : <p className="empty-state">点击左侧结果查看 ASIN。</p>}
          </div>
        </section>

        {maintenancePanel === 'online' ? <section className="editor-panel" ref={editorPanelRef}>
          <div className="section-heading editor-heading">
            <div><span className="eyebrow">实时同步到监控汇总</span><h2>在线添加</h2></div>
            <div className="editor-actions">
              <button className="icon-command" title="重新匹配映射信息" type="button" onClick={refreshOnlineMappings}><RefreshCw className={isRefreshingOnline ? 'spin-icon' : ''} size={16} /></button>
              <button type="button" onClick={startOnlineAdd}><Plus size={16} />新增</button>
              <button type="button" onClick={saveOnlineRow}><Save size={16} />{editingIndex === null ? '添加' : '更新'}</button>
              <button type="button" onClick={() => exportRows('monitor', onlineRows, '在线添加SKU-ASIN.xlsx')}><Download size={16} />导出在线数据</button>
              <button className="danger-command" disabled={!onlineRows.length} type="button" onClick={clearOnlineRows}><Trash2 size={16} />一键清除</button>
            </div>
          </div>
          <div className="online-form">
            <label><span>平台 SKU</span><input placeholder="手动填写" value={onlineForm.sku} onChange={(event) => updateOnlineSku(event.target.value)} /></label>
            <label><span>ASIN 分类</span><input placeholder="例如 KMASIN" value={onlineForm.asinType} onChange={(event) => setOnlineForm({ ...onlineForm, asinType: event.target.value })} /></label>
            <label><span>ASIN</span><input placeholder="手动填写" value={onlineForm.asin} onChange={(event) => setOnlineForm({ ...onlineForm, asin: event.target.value })} /></label>
            <label><span>账号</span><input className={onlineFormMapping ? 'mapped-field' : ''} placeholder={onlineFormNeedsManualMapping ? '未匹配，请手动填写' : '映射表自动匹配'} readOnly={Boolean(onlineFormMapping)} value={onlineForm.account} onChange={(event) => setOnlineForm({ ...onlineForm, account: event.target.value })} /></label>
            <label><span>组别</span><input className={onlineFormMapping ? 'mapped-field' : ''} placeholder={onlineFormNeedsManualMapping ? '未匹配，请手动填写' : '映射表自动匹配'} readOnly={Boolean(onlineFormMapping)} value={onlineForm.group} onChange={(event) => setOnlineForm({ ...onlineForm, group: event.target.value })} /></label>
            <label><span>运营</span><input className={onlineFormMapping ? 'mapped-field' : ''} placeholder={onlineFormNeedsManualMapping ? '未匹配，请手动填写' : '映射表自动匹配'} readOnly={Boolean(onlineFormMapping)} value={onlineForm.owner} onChange={(event) => setOnlineForm({ ...onlineForm, owner: event.target.value })} /></label>
          </div>
          {onlineFormNeedsManualMapping ? <div className="mapping-status mapping-status-warning"><AlertTriangle size={17} /><span>未在映射表找到该平台 SKU，请手动填写账号、组别和运营。</span></div> : onlineFormMapping ? <div className="mapping-status mapping-status-success"><Database size={17} /><span>已匹配映射表：{onlineFormMapping.systemSku || onlineFormMapping.sku}</span></div> : null}
          <div className="mini-table-wrap online-table-wrap"><table className="data-table mini-table"><thead><tr><th>平台 SKU</th><th>ASIN 分类</th><th>ASIN</th><th>账号</th><th>组别</th><th>运营</th><th>匹配状态</th><th>操作</th></tr></thead><tbody>{onlineRows.length ? onlineRows.map((row, index) => {
            const matched = mappingBySku.has(normalized(row.sku))
            return <tr key={`${monitorRowKey(row)}-${index}`}><td className="sku-cell">{row.sku}</td><td>{row.asinType}</td><td className="asin-cell">{row.asin}</td><td>{row.account || '-'}</td><td>{row.group || '-'}</td><td>{row.owner || '-'}</td><td>{matched ? <span className="source-tag source-tag-mapped">已映射</span> : <span className="source-tag source-tag-manual"><AlertTriangle size={13} />手动填写</span>}</td><td><button className="row-icon" title="编辑" type="button" onClick={() => startOnlineEdit(index)}><Edit3 size={14} /></button><button className="row-icon danger" title="删除" type="button" onClick={() => deleteOnlineRow(index)}><Trash2 size={14} /></button></td></tr>
          }) : <tr><td className="empty-cell" colSpan={8}>暂无在线添加数据</td></tr>}</tbody></table></div>
        </section> : null}

        {maintenancePanel === 'monitor' || maintenancePanel === 'mapping' ? <section className="editor-panel source-preview-panel" ref={editorPanelRef}>
          <div className="section-heading editor-heading">
            <div><span className="eyebrow">固定数据源 · 仅重新上传时更新</span><h2>{maintenancePanel === 'monitor' ? 'SKU / ASIN 监控清单' : '映射信息'}</h2></div>
            <div className="editor-actions">
              <button type="button" onClick={() => maintenancePanel === 'monitor' ? exportRows('monitor', combinedMonitorRows, 'SKU-ASIN监控清单汇总.xlsx') : exportRows('mapping', mappingRows)}><Download size={16} />{maintenancePanel === 'monitor' ? '导出最新汇总' : '导出映射信息'}</button>
            </div>
          </div>
          <div className="source-meta"><FileSpreadsheet size={17} /><span>{sourceReports[maintenancePanel].fileName}</span><span>{sourceReports[maintenancePanel].updatedAt}</span><strong>{sourceReports[maintenancePanel].imported.toLocaleString()} 条</strong></div>
          {maintenancePanel === 'monitor' ? <>
            <div className="source-summary-grid"><div><span>固定源</span><strong>{monitorRows.length.toLocaleString()}</strong></div><div><span>在线添加</span><strong>{onlineRows.length.toLocaleString()}</strong></div><div><span>最新汇总</span><strong>{combinedMonitorRows.length.toLocaleString()}</strong></div><div><span>监控 ASIN</span><strong>{new Set(combinedMonitorRows.map((row) => normalized(row.asin))).size.toLocaleString()}</strong></div></div>
            <div className="mini-table-wrap source-table-wrap"><table className="data-table mini-table"><thead><tr><th>运营</th><th>组别</th><th>账号</th><th>平台 SKU</th><th>ASIN 分类</th><th>ASIN</th><th>数据层</th></tr></thead><tbody>{monitorRows.slice(0, 300).map((row, index) => <tr key={`${monitorRowKey(row)}-${index}`}><td>{row.owner || '-'}</td><td>{row.group || '-'}</td><td>{row.account || '-'}</td><td className="sku-cell">{row.sku}</td><td>{row.asinType || '-'}</td><td className="asin-cell">{row.asin}</td><td><span className="source-tag">固定源</span></td></tr>)}</tbody></table></div>
            {monitorRows.length > 300 ? <p className="preview-note">页面展示前 300 条，导出包含全部 {combinedMonitorRows.length.toLocaleString()} 条最新汇总数据。</p> : null}
          </> : <>
            <div className="source-summary-grid"><div><span>映射记录</span><strong>{mappingRows.length.toLocaleString()}</strong></div><div><span>平台 SKU</span><strong>{new Set(mappingRows.map((row) => normalized(row.sku))).size.toLocaleString()}</strong></div><div><span>在线已匹配</span><strong>{onlineRows.filter((row) => mappingBySku.has(normalized(row.sku))).length.toLocaleString()}</strong></div><div><span>在线待补映射</span><strong>{onlineRows.filter((row) => !mappingBySku.has(normalized(row.sku))).length.toLocaleString()}</strong></div></div>
            <div className="mini-table-wrap source-table-wrap"><table className="data-table mini-table"><thead><tr><th>平台 SKU</th><th>系统 SKU</th><th>运营</th><th>小组</th><th>店铺别名</th></tr></thead><tbody>{mappingRows.slice(0, 300).map((row, index) => <tr key={`${normalized(row.sku)}-${index}`}><td className="sku-cell">{row.sku}</td><td>{row.systemSku || '-'}</td><td>{row.owner || '-'}</td><td>{row.group || '-'}</td><td>{row.account || '-'}</td></tr>)}</tbody></table></div>
            {mappingRows.length > 300 ? <p className="preview-note">页面展示前 300 条，导出包含全部 {mappingRows.length.toLocaleString()} 条映射数据。</p> : null}
          </>}
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
                  <button className="action-pill" type="button" onClick={() => exportMonitorWithDuplicateMarks(combinedMonitorRows)}>下载删减后重传表</button>
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
            <div className="dialog-heading"><Save size={20} /><div><span className="eyebrow">上传每日 Keepa 数据</span><h2 id="keepa-save-title">这份数据作为哪一天？</h2></div></div>
            <p id="keepa-save-description">
              请选择 <strong>{pendingKeepaUpload.file.name}</strong> 的日期角色。作为昨日基准会替换右侧昨日数据并清空当前今日结果；作为今日数据会{yesterdayBuyBox.date ? <>与 <strong>{yesterdayBuyBox.date}</strong> 的昨日基准比较</> : '因没有昨日基准而无法判断今日丢失或恢复'}。
            </p>
            <div className="dialog-actions">
              <button className="dialog-button dialog-button-cancel" type="button" onClick={cancelKeepaUpload}>取消上传</button>
              <button className="dialog-button" type="button" onClick={() => void continueKeepaUpload('yesterday')}><Save size={16} />作为昨日基准</button>
              <button autoFocus className="dialog-button dialog-button-primary" type="button" onClick={() => void continueKeepaUpload('today')}><Upload size={16} />作为今日数据</button>
            </div>
          </section>
        </div>
      ) : null}
    </main>
  )
}

export default App
