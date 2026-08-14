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
  Eye,
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
import {
  getSharedStateConfig,
  isSharedStateConfigured,
  readSharedState,
  writeSharedState,
} from './lib/sharedState'
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

type MappingIndex = {
  bySku: Map<string, MappingRow[]>
  bySystemSkuCandidate: Map<string, MappingRow[]>
  rowOrder: Map<MappingRow, number>
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

type KeepaSnapshotRow = Pick<KeepaRow, 'asin' | 'price' | 'newCurrent' | 'rank'>

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

type ResultFilterKey = 'owner' | 'sku' | 'asinType' | 'brand' | 'asin' | 'price' | 'todayRank' | 'yesterdayRank' | 'stockStatus'

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

type DuplicateActionRow = {
  sku: string
  asin: string
  action: string
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

type SourceUploadReport = {
  status: 'idle' | 'processing' | 'success' | 'error'
  fileName: string
  imported: number
  notes: string[]
  errors: string[]
}

type SourceUploadReports = Record<EditMode, SourceUploadReport>

type DataPreviewMode = EditMode | null

type MonitorSyncReport = {
  updatedAt: string
  added: number
  deleted: number
  total: number
  action: 'idle' | 'add' | 'delete' | 'import' | 'upload' | 'duplicate-cleanup' | 'migration'
}

type MonitorDailyAdditions = {
  date: string
  keys: string[]
}

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
  monitorSyncReport: MonitorSyncReport
  monitorDailyAdditions: MonitorDailyAdditions
}

const typedSeed = seedData as unknown as {
  keepaRows: KeepaRow[]
  mappingRows: MappingRow[]
  monitorRows: MonitorRow[]
}

const storageKey = 'buybox-monitor-state-v2'
const mappingStorageKey = 'buybox-monitor-mapping-v1'

const uploadLabels: Record<UploadKind, string> = {
  keepa: '每日 Keepa 数据源',
  mapping: '映射信息',
  monitor: 'SKU / ASIN 监控清单',
}

const monitorSyncActionLabels: Record<MonitorSyncReport['action'], string> = {
  idle: '等待回传',
  add: '在线添加',
  delete: '在线删除',
  import: '表格导入',
  upload: '清单上传',
  'duplicate-cleanup': '重复清理',
  migration: '旧数据合并',
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
const mappingRequiredHeaders = ['平台SKU', '系统SKU', '运营', '小组', '店铺别名']
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

const downloadArrayBufferFile = (buffer: ArrayBuffer, fileName: string) => {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

const readWorkbookRows = async (file: File, kind: EditMode | 'monitor'): Promise<AnyRow[]> => {
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
      const matched = group.some((candidate) => cells.some((cell) => cell === normalized(candidate)))
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

const mappingCompositeKey = (row: Pick<MappingRow, 'sku' | 'account'>) =>
  `${normalized(row.sku)}::${normalized(row.account)}`

const mappingAssignmentKey = (row: MappingRow) =>
  `${normalized(row.sku)}::${normalized(row.systemSku)}::${normalized(row.owner)}::${normalized(row.group)}::${normalized(row.account)}`

const uniqueMappingRows = (rows: MappingRow[]) => {
  const seen = new Set<string>()
  return rows.filter((row) => {
    const key = mappingAssignmentKey(row)
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

const analyzeMappingRows = (rows: AnyRow[], parsedRows: MappingRow[]) => {
  const missingHeaders = mappingRequiredHeaders.filter((header) =>
    !Object.keys(rows[0] ?? {}).some((key) => normalized(key) === normalized(header)),
  )
  const dataRows = rows.filter((row) => Object.values(row).some((value) => String(value ?? '').trim()))
  const invalidRows = dataRows.flatMap((row, index) => {
    const parsed = parseMapping([row])[0]
    const missing = [
      ['平台SKU', parsed?.sku],
      ['系统SKU', parsed?.systemSku],
      ['运营', parsed?.owner],
      ['小组', parsed?.group],
      ['店铺别名', parsed?.account],
    ].filter(([, value]) => !value).map(([label]) => label)
    return missing.length ? [{ row: index + 1, missing }] : []
  })
  const compositeGroups = new Map<string, MappingRow[]>()
  for (const row of parsedRows) {
    const key = mappingCompositeKey(row)
    compositeGroups.set(key, [...(compositeGroups.get(key) ?? []), row])
  }
  const compositeConflicts = [...compositeGroups.entries()].filter(([, group]) =>
    new Set(group.map(mappingAssignmentKey)).size > 1,
  )
  const skuGroups = new Map<string, MappingRow[]>()
  for (const row of parsedRows) {
    const key = normalized(row.sku)
    skuGroups.set(key, [...(skuGroups.get(key) ?? []), row])
  }
  const multiAccountSkus = [...skuGroups.values()].filter((group) =>
    new Set(group.map((row) => normalized(row.account))).size > 1,
  ).length
  return { missingHeaders, invalidRows, compositeConflicts, multiAccountSkus }
}

const splitMultiValueCell = (value: string) =>
  value
    .split(/[\r\n/;,，；]+/)
    .map((item) => item.trim())
    .filter(Boolean)

const expandMonitorRows = (rows: MonitorRow[]) =>
  rows.flatMap((row) => {
    const skuParts = splitMultiValueCell(row.sku)
    const bundleParts = splitMultiValueCell(row.bundleSku)
    const targetSkus = skuParts.length ? skuParts : [row.sku.trim()]

    return targetSkus
      .map((sku, index) => ({
        ...row,
        sku,
        bundleSku: bundleParts[index] ?? bundleParts[0] ?? row.bundleSku,
      }))
      .filter((item) => item.sku && item.asin)
  })

const parseMonitor = (rows: AnyRow[]): MonitorRow[] =>
  expandMonitorRows(
    rows.map((row) => ({
      owner: String(pick(row, ['运营', 'Owner']) || '').trim(),
      group: String(pick(row, ['组别', '小组', 'Group']) || '').trim(),
      account: String(pick(row, ['账号', '店铺别名', 'Account']) || '').trim(),
      sku: String(pick(row, ['平台sku', '平台SKU', 'SKU']) || '').trim(),
      bundleSku: String(pick(row, ['Bundle主SKU', 'Bundle 主 SKU']) || '').trim(),
      asinType: String(pick(row, ['ASIN分类', 'ASIN 分类', '类型']) || '').trim(),
      asin: String(pick(row, ['ASIN', 'asin']) || '').trim(),
      note: String(pick(row, ['竞对备注列', '备注', '库存状态备注']) || '').trim(),
    })),
  )

const monitorRowKey = (row: Pick<MonitorRow, 'sku' | 'asinType' | 'asin'>) =>
  `${normalized(row.sku)}::${normalized(row.asinType)}::${normalized(row.asin)}`

const mergeMonitorRows = (sourceRows: MonitorRow[], onlineRows: MonitorRow[]) => {
  const sourceKeys = new Set(sourceRows.map(monitorRowKey))
  return [...sourceRows, ...onlineRows.filter((row) => !sourceKeys.has(monitorRowKey(row)))]
}

const upsertMonitorRows = (currentRows: MonitorRow[], incomingRows: MonitorRow[]) => {
  const next = [...currentRows]
  const addedRows: MonitorRow[] = []
  for (const row of incomingRows) {
    const duplicateIndex = next.findIndex((item) => monitorRowKey(item) === monitorRowKey(row))
    if (duplicateIndex >= 0) next[duplicateIndex] = row
    else {
      next.unshift(row)
      addedRows.push(row)
    }
  }
  return { rows: next, added: addedRows.length, addedRows }
}

const getSystemSkuCandidatesFromPlatformSku = (sku: string) => {
  const raw = sku.trim()
  if (!raw) return []
  const candidates = new Set<string>([raw])
  const prefixStripped = [
    raw.replace(/^KM1/i, ''),
    raw.replace(/^KM2/i, ''),
    raw.replace(/^KM/i, ''),
    raw.replace(/^UTV1/i, ''),
    raw.replace(/^UTV/i, ''),
  ]
  const suffixStripped = [
    raw.replace(/-FBM$/i, ''),
    raw.replace(/-FBA$/i, ''),
    raw.replace(/-RE$/i, ''),
  ]
  ;[...prefixStripped, ...suffixStripped].map((item) => item.trim()).filter(Boolean).forEach((item) => candidates.add(item))
  return [...candidates]
}

const appendMappingIndexEntry = (index: Map<string, MappingRow[]>, key: string, row: MappingRow) => {
  if (!key) return
  const matches = index.get(key)
  if (matches) matches.push(row)
  else index.set(key, [row])
}

const createMappingIndex = (mappings: MappingRow[]): MappingIndex => {
  const bySku = new Map<string, MappingRow[]>()
  const bySystemSkuCandidate = new Map<string, MappingRow[]>()
  const rowOrder = new Map<MappingRow, number>()

  mappings.forEach((mapping, index) => {
    rowOrder.set(mapping, index)
    appendMappingIndexEntry(bySku, normalized(mapping.sku), mapping)
    const systemCandidates = new Set(
      getSystemSkuCandidatesFromPlatformSku(mapping.systemSku).map(normalized).filter(Boolean),
    )
    systemCandidates.forEach((candidate) => appendMappingIndexEntry(bySystemSkuCandidate, candidate, mapping))
  })

  return { bySku, bySystemSkuCandidate, rowOrder }
}

const getIndexedSystemMatches = (sku: string, mappingIndex: MappingIndex) => {
  const matches = new Set<MappingRow>()
  const candidates = new Set(getSystemSkuCandidatesFromPlatformSku(sku).map(normalized).filter(Boolean))
  candidates.forEach((candidate) => {
    mappingIndex.bySystemSkuCandidate.get(candidate)?.forEach((mapping) => matches.add(mapping))
  })
  return [...matches].sort(
    (left, right) => (mappingIndex.rowOrder.get(left) ?? 0) - (mappingIndex.rowOrder.get(right) ?? 0),
  )
}

const resolveDirectMapping = (row: Pick<MonitorRow, 'sku' | 'account'>, mappingIndex: MappingIndex) => {
  const sku = normalized(row.sku)
  const account = normalized(row.account)
  const exactSku = mappingIndex.bySku.get(sku) ?? []
  const exactAccount = account ? exactSku.filter((mapping) => normalized(mapping.account) === account) : []
  if (exactAccount.length === 1) return { mapping: exactAccount[0], status: 'exact' as const, candidates: exactAccount }
  if (exactSku.length === 1) return { mapping: exactSku[0], status: 'unique' as const, candidates: exactSku }
  if (exactSku.length > 1) return { mapping: null, status: 'ambiguous' as const, candidates: exactSku }

  const systemMatches = getIndexedSystemMatches(row.sku, mappingIndex)
  const systemAccount = account ? systemMatches.filter((mapping) => normalized(mapping.account) === account) : []
  if (systemAccount.length === 1) return { mapping: systemAccount[0], status: 'system-exact' as const, candidates: systemAccount }
  if (systemMatches.length === 1) return { mapping: systemMatches[0], status: 'system-unique' as const, candidates: systemMatches }
  if (systemMatches.length > 1) return { mapping: null, status: 'ambiguous' as const, candidates: systemMatches }
  return { mapping: null, status: 'missing' as const, candidates: [] }
}

const getCompositePrimarySku = (sku: string) => {
  const parts = sku.split(/[+＋]/).map((part) => part.trim()).filter(Boolean)
  return parts.length > 1 ? parts[0] : ''
}

const resolveMapping = (row: Pick<MonitorRow, 'sku' | 'account'>, mappingIndex: MappingIndex) => {
  const direct = resolveDirectMapping(row, mappingIndex)
  if (direct.mapping || direct.status === 'ambiguous') {
    return { ...direct, matchedSku: direct.mapping ? row.sku.trim() : '' }
  }

  const primarySku = getCompositePrimarySku(row.sku)
  if (!primarySku) return { ...direct, matchedSku: '' }
  const primary = resolveDirectMapping({ sku: primarySku, account: row.account }, mappingIndex)
  if (primary.mapping) return { ...primary, status: 'composite-primary' as const, matchedSku: primarySku }
  return { ...primary, matchedSku: primary.status === 'ambiguous' ? primarySku : '' }
}

const applyMappingsToMonitorRows = (rows: MonitorRow[], mappingIndex: MappingIndex) => {
  return rows.map((row) => {
    const mapping = resolveMapping(row, mappingIndex).mapping
    return mapping
      ? {
          ...row,
          owner: row.owner || mapping.owner,
          group: row.group || mapping.group,
          account: row.account || mapping.account,
        }
      : row
  })
}

const applyMappingsToOnlineRows = (rows: MonitorRow[], mappingIndex: MappingIndex) => {
  return applyMappingsToMonitorRows(rows, mappingIndex)
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
    price: typeof row.price === 'number' && Number.isFinite(row.price) ? row.price : null,
    newCurrent: typeof row.newCurrent === 'number' && Number.isFinite(row.newCurrent) ? row.newCurrent : null,
    rank: typeof row.rank === 'number' && Number.isFinite(row.rank) ? row.rank : null,
  })).filter((row) => row.asin)

const compactKeepaRowsForStorage = (rows: KeepaRow[]): KeepaRow[] =>
  rows.map((row) => ({ ...row }))

const formatSyncTime = (value: string) => {
  if (!value) return '尚未同步'
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return value
  return new Intl.DateTimeFormat('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  }).format(date)
}

const emptyMonitorSyncReport = (total = 0): MonitorSyncReport => ({
  updatedAt: '',
  added: 0,
  deleted: 0,
  total,
  action: 'idle',
})

const emptyMonitorDailyAdditions = (): MonitorDailyAdditions => ({
  date: new Date().toLocaleDateString('sv-SE'),
  keys: [],
})

const normalizeMonitorSyncReport = (
  report: Partial<MonitorSyncReport> | null | undefined,
  total: number,
): MonitorSyncReport => ({
  updatedAt: String(report?.updatedAt ?? ''),
  added: Number.isFinite(report?.added) ? Number(report?.added) : 0,
  deleted: Number.isFinite(report?.deleted) ? Number(report?.deleted) : 0,
  total,
  action: report?.action ?? 'idle',
})

const deserializeStoredState = (raw: Partial<StoredState> | null | undefined): StoredState | null => {
  if (!raw) return null
  const seedKeepaRows = normalizeKeepaRows(typedSeed.keepaRows)
  const seedMonitorRows = expandMonitorRows(typedSeed.monitorRows)
  const fallback: StoredState = {
    keepaRows: seedKeepaRows,
    yesterdayKeepaRows: [],
    mappingRows: typedSeed.mappingRows,
    monitorRows: seedMonitorRows,
    onlineRows: [],
    sourceReports: {
      monitor: { fileName: 'seedData.json', imported: seedMonitorRows.length, updatedAt: '内置数据', source: 'seed' },
      mapping: { fileName: 'seedData.json', imported: typedSeed.mappingRows.length, updatedAt: '内置数据', source: 'seed' },
    },
    history: initialHistory(seedKeepaRows),
    todayBuyBox: emptyBuyBoxBoard(),
    yesterdayBuyBox: emptyBuyBoxBoard(),
    keepaUploadReports: emptyKeepaUploadReports(),
    monitorSyncReport: emptyMonitorSyncReport(seedMonitorRows.length),
    monitorDailyAdditions: emptyMonitorDailyAdditions(),
  }

  const storedKeepaRows = normalizeKeepaRows(raw.keepaRows)
  const storedYesterdayRows = raw.yesterdayKeepaRows
    ? toKeepaSnapshotRows(raw.yesterdayKeepaRows)
    : raw.yesterdayBuyBox?.date && !raw.todayBuyBox?.date
      ? toKeepaSnapshotRows(raw.keepaRows)
      : []
  const mappingRows = raw.mappingRows ?? fallback.mappingRows
  const storedMonitorRows = expandMonitorRows(raw.monitorRows ?? fallback.monitorRows)
  const legacyOnlineRows = expandMonitorRows(raw.onlineRows ?? fallback.onlineRows)
  const monitorRows = mergeMonitorRows(storedMonitorRows, legacyOnlineRows)
  const mappingIndex = createMappingIndex(mappingRows)
  const migratedCount = monitorRows.length - storedMonitorRows.length
  const todayDate = String(raw.todayBuyBox?.date ?? '')
  const yesterdayDate = String(raw.yesterdayBuyBox?.date ?? '')
  return {
    keepaRows: storedKeepaRows,
    yesterdayKeepaRows: storedYesterdayRows,
    mappingRows,
    monitorRows,
    onlineRows: [],
    sourceReports: raw.sourceReports ?? fallback.sourceReports,
    history: raw.history?.length ? keepRecentFiveDays(raw.history) : initialHistory(storedKeepaRows),
    todayBuyBox: todayDate
      ? buildBuyBoxBoard(monitorRows, mappingIndex, storedKeepaRows, storedYesterdayRows, todayDate, 'changes')
      : emptyBuyBoxBoard(),
    yesterdayBuyBox: yesterdayDate
      ? buildBuyBoxBoard(monitorRows, mappingIndex, storedYesterdayRows, [], yesterdayDate, 'snapshot')
      : emptyBuyBoxBoard(),
    keepaUploadReports: {
      yesterday: normalizeKeepaUploadReport(raw.keepaUploadReports?.yesterday),
      today: normalizeKeepaUploadReport(raw.keepaUploadReports?.today),
    },
    monitorSyncReport: migratedCount
      ? {
          updatedAt: new Date().toISOString(),
          added: migratedCount,
          deleted: 0,
          total: monitorRows.length,
          action: 'migration',
        }
      : normalizeMonitorSyncReport(raw.monitorSyncReport, monitorRows.length),
    monitorDailyAdditions: {
      date: String(raw.monitorDailyAdditions?.date ?? ''),
      keys: Array.isArray(raw.monitorDailyAdditions?.keys)
        ? [...new Set(raw.monitorDailyAdditions.keys.map(String).filter(Boolean))]
        : [],
    },
  }
}

const getImportDate = (fileName: string) => {
  const fileDate = fileName.match(/\d{4}-\d{2}-\d{2}/)?.[0]
  return fileDate ?? new Date().toLocaleDateString('sv-SE')
}

const buildBuyBoxBoard = (
  monitorRows: MonitorRow[],
  mappingIndex: MappingIndex,
  currentRows: Array<KeepaRow | KeepaSnapshotRow>,
  previousRows: Array<KeepaRow | KeepaSnapshotRow>,
  date: string,
  mode: 'snapshot' | 'changes',
): BuyBoxDayBoard => {
  const currentByAsin = new Map(currentRows.map((row) => [normalized(row.asin), row.newCurrent]))
  const previousByAsin = new Map(previousRows.map((row) => [normalized(row.asin), row.newCurrent]))
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
      owner: row.owner || resolveMapping(row, mappingIndex).mapping?.owner || '',
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

const getDirectMetricChange = (
  current: number | null | undefined,
  previous: number | null | undefined,
  metric: 'price' | 'rank',
): MetricChange | null => {
  if (typeof current !== 'number' || typeof previous !== 'number' || current === previous) return null
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
  if (row.owner || row.group || row.account || mapping) {
    return {
      owner: row.owner || mapping?.owner || '',
      group: row.group || mapping?.group || '',
      account: row.account || mapping?.account || '',
      source: row.owner && row.group && row.account ? '监控清单精确值' : '监控清单 + 映射补空值',
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
  const seedMonitorRows = expandMonitorRows(typedSeed.monitorRows)
  const fallback: StoredState = {
    keepaRows: seedKeepaRows,
    yesterdayKeepaRows: [],
    mappingRows: typedSeed.mappingRows,
    monitorRows: seedMonitorRows,
    onlineRows: [],
    sourceReports: {
      monitor: { fileName: 'seedData.json', imported: seedMonitorRows.length, updatedAt: '内置数据', source: 'seed' },
      mapping: { fileName: 'seedData.json', imported: typedSeed.mappingRows.length, updatedAt: '内置数据', source: 'seed' },
    },
    history: initialHistory(seedKeepaRows),
    todayBuyBox: emptyBuyBoxBoard(),
    yesterdayBuyBox: emptyBuyBoxBoard(),
    keepaUploadReports: emptyKeepaUploadReports(),
    monitorSyncReport: emptyMonitorSyncReport(seedMonitorRows.length),
    monitorDailyAdditions: emptyMonitorDailyAdditions(),
  }
  try {
    const raw = localStorage.getItem(storageKey)
    const storedState = raw
      ? deserializeStoredState(JSON.parse(raw) as Partial<StoredState>) ?? fallback
      : fallback
    const rawMapping = localStorage.getItem(mappingStorageKey)
    if (!rawMapping) return storedState
    const mappingSnapshot = JSON.parse(rawMapping) as { rows?: MappingRow[]; report?: SourceReport }
    if (!Array.isArray(mappingSnapshot.rows) || !mappingSnapshot.rows.length) return storedState
    return {
      ...storedState,
      mappingRows: mappingSnapshot.rows,
      sourceReports: {
        ...storedState.sourceReports,
        ...(mappingSnapshot.report ? { mapping: mappingSnapshot.report } : {}),
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

const emptySourceUploadReport = (): SourceUploadReport => ({
  status: 'idle',
  fileName: '',
  imported: 0,
  notes: [],
  errors: [],
})

const emptySourceUploadReports = (): SourceUploadReports => ({
  mapping: emptySourceUploadReport(),
  monitor: emptySourceUploadReport(),
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
  const guidance = kind === 'mapping'
    ? ['填写说明：平台SKU + 店铺别名共同构成唯一匹配键；五个字段均为必填。上传会在完整校验成功后一次性覆盖旧映射，失败时旧数据保持不变。']
    : []
  const sheet = XLSX.utils.aoa_to_sheet([...(guidance.length ? [guidance] : []), headers, example])
  sheet['!cols'] = kind === 'mapping'
    ? [{ wch: 26 }, { wch: 24 }, { wch: 22 }, { wch: 24 }, { wch: 20 }]
    : undefined
  sheet['!autofilter'] = { ref: guidance.length ? `A2:${XLSX.utils.encode_col(headers.length - 1)}2` : `A1:${XLSX.utils.encode_col(headers.length - 1)}1` }
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

const normalizeAction = (value: string) => normalized(value).replace(/\s+/g, '')

const parseDuplicateActionRows = (rows: AnyRow[]): DuplicateActionRow[] =>
  rows
    .map((row) => ({
      sku: String(pick(row, ['平台SKU', 'SKU', 'sku']) || '').trim(),
      asin: String(pick(row, ['ASIN', 'asin']) || '').trim(),
      action: String(pick(row, ['处理动作', '操作', '处理结果']) || '').trim(),
    }))
    .filter((row) => row.sku && row.asin && row.action)

const applyDuplicateActionsAcrossSources = (
  monitorRows: MonitorRow[],
  onlineRows: MonitorRow[],
  actionRows: DuplicateActionRow[],
) => {
  const groupedActions = new Map<string, DuplicateActionRow[]>()
  actionRows.forEach((row) => {
    const key = `${normalized(row.sku)}::${normalized(row.asin)}`
    groupedActions.set(key, [...(groupedActions.get(key) ?? []), row])
  })

  const removeKeys = new Set<string>()
  const keepKeys = new Set<string>()
  for (const [key, rows] of groupedActions.entries()) {
    const actions = rows.map((row) => normalizeAction(row.action))
    if (actions.some((action) => action.includes('剔除'))) {
      removeKeys.add(key)
      continue
    }
    if (actions.some((action) => action.includes('保留'))) keepKeys.add(key)
  }

  const seenDuplicateKeys = new Set<string>()
  const nextMonitorRows: MonitorRow[] = []
  const nextOnlineRows: MonitorRow[] = []
  const removedKeys = new Set<string>()
  const keptActionKeys = new Set<string>()
  const orderedRows = [
    ...monitorRows.map((row) => ({ row, source: 'monitor' as const })),
    ...onlineRows.map((row) => ({ row, source: 'online' as const })),
  ]

  for (const item of orderedRows) {
    const row = item.row
    const duplicateKey = `${normalized(row.sku)}::${normalized(row.asin)}`
    const isTargetedForRemoval = removeKeys.has(duplicateKey)
    const alreadySeen = seenDuplicateKeys.has(duplicateKey)

    if (!isTargetedForRemoval || !alreadySeen) {
      if (isTargetedForRemoval) seenDuplicateKeys.add(duplicateKey)
      if (item.source === 'monitor') nextMonitorRows.push(row)
      else nextOnlineRows.push(row)
      if (keepKeys.has(duplicateKey)) keptActionKeys.add(duplicateKey)
      continue
    }

    removedKeys.add(duplicateKey)
  }

  return {
    monitorRows: nextMonitorRows,
    onlineRows: nextOnlineRows,
    removedKeys: [...removedKeys],
    keptKeys: [...keptActionKeys],
    touchedCount: actionRows.length,
  }
}

const exportMonitorWithDuplicateMarks = async (rows: MonitorRow[]) => {
  const ExcelJS = await import('exceljs')
  const grouped = new Map<string, number[]>()
  rows.forEach((row, index) => {
    const key = `${normalized(row.sku)}::${normalized(row.asin)}`
    grouped.set(key, [...(grouped.get(key) ?? []), index])
  })
  const duplicateIndexes = new Set<number>()
  for (const indexes of grouped.values()) {
    if (indexes.length > 1) indexes.forEach((index) => duplicateIndexes.add(index))
  }
  const markedData = rows
    .map((row, index) => ({
      重复标记: duplicateIndexes.has(index) ? '重复' : '',
      处理动作: '',
      重复组键: `${row.sku} / ${row.asin}`,
      重复组标准键: `${normalized(row.sku)}::${normalized(row.asin)}`,
      运营: row.owner,
      组别: row.group,
      账号: row.account,
      平台SKU: row.sku,
      Bundle主SKU: row.bundleSku,
      ASIN分类: row.asinType,
      ASIN: row.asin,
      备注: row.note,
    }))
    .filter((row) => row.重复标记 === '重复')
    .sort((a, b) => {
      const groupCompare = String(b.重复组标准键).localeCompare(String(a.重复组标准键))
      if (groupCompare !== 0) return groupCompare
      const skuCompare = String(b.平台SKU).localeCompare(String(a.平台SKU))
      if (skuCompare !== 0) return skuCompare
      const asinCompare = String(b.ASIN).localeCompare(String(a.ASIN))
      if (asinCompare !== 0) return asinCompare
      return String(b.ASIN分类).localeCompare(String(a.ASIN分类))
    })
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

  const workbook = new ExcelJS.Workbook()
  const markedSheet = workbook.addWorksheet('重复监控处理表')
  const dedupedSheet = workbook.addWorksheet('剔重后参考')
  const optionSheet = workbook.addWorksheet('处理动作说明')

  const markedHeaders = Object.keys(markedData[0] ?? {
    重复标记: '',
    处理动作: '',
    重复组键: '',
    重复组标准键: '',
    运营: '',
    组别: '',
    账号: '',
    平台SKU: '',
    Bundle主SKU: '',
    ASIN分类: '',
    ASIN: '',
    备注: '',
  })
  markedSheet.addRow(markedHeaders)
  markedData.forEach((row) => markedSheet.addRow(markedHeaders.map((header) => row[header as keyof typeof row])))

  const dedupedHeaders = Object.keys(dedupedData[0] ?? {
    运营: '',
    组别: '',
    账号: '',
    平台SKU: '',
    Bundle主SKU: '',
    ASIN分类: '',
    ASIN: '',
    备注: '',
  })
  dedupedSheet.addRow(dedupedHeaders)
  dedupedData.forEach((row) => dedupedSheet.addRow(dedupedHeaders.map((header) => row[header as keyof typeof row])))
  optionSheet.addRows([
    ['处理动作可选值'],
    ['剔除'],
    ['保留'],
  ])

  ;[markedSheet, dedupedSheet].forEach((sheet) => {
    sheet.getRow(1).font = { bold: true }
    sheet.views = [{ state: 'frozen', ySplit: 1 }]
    sheet.columns.forEach((column) => {
      let maxLength = 12
      column?.eachCell?.({ includeEmpty: true }, (cell) => {
        maxLength = Math.max(maxLength, String(cell.value ?? '').length + 2)
      })
      if (column) column.width = Math.min(maxLength, 36)
    })
  })
  optionSheet.getRow(1).font = { bold: true }
  optionSheet.getColumn(1).width = 18

  const actionColumnIndex = markedHeaders.indexOf('处理动作') + 1
  const groupKeyColumnIndex = markedHeaders.indexOf('重复组键') + 1
  const normalizedGroupKeyColumnIndex = markedHeaders.indexOf('重复组标准键') + 1

  if (normalizedGroupKeyColumnIndex > 0) {
    markedSheet.getColumn(normalizedGroupKeyColumnIndex).hidden = true
  }

  const borderColor = 'FF8FA1B3'
  const thickBorder = { style: 'medium' as const, color: { argb: borderColor } }
  const thinBorder = { style: 'thin' as const, color: { argb: 'FFD6DEE8' } }

  if (groupKeyColumnIndex > 0 && normalizedGroupKeyColumnIndex > 0) {
    const lastVisibleColumnIndex = markedHeaders.length - 1
    let groupStartRow = 2
    let currentGroupKey = markedData.length
      ? String(markedSheet.getCell(2, normalizedGroupKeyColumnIndex).value ?? '')
      : ''

    const applyGroupOutline = (startRow: number, endRow: number) => {
      if (!startRow || !endRow || endRow < startRow) return
      for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
        for (let colIndex = 1; colIndex <= lastVisibleColumnIndex; colIndex += 1) {
          const cell = markedSheet.getCell(rowIndex, colIndex)
          cell.border = {
            top: rowIndex === startRow ? thickBorder : thinBorder,
            bottom: rowIndex === endRow ? thickBorder : thinBorder,
            left: colIndex === 1 ? thickBorder : thinBorder,
            right: colIndex === lastVisibleColumnIndex ? thickBorder : thinBorder,
          }
        }
      }
    }

    for (let rowIndex = 3; rowIndex <= markedData.length + 1; rowIndex += 1) {
      const nextGroupKey = String(markedSheet.getCell(rowIndex, normalizedGroupKeyColumnIndex).value ?? '')
      if (nextGroupKey !== currentGroupKey) {
        applyGroupOutline(groupStartRow, rowIndex - 1)
        groupStartRow = rowIndex
        currentGroupKey = nextGroupKey
      }
    }
    if (markedData.length) applyGroupOutline(groupStartRow, markedData.length + 1)
  }

  if (actionColumnIndex > 0) {
    for (let rowIndex = 2; rowIndex <= markedData.length + 1; rowIndex += 1) {
      markedSheet.getCell(rowIndex, actionColumnIndex).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['\'处理动作说明\'!$A$2:$A$3'],
        showErrorMessage: true,
        errorStyle: 'error',
        errorTitle: '仅可选择预设动作',
        error: '处理动作只允许填写“剔除”或“保留”。',
        showInputMessage: true,
        promptTitle: '处理动作',
        prompt: '请选择“剔除”或“保留”。',
      }
    }
  }

  const buffer = await workbook.xlsx.writeBuffer()
  downloadArrayBufferFile(buffer as ArrayBuffer, '重复监控关系检查表.xlsx')
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
  const [sourceReports, setSourceReports] = useState<SourceReports>(initialState.sourceReports)
  const [monitorSyncReport, setMonitorSyncReport] = useState<MonitorSyncReport>(initialState.monitorSyncReport)
  const [monitorDailyAdditions, setMonitorDailyAdditions] = useState<MonitorDailyAdditions>(initialState.monitorDailyAdditions)
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
  const [sourceUploadReports, setSourceUploadReports] = useState<SourceUploadReports>(emptySourceUploadReports)
  const [dataPreviewMode, setDataPreviewMode] = useState<DataPreviewMode>(null)
  const [pendingMonitorDelete, setPendingMonitorDelete] = useState<MonitorRow[] | null>(null)
  const [selectedMonitorRowKeys, setSelectedMonitorRowKeys] = useState<string[]>([])
  const [sharedSyncAt, setSharedSyncAt] = useState('')
  const [resultFilters, setResultFilters] = useState<Record<ResultFilterKey, string>>({
    owner: '',
    sku: '',
    asinType: '',
    brand: '',
    asin: '',
    price: '',
    todayRank: '',
    yesterdayRank: '',
    stockStatus: '',
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
    stockStatus: 112,
  })
  const [resultsPanelHeight, setResultsPanelHeight] = useState(430)
  const [hasLoadedSharedState, setHasLoadedSharedState] = useState(false)
  const [status, setStatus] = useState(
    `已自动载入原 Excel：${initialState.monitorRows.length} 条监控清单、${initialState.mappingRows.length} 条映射、${initialState.keepaRows.length} 条 Keepa。趋势仅保留最近 5 天。`,
  )

  const buildStoredStatePayload = (): StoredState => ({
    keepaRows: compactKeepaRowsForStorage(keepaRows),
    yesterdayKeepaRows,
    mappingRows,
    monitorRows,
    onlineRows: [],
    sourceReports,
    history,
    todayBuyBox,
    yesterdayBuyBox,
    keepaUploadReports,
    monitorSyncReport,
    monitorDailyAdditions,
  })

  useEffect(() => {
    let active = true
    if (!isSharedStateConfigured) {
      setHasLoadedSharedState(true)
      return undefined
    }

    const loadSharedState = async () => {
      try {
        const remote = await readSharedState()
        if (!active) return
        if (remote?.payload) {
          const nextState = deserializeStoredState(remote.payload as Partial<StoredState>)
          if (nextState) {
            setKeepaRows(nextState.keepaRows)
            setYesterdayKeepaRows(nextState.yesterdayKeepaRows)
            setMappingRows(nextState.mappingRows)
            setMonitorRows(nextState.monitorRows)
            setSourceReports(nextState.sourceReports)
            setMonitorSyncReport(nextState.monitorSyncReport)
            setMonitorDailyAdditions(nextState.monitorDailyAdditions)
            setHistory(nextState.history)
            setTodayBuyBox(nextState.todayBuyBox)
            setYesterdayBuyBox(nextState.yesterdayBuyBox)
            setKeepaUploadReports(nextState.keepaUploadReports)
            setSharedSyncAt(String(remote.updated_at ?? ''))
            setStatus(`已从共享云端载入最新数据${remote.updated_at ? `（更新时间 ${remote.updated_at}）` : ''}。`)
          }
        } else {
          await writeSharedState(buildStoredStatePayload())
          if (!active) return
          setSharedSyncAt(new Date().toISOString())
          setStatus(`已初始化共享云端数据表：${getSharedStateConfig().table}。`)
        }
      } catch (error) {
        if (!active) return
        const message = error instanceof Error ? error.message : '共享云端数据读取失败。'
        setStatus(`共享数据同步失败，当前先使用本地缓存。${message}`)
      } finally {
        if (!active) return
        setHasLoadedSharedState(true)
      }
    }

    void loadSharedState()

    return () => {
      active = false
    }
  }, [])

  useEffect(() => {
    const payload = buildStoredStatePayload()
    try {
      localStorage.setItem(storageKey, JSON.stringify(payload))
    } catch {
      setStatus('数据已在当前页面完成解析；完整状态超过浏览器容量，但映射信息将继续使用独立缓存保留。')
    }
  }, [keepaRows, keepaUploadReports, mappingRows, monitorDailyAdditions, monitorRows, monitorSyncReport, sourceReports, todayBuyBox, yesterdayBuyBox, yesterdayKeepaRows])

  useEffect(() => {
    try {
      localStorage.setItem(mappingStorageKey, JSON.stringify({ rows: mappingRows, report: sourceReports.mapping }))
    } catch {
      setStatus('映射已在当前页面更新，但浏览器无法保存独立映射缓存；请导出映射备份后检查浏览器存储空间。')
    }
  }, [mappingRows, sourceReports.mapping])

  useEffect(() => {
    if (!isSharedStateConfigured || !hasLoadedSharedState) return undefined
    const payload = buildStoredStatePayload()
    const timer = window.setTimeout(() => {
      void writeSharedState(payload)
        .then(() => {
          setSharedSyncAt(new Date().toISOString())
        })
        .catch((error) => {
          const message = error instanceof Error ? error.message : '共享云端数据写入失败。'
          setStatus(`共享数据写入失败，当前已保留在本地缓存。${message}`)
        })
    }, 600)
    return () => window.clearTimeout(timer)
  }, [hasLoadedSharedState, keepaRows, keepaUploadReports, mappingRows, monitorDailyAdditions, monitorRows, monitorSyncReport, sourceReports, todayBuyBox, yesterdayBuyBox, yesterdayKeepaRows])

  useEffect(() => {
    if (!pendingKeepaUpload && !pendingMonitorDelete) return undefined
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        if (pendingMonitorDelete) setPendingMonitorDelete(null)
        else cancelKeepaUpload()
      }
    }
    document.addEventListener('keydown', closeOnEscape)
    return () => document.removeEventListener('keydown', closeOnEscape)
  }, [pendingKeepaUpload, pendingMonitorDelete])

  const keepaByAsin = useMemo(() => new Map(keepaRows.map((row) => [normalized(row.asin), row])), [keepaRows])
  const yesterdayKeepaByAsin = useMemo(
    () => new Map(yesterdayKeepaRows.map((row) => [normalized(row.asin), row])),
    [yesterdayKeepaRows],
  )
  const mappingIndex = useMemo(() => createMappingIndex(mappingRows), [mappingRows])
  const mappedMonitorRows = useMemo(
    () => applyMappingsToMonitorRows(monitorRows, mappingIndex),
    [mappingIndex, monitorRows],
  )
  const combinedMonitorRows = mappedMonitorRows
  const todayAddedMonitorRows = useMemo(() => {
    const today = new Date().toLocaleDateString('sv-SE')
    if (monitorDailyAdditions.date !== today) return []
    const byKey = new Map(monitorRows.map((row) => [monitorRowKey(row), row]))
    return monitorDailyAdditions.keys.flatMap((key) => {
      const row = byKey.get(key)
      return row ? [row] : []
    })
  }, [monitorDailyAdditions, monitorRows])
  const resizingColumnRef = useRef<{ key: ResultColumnKey; startX: number; startWidth: number } | null>(null)

  useEffect(() => {
    if (yesterdayBuyBox.date && yesterdayKeepaRows.length) {
      setYesterdayBuyBox(buildBuyBoxBoard(
        combinedMonitorRows,
        mappingIndex,
        yesterdayKeepaRows,
        [],
        yesterdayBuyBox.date,
        'snapshot',
      ))
    }
    if (todayBuyBox.date && keepaRows.length) {
      setTodayBuyBox(buildBuyBoxBoard(
        combinedMonitorRows,
        mappingIndex,
        keepaRows,
        yesterdayKeepaRows,
        todayBuyBox.date,
        'changes',
      ))
    }
  }, [combinedMonitorRows, keepaRows, mappingIndex, todayBuyBox.date, yesterdayBuyBox.date, yesterdayKeepaRows])

  const enrichedRows = useMemo(
    () =>
      combinedMonitorRows.map((row) => {
        const keepa = keepaByAsin.get(normalized(row.asin))
        const yesterdayKeepa = yesterdayKeepaByAsin.get(normalized(row.asin))
        const mappingResult = resolveMapping(row, mappingIndex)
        const mapping = mappingResult.mapping ?? undefined
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
          stockStatus: mapping ? '备货' : '不备货',
          mappingStatus: mappingResult.status,
          matchedSku: mappingResult.matchedSku,
        }
      }),
    [combinedMonitorRows, keepaByAsin, mappingIndex, yesterdayKeepaByAsin],
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
        (!columnNeedles.yesterdayRank || normalized(yesterdayRank).includes(columnNeedles.yesterdayRank)) &&
        (!columnNeedles.stockStatus || normalized(row.stockStatus).includes(columnNeedles.stockStatus))

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

  const historyDayCount = useMemo(() => new Set(history.map((point) => point.date)).size, [history])

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
      const relationKey = `${normalized(row.sku)}::${normalized(row.asin)}`
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
      const [sku, asin] = key.split('::')
      const duplicateIndexes = duplicateIndexMap.get(key) ?? []
      const rows = duplicateIndexes.map((index) => enrichedRows[index]).filter(Boolean)
      duplicateRelations.push({
        id: `duplicate-${key}`,
        message: `${sku.toUpperCase()} 下 ${asin.toUpperCase()} 重复监控`,
        asin: asin.toUpperCase(),
        sku: sku.toUpperCase(),
        owner: [...new Set(rows.map((row) => row.owner || '未填运营'))].join('、'),
        group: rows[0]?.group,
        account: rows[0]?.account,
        category: rows[0]?.asinType,
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
    () => {
      const seenSkus = new Set<string>()
      return enrichedRows.filter((row) => {
        if (row.hasMapping) return false
        const skuKey = normalized(row.sku)
        if (!skuKey || seenSkus.has(skuKey)) return false
        seenSkus.add(skuKey)
        return true
      })
    },
    [enrichedRows],
  )

  const ambiguousMonitorRows = useMemo(
    () => enrichedRows.filter((row) => row.mappingStatus === 'ambiguous'),
    [enrichedRows],
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
  const visibleRowKeys = visibleRows.map(monitorRowKey)
  const allVisibleRowsSelected = Boolean(visibleRowKeys.length) && visibleRowKeys.every((key) => selectedMonitorRowKeys.includes(key))

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

  const recordMonitorSync = (
    action: MonitorSyncReport['action'],
    added: number,
    deleted: number,
    total: number,
  ) => {
    const updatedAt = new Date().toISOString()
    setMonitorSyncReport({ updatedAt, added, deleted, total, action })
    setSourceReports((current) => ({
      ...current,
      monitor: {
        ...current.monitor,
        imported: total,
        updatedAt: formatSyncTime(updatedAt),
      },
    }))
  }

  const recordDailyMonitorAdditions = (rows: MonitorRow[]) => {
    if (!rows.length) return
    const date = new Date().toLocaleDateString('sv-SE')
    const nextKeys = rows.map(monitorRowKey)
    setMonitorDailyAdditions((current) => ({
      date,
      keys: [...new Set([...(current.date === date ? current.keys : []), ...nextKeys])],
    }))
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
        mappingIndex,
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

  const copyLatestMonitorAsins = async () => {
    const asins = [...new Set(
      combinedMonitorRows
        .map((row) => row.asin.trim())
        .filter(Boolean),
    )]

    if (!asins.length) {
      setStatus('当前没有可复制的监控 ASIN。')
      return
    }

    const text = asins.join('\n')

    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text)
      } else {
        const textarea = document.createElement('textarea')
        textarea.value = text
        textarea.setAttribute('readonly', 'true')
        textarea.style.position = 'fixed'
        textarea.style.opacity = '0'
        document.body.appendChild(textarea)
        textarea.select()
        document.execCommand('copy')
        document.body.removeChild(textarea)
      }
      setStatus(`已复制最新监控 ASIN ${asins.length} 个，可直接粘贴到 Keepa。`)
    } catch {
      setStatus('复制失败，请检查浏览器剪贴板权限后重试。')
    }
  }

  const handleDuplicateRefreshUpload = async (file: File | null) => {
    if (!file) return
    setStatus(`正在解析 ${file.name} ...`)
    try {
      const rows = await readWorkbookRows(file, 'monitor')
      const actions = parseDuplicateActionRows(rows)
      if (!actions.length) throw new Error('未识别到有效重复监控处理动作，请填写“剔除”或“保留”后再上传。')
      const result = applyDuplicateActionsAcrossSources(monitorRows, [], actions)
      setMonitorRows(result.monitorRows)
      setSelectedAsin(result.monitorRows[0]?.asin ?? '')
      recordMonitorSync('duplicate-cleanup', 0, monitorRows.length - result.monitorRows.length, result.monitorRows.length)
      const noteLines = [
        `已处理 ${result.touchedCount} 条重复监控指令`,
        result.removedKeys.length ? `已剔除重复监控组 ${result.removedKeys.length} 个` : '',
        result.keptKeys.length ? `保留重复监控组 ${result.keptKeys.length} 个` : '',
      ].filter(Boolean)
      setUploadSummary({
        kind: 'monitor',
        fileName: file.name,
        imported: result.monitorRows.length,
        notes: noteLines,
        errors: [],
      })
      setStatus(noteLines.join('；'))
    } catch (error) {
      const message = error instanceof Error ? error.message : '重复监控处理表解析失败，原有监控未改变。'
      setUploadSummary({
        kind: 'monitor',
        fileName: file.name,
        imported: 0,
        notes: ['本次重复监控回传未写入，现有监控清单保持不变。'],
        errors: [message],
      })
      setStatus(message)
    }
  }

  const handleOnlineMonitorImportUpload = async (file: File | null) => {
    if (!file) return
    setStatus(`正在解析 ${file.name} ...`)
    try {
      const rows = await readWorkbookRows(file, 'monitor')
      const parsed = parseMonitor(rows)
      if (!parsed.length) throw new Error('批量导入表未识别到有效的平台SKU、ASIN分类与ASIN。')
      const mappedParsed = applyMappingsToOnlineRows(parsed, mappingIndex).map((row) => ({
        ...row,
        note: row.note || '批量导入',
      }))
      const result = upsertMonitorRows(monitorRows, mappedParsed)
      setMonitorRows(result.rows)
      recordDailyMonitorAdditions(result.addedRows)
      recordMonitorSync('import', result.added, 0, result.rows.length)
      setSelectedAsin(mappedParsed[0]?.asin ?? '')
      setUploadSummary({
        kind: 'monitor',
        fileName: file.name,
        imported: mappedParsed.length,
        notes: [
          `已解析 ${mappedParsed.length} 条，其中新增 ${result.added} 条`,
          '已写回完整 SKU / ASIN 监控清单，并进入云端自动同步队列',
          '运营、组别、账号优先保留导入值，仅空值使用当前映射补齐',
          '若平台SKU暂无映射，将保留手动填写或待后续映射刷新',
        ],
        errors: [],
      })
      setStatus(`已导入 ${mappedParsed.length} 条监控关系，新增 ${result.added} 条；完整监控清单共 ${result.rows.length} 条。`)
    } catch (error) {
      const message = error instanceof Error ? error.message : '批量导入表解析失败，现有在线监控未改变。'
      setUploadSummary({
        kind: 'monitor',
        fileName: file?.name ?? '',
        imported: 0,
        notes: ['本次批量导入未写入，现有在线监控保持不变。'],
        errors: [message],
      })
      setStatus(message)
    }
  }

  const handleUpload = async (kind: UploadKind, file: File | null) => {
    if (!file) return
    if (kind === 'keepa') {
      setPendingKeepaUpload({ file })
      setStatus('请选择将这份 Keepa 数据作为昨日基准或今日数据。')
      return
    }
    setStatus(`正在解析 ${file.name} ...`)
    setSourceUploadReports((current) => ({
      ...current,
      [kind]: { status: 'processing', fileName: file.name, imported: 0, notes: ['正在完整校验文件，现有云端数据暂不变更。'], errors: [] },
    }))
    try {
      const rows = await readWorkbookRows(file, kind)
      const updatedAt = new Date().toLocaleString('zh-CN', { hour12: false })
      if (kind === 'mapping') {
        const parsed = parseMapping(rows)
        if (!parsed.length) throw new Error('映射文件没有识别到有效的平台 SKU 行，现有映射信息未改变。')
        const analysis = analyzeMappingRows(rows, parsed)
        if (analysis.missingHeaders.length) throw new Error(`映射模板缺少必填列：${analysis.missingHeaders.join('、')}。请下载统一模板后重试，现有映射未改变。`)
        if (analysis.invalidRows.length) {
          const samples = analysis.invalidRows.slice(0, 5).map((item) => `数据第 ${item.row} 行缺少${item.missing.join('/')}`).join('；')
          throw new Error(`映射文件有 ${analysis.invalidRows.length} 行必填字段不完整：${samples}${analysis.invalidRows.length > 5 ? '……' : ''}。现有映射未改变。`)
        }
        if (analysis.compositeConflicts.length) throw new Error(`发现 ${analysis.compositeConflicts.length} 个“平台SKU + 店铺别名”重复且归属冲突，无法安全覆盖。现有映射未改变。`)
        const validated = uniqueMappingRows(parsed)
        const validatedIndex = createMappingIndex(validated)
        setMappingRows(validated)
        setMonitorRows((current) => applyMappingsToMonitorRows(current, validatedIndex))
        setSourceReports((current) => ({
          ...current,
          mapping: { fileName: file.name, imported: validated.length, updatedAt, source: 'upload' },
        }))
        const notes = [
          `完整校验通过，固定映射源由原 ${mappingRows.length} 条替换为 ${validated.length} 条`,
          `唯一键：平台SKU + 店铺别名；跨账号重复 SKU ${analysis.multiAccountSkus} 个`,
          `精确重复行已去重 ${parsed.length - validated.length} 条；复合键归属冲突 0 个`,
          isSharedStateConfigured
            ? '监控清单原归属未被覆盖，仅空字段已补齐；变更进入云端自动同步队列'
            : '监控清单原归属未被覆盖，仅空字段已补齐并保存本地；当前环境未配置 Supabase',
        ]
        setUploadSummary({
          kind,
          fileName: file.name,
          imported: validated.length,
          notes,
          errors: [],
        })
        setSourceUploadReports((current) => ({ ...current, mapping: { status: 'success', fileName: file.name, imported: validated.length, notes, errors: [] } }))
        setStatus(
          `映射上传成功：${validated.length} 条；监控清单已有运营、组别、账号保持不变，仅补齐空字段${isSharedStateConfigured ? '并等待云端同步' : '，当前已保存本地（云端未配置）'}。`,
        )
      }
      if (kind === 'monitor') {
        const parsed = parseMonitor(rows)
        if (!parsed.length) throw new Error('监控清单没有识别到有效的 SKU 和 ASIN，现有固定监控源未改变。')
        const prepared = applyMappingsToMonitorRows(parsed, mappingIndex)
        const previousKeys = new Set(monitorRows.map(monitorRowKey))
        const nextKeys = new Set(prepared.map(monitorRowKey))
        const added = prepared.filter((row) => !previousKeys.has(monitorRowKey(row))).length
        const deleted = monitorRows.filter((row) => !nextKeys.has(monitorRowKey(row))).length
        setMonitorRows(prepared)
        recordMonitorSync('upload', added, deleted, prepared.length)
        setSelectedAsin(prepared[0]?.asin ?? '')
        setSourceReports((current) => ({
          ...current,
          monitor: { fileName: file.name, imported: prepared.length, updatedAt, source: 'upload' },
        }))
        setUploadSummary({
          kind,
          fileName: file.name,
          imported: prepared.length,
          notes: [
            `完整监控清单已替换为 ${prepared.length} 条关系`,
            `相较替换前新增 ${added} 条、移除 ${deleted} 条`,
            '上传值为运营归属依据，仅空字段使用映射信息补齐',
          ],
          errors: [],
        })
        setSourceUploadReports((current) => ({
          ...current,
          monitor: {
            status: 'success',
            fileName: file.name,
            imported: prepared.length,
            notes: [`完整监控清单已替换为 ${prepared.length} 条关系`, `相较替换前新增 ${added} 条、移除 ${deleted} 条`],
            errors: [],
          },
        }))
        setStatus(`已更新完整监控清单：${prepared.length} 条；新增 ${added} 条、移除 ${deleted} 条。`)
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : `${uploadLabels[kind]}解析失败，原有数据未改变。`
      setUploadSummary({ kind, fileName: file.name, imported: 0, notes: ['本次上传未写入，固定数据源保持不变。'], errors: [message] })
      setSourceUploadReports((current) => ({
        ...current,
        [kind]: { status: 'error', fileName: file.name, imported: 0, notes: ['本次上传未写入，原有本地与云端数据保持不变。'], errors: [message] },
      }))
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

  const startOnlineEdit = (target: MonitorRow) => {
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
      if (key === 'sku' || key === 'account') {
        const nextRow = { ...row, [key]: value }
        const mapping = resolveMapping(nextRow, mappingIndex).mapping
        return {
          ...nextRow,
          owner: nextRow.owner || mapping?.owner || '',
          group: nextRow.group || mapping?.group || '',
          account: nextRow.account || mapping?.account || '',
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
      const mapping = resolveMapping({ sku: value, account: next.account }, mappingIndex).mapping
      rows[rowIndex] = {
        ...next,
        sku: value,
        owner: next.owner || mapping?.owner || '',
        group: next.group || mapping?.group || '',
        account: next.account || mapping?.account || '',
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
        const mapping = resolveMapping(row, mappingIndex).mapping
        return {
          ...row,
          sku: row.sku.trim(),
          asinType: row.asinType.trim(),
          asin: row.asin.trim(),
          owner: (row.owner || mapping?.owner || '').trim(),
          group: (row.group || mapping?.group || '').trim(),
          account: (row.account || mapping?.account || '').trim(),
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

    const result = upsertMonitorRows(monitorRows, prepared)
    setMonitorRows(result.rows)
    recordDailyMonitorAdditions(result.addedRows)
    recordMonitorSync('add', result.added, 0, result.rows.length)
    setSelectedAsin(prepared[0]?.asin ?? '')
    setBatchOnlineRows(Array.from({ length: 8 }, () => ({ ...emptyOnlineMonitor })))
    setStatus(`已保存 ${prepared.length} 条监控关系，其中新增 ${result.added} 条；已写回完整监控清单并等待云端同步。`)
  }

  const requestMonitorDelete = (row: MonitorRow) => {
    setPendingMonitorDelete([row])
  }

  const toggleMonitorRowSelection = (row: MonitorRow, checked: boolean) => {
    const key = monitorRowKey(row)
    setSelectedMonitorRowKeys((current) => checked
      ? [...new Set([...current, key])]
      : current.filter((item) => item !== key))
  }

  const toggleVisibleMonitorRows = (checked: boolean) => {
    setSelectedMonitorRowKeys((current) => checked
      ? [...new Set([...current, ...visibleRowKeys])]
      : current.filter((key) => !visibleRowKeys.includes(key)))
  }

  const requestSelectedMonitorRowsDelete = () => {
    if (!selectedMonitorRowKeys.length) return
    const selectedKeys = new Set(selectedMonitorRowKeys)
    const selectedRows = monitorRows.filter((row) => selectedKeys.has(monitorRowKey(row)))
    if (selectedRows.length) setPendingMonitorDelete(selectedRows)
  }

  const confirmMonitorDelete = () => {
    const targets = pendingMonitorDelete
    if (!targets?.length) return
    const targetKeys = new Set(targets.map(monitorRowKey))
    const nextRows = monitorRows.filter((row) => !targetKeys.has(monitorRowKey(row)))
    const deleted = monitorRows.length - nextRows.length
    setMonitorRows(nextRows)
    setMonitorDailyAdditions((current) => ({
      ...current,
      keys: current.keys.filter((key) => !targetKeys.has(key)),
    }))
    setPendingMonitorDelete(null)
    setSelectedMonitorRowKeys((current) => current.filter((key) => !targetKeys.has(key)))
    recordMonitorSync('delete', 0, deleted, nextRows.length)
    if (targets.some((target) => normalized(selectedAsin) === normalized(target.asin))) setSelectedAsin(nextRows[0]?.asin ?? '')
    setStatus(`已从完整 SKU / ASIN 监控清单删除 ${deleted} 条关系，并进入云端自动同步队列。`)
  }

  const refreshOnlineMappings = () => {
    setIsRefreshingOnline(true)
    window.setTimeout(() => {
      setMonitorRows((rows) => applyMappingsToMonitorRows(rows, mappingIndex))
      setBatchOnlineRows((rows) => rows.map((row) => {
        const mapping = resolveMapping(row, mappingIndex).mapping
        return mapping ? {
          ...row,
          owner: row.owner || mapping.owner,
          group: row.group || mapping.group,
          account: row.account || mapping.account,
        } : row
      }))
      setIsRefreshingOnline(false)
      setStatus(`已检查 ${monitorRows.length} 条监控关系；已有归属保持不变，仅空字段按映射补齐。`)
    }, 450)
  }

  const resetDashboardFilters = () => {
    setOwnerQuery('')
    setSkuQuery('')
    setAsinQuery('')
    setGroupQuery('')
    setAccountQuery('')
    setBrandQuery('')
    setResultFilters({
      owner: '',
      sku: '',
      asinType: '',
      brand: '',
      asin: '',
      price: '',
      todayRank: '',
      yesterdayRank: '',
      stockStatus: '',
    })
    setOpenResultFilter(null)
    setStatus('已清空运营看板筛选条件。')
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

        <section className="erp-nav-section">
          <div className="panel-title"><Upload size={16} />数据更新</div>
          <div className="erp-nav-list">
            <button className={activePage === 'data-update' ? 'active-nav' : ''} type="button" onClick={() => openPage('data-update')}>
              <FileSpreadsheet size={15} />数据更新
              <span>{keepaRows.length}</span>
            </button>
          </div>
        </section>

        <section className="erp-nav-section">
          <div className="panel-title"><Bell size={16} />运营看板</div>
          <div className="erp-nav-list">
            <button className={activePage === 'dashboard' ? 'active-nav' : ''} type="button" onClick={() => openPage('dashboard')}>
              <BarChart3 size={15} />运营看板
            </button>
            <button className={activePage === 'buybox' ? 'active-nav' : ''} type="button" onClick={() => openPage('buybox')}>
              <Database size={15} />BuyBox丢失|恢复
              <span>{todayBuyBox.lost.length + todayBuyBox.recovered.length}</span>
            </button>
          </div>
        </section>

        <section className="erp-nav-section">
          <div className="panel-title"><Plus size={16} />添加监控ASIN</div>
          <div className="erp-nav-list">
            <button className={activePage === 'add-monitor' ? 'active-nav' : ''} type="button" onClick={() => openPage('add-monitor')}>
              <Plus size={15} />添加监控ASIN
              <span>{monitorRows.length}</span>
            </button>
          </div>
        </section>

        <p className="status-text">{status}</p>
      </aside>

      <section className="workspace erp-workspace">
        <div className="erp-tabbar">
          {openPages.map((page) => (
            <button className={activePage === page ? 'erp-tab active' : 'erp-tab'} key={page} type="button" onClick={() => setActivePage(page)}>
              <span>{workspacePageLabels[page]}</span>
              {openPages.length > 1 || page !== 'dashboard' ? (
                <CircleX
                  size={14}
                  onClick={(event) => {
                    event.stopPropagation()
                    closePage(page)
                  }}
                />
              ) : null}
            </button>
          ))}
        </div>

        <div className="erp-page">
          {activePage === 'dashboard' ? (
            <>
              <header className="topbar">
                <div className="topbar-intro">
                  <span className="eyebrow">运营检索台</span>
                  <h2>按运营、SKU、ASIN 快速定位监控关系</h2>
                </div>
                <div className="filter-grid dashboard-filter-grid">
                  <div className="search-box"><Search size={18} /><input list="owner-options" placeholder="检索运营/人名" value={ownerQuery} onChange={(event) => setOwnerQuery(event.target.value)} /><datalist id="owner-options">{ownerOptions.map((owner) => <option key={owner} value={owner} />)}</datalist></div>
                  <div className="search-box"><input list="sku-options" placeholder="检索 SKU" value={skuQuery} onChange={(event) => setSkuQuery(event.target.value)} /><datalist id="sku-options">{skuOptions.map((sku) => <option key={sku} value={sku} />)}</datalist></div>
                  <div className="search-box"><input list="asin-options" placeholder="检索 ASIN" value={asinQuery} onChange={(event) => setAsinQuery(event.target.value)} /><datalist id="asin-options">{asinOptions.map((asin) => <option key={asin} value={asin} />)}</datalist></div>
                  <div className="search-box"><input list="group-options" placeholder="检索组别" value={groupQuery} onChange={(event) => setGroupQuery(event.target.value)} /><datalist id="group-options">{groupOptions.map((group) => <option key={group} value={group} />)}</datalist></div>
                  <div className="search-box"><input list="account-options" placeholder="检索账号" value={accountQuery} onChange={(event) => setAccountQuery(event.target.value)} /><datalist id="account-options">{accountOptions.map((account) => <option key={account} value={account} />)}</datalist></div>
                  <div className="search-box"><input list="brand-options" placeholder="检索品牌" value={brandQuery} onChange={(event) => setBrandQuery(event.target.value)} /><datalist id="brand-options">{brandOptions.map((brand) => <option key={brand} value={brand} />)}</datalist></div>
                  <button className="dashboard-reset-button" type="button" onClick={resetDashboardFilters}>重置</button>
                </div>
              </header>

              <section className="results-stack dashboard-results-stack">
                <div className="table-panel full-span-panel">
                  <div className="section-heading">
                    <h2>检索结果</h2>
                    <div className="results-heading-actions">
                      <span>{filteredRows.length} 条 · 已选 {selectedMonitorRowKeys.length} 条</span>
                      <button className="bulk-delete-button" disabled={!selectedMonitorRowKeys.length} type="button" onClick={requestSelectedMonitorRowsDelete}><Trash2 size={14} />一键删除</button>
                    </div>
                  </div>
                  <div className="data-table-wrap resizable-results-wrap" style={{ maxHeight: `${resultsPanelHeight}px` }}>
                    <table className="data-table">
                      <colgroup>
                        <col style={{ width: `${resultColumnWidths.owner}px` }} />
                        <col style={{ width: `${resultColumnWidths.sku}px` }} />
                        <col style={{ width: `${resultColumnWidths.asinType}px` }} />
                        <col style={{ width: `${resultColumnWidths.brand}px` }} />
                        <col style={{ width: `${resultColumnWidths.asin}px` }} />
                        <col style={{ width: `${resultColumnWidths.price}px` }} />
                        <col style={{ width: `${resultColumnWidths.todayRank}px` }} />
                        <col style={{ width: `${resultColumnWidths.yesterdayRank}px` }} />
                        <col style={{ width: `${resultColumnWidths.stockStatus}px` }} />
                        <col style={{ width: '72px' }} />
                        <col style={{ width: '64px' }} />
                      </colgroup>
                      <thead><tr>{renderResultHeader('owner', '运营')}{renderResultHeader('sku', 'SKU')}{renderResultHeader('asinType', '类型')}{renderResultHeader('brand', '品牌')}{renderResultHeader('asin', 'ASIN')}{renderResultHeader('price', '价格')}{renderResultHeader('todayRank', '今日排名')}{renderResultHeader('yesterdayRank', '昨日排名')}{renderResultHeader('stockStatus', '备货状态')}<th className="result-action-heading">操作</th><th className="result-select-heading"><label><input aria-label="全选当前检索结果" checked={allVisibleRowsSelected} type="checkbox" onChange={(event) => toggleVisibleMonitorRows(event.target.checked)} /><span>多选</span></label></th></tr></thead>
                      <tbody>{visibleRows.map((row, index) => {
                        const previous = visibleRows[index - 1]
                        const next = visibleRows[index + 1]
                        const isGroupStart = !previous || previous.sku !== row.sku
                        const isGroupEnd = !next || next.sku !== row.sku
                        const typeClass = normalized(row.asinType).includes('kmasin') ? 'type-km' : normalized(row.asinType).includes('竞对') ? 'type-competitor' : 'type-neutral'
                        const priceChange = getDirectMetricChange(row.keepa?.price, row.yesterdayKeepa?.price, 'price')
                        return <tr className={`${row.asin === selectedAsin ? 'selected-row' : ''} ${isGroupStart ? 'sku-group-start' : ''} ${isGroupEnd ? 'sku-group-end' : ''}`} key={`${row.sku}-${row.asin}-${index}`} onClick={() => setSelectedAsin(row.asin)}><td>{row.owner || '-'}</td><td className="sku-cell">{row.sku}</td><td><span className={`type-tag ${typeClass}`}>{row.asinType || '-'}</span></td><td>{row.keepa?.brand || '-'}</td><td className="asin-cell"><button className="asin-trigger" type="button" onClick={(event) => {
                          event.stopPropagation()
                          setSelectedAsin(row.asin)
                        }}>{row.asin}</button></td><td>{renderMetric(row.keepa?.price, priceChange, 'price')}</td><td>{renderRankComparison(row.keepa?.rank, row.yesterdayKeepa?.rank)}</td><td>{typeof row.yesterdayKeepa?.rank === 'number' ? row.yesterdayKeepa.rank.toLocaleString() : '-'}</td><td><span className={`stock-status-tag ${row.hasMapping ? 'stock-status-active' : 'stock-status-inactive'}`} title={row.mappingStatus === 'composite-primary' ? `A+B 已按主 SKU ${row.matchedSku} 匹配` : row.hasMapping ? '已匹配映射信息' : '未匹配映射信息'}>{row.stockStatus}</span></td><td className="result-action-cell"><button aria-label={`删除 ${row.sku} ${row.asin}`} className="row-icon danger" title="从监控清单删除" type="button" onClick={(event) => { event.stopPropagation(); requestMonitorDelete(row) }}><Trash2 size={15} /></button></td><td className="result-select-cell"><input aria-label={`选择 ${row.sku} ${row.asin}`} checked={selectedMonitorRowKeys.includes(monitorRowKey(row))} type="checkbox" onClick={(event) => event.stopPropagation()} onChange={(event) => toggleMonitorRowSelection(row, event.target.checked)} /></td></tr>
                      })}</tbody>
                    </table>
                  </div>
                  <div className="results-resize-bar">
                    <input
                      aria-label="调整检索结果高度"
                      className="results-height-slider"
                      max="900"
                      min="280"
                      type="range"
                      value={resultsPanelHeight}
                      onChange={(event) => setResultsPanelHeight(Number(event.target.value))}
                    />
                    <span>{resultsPanelHeight}px</span>
                  </div>
                </div>
              </section>

              <section className="chart-grid">
                <div className="chart-panel"><div className="section-heading"><h2>价格趋势</h2><span>Price</span></div><ResponsiveContainer height={230} width="100%"><LineChart data={selectedHistory}><CartesianGrid stroke="#e6ebf2" vertical={false} /><XAxis dataKey="date" tickLine={false} /><YAxis tickLine={false} width={54} /><Tooltip /><Line dataKey="price" stroke="#1f7a6d" strokeWidth={2.5} type="monotone" /></LineChart></ResponsiveContainer></div>
                <div className="chart-panel"><div className="section-heading"><h2>排名趋势</h2><span>Rank</span></div><ResponsiveContainer height={230} width="100%"><AreaChart data={selectedHistory}><CartesianGrid stroke="#e6ebf2" vertical={false} /><XAxis dataKey="date" tickLine={false} /><YAxis tickLine={false} width={66} /><Tooltip /><Area dataKey="rank" fill="#dfeeea" stroke="#5067a3" strokeWidth={2.5} type="monotone" /></AreaChart></ResponsiveContainer></div>
              </section>

              {selectedRows[0] ? (
                <section className="detail-panel asin-detail-card">
                  <div className="section-heading"><h2>ASIN 详情</h2><span>{selectedRows[0].asin}</span></div>
                  <div className="detail-stack">
                    <div>
                      <span className="eyebrow">商品</span>
                      <h3>{selectedRows[0].keepa?.title || selectedRows[0].asin}</h3>
                    </div>
                    {getImageCandidates(selectedRows[0].keepa?.image).length ? (
                      <div className="detail-image-wrap">
                        <img
                          alt={selectedRows[0].asin}
                          className="detail-image"
                          key={selectedRows[0].asin}
                          onLoad={(event) => {
                            event.currentTarget.style.display = 'block'
                            event.currentTarget.dataset.index = '0'
                          }}
                          onError={(event) => {
                            const candidates = getImageCandidates(selectedRows[0].keepa?.image)
                            const currentIndex = Number(event.currentTarget.dataset.index || '0')
                            const nextIndex = currentIndex + 1
                            if (nextIndex < candidates.length) {
                              event.currentTarget.dataset.index = String(nextIndex)
                              event.currentTarget.src = candidates[nextIndex]
                              return
                            }
                            event.currentTarget.style.display = 'none'
                          }}
                          src={getPrimaryImageUrl(selectedRows[0].keepa?.image)}
                        />
                      </div>
                    ) : <div className="detail-image-wrap detail-image-empty">暂无主图</div>}
                  </div>
                </section>
              ) : null}

            </>
          ) : null}

          {activePage === 'data-update' ? (
            <>
              <header className="page-header-card">
                <div>
                  <span className="eyebrow">数据更新</span>
                  <h2>上传 Keepa、映射信息和 SKU / ASIN 监控清单</h2>
                </div>
              </header>

              <section className="stat-grid update-stat-grid">{updateStats.map((stat) => <div className="stat-card" key={stat.label}><span>{stat.label}</span><strong>{stat.value.toLocaleString()}</strong></div>)}</section>

              <section className="erp-grid erp-grid-2">
                <section className="sidebar-info-panel page-panel">
                  <div className="panel-title"><Upload size={16} />数据上传</div>
                  <div className="upload-stack">
                    {(Object.keys(uploadLabels) as UploadKind[]).map((kind) => (
                      <div className={`upload-row ${kind === 'keepa' ? '' : 'upload-row-preview'}`} key={kind}>
                        <label className="upload-button">
                          <Upload size={16} />
                          <span>{uploadLabels[kind]}</span>
                          <input accept=".xlsx,.xls,.csv" type="file" onChange={(event) => { const file = event.currentTarget.files?.[0] ?? null; event.currentTarget.value = ''; void handleUpload(kind, file) }} />
                        </label>
                        <button className="icon-button" title={downloadLabels[kind]} type="button" onClick={() => handleSidebarDownload(kind)}>
                          <Download size={16} />
                        </button>
                        {kind !== 'keepa' ? (
                          <button className="icon-button" title={`查看全部${uploadLabels[kind]}数据`} type="button" onClick={() => setDataPreviewMode(kind)}>
                            <Eye size={16} />
                          </button>
                        ) : null}
                      </div>
                    ))}
                  </div>
                  <button className="reset-button" type="button" onClick={() => void copyLatestMonitorAsins()}>
                    <FileSpreadsheet size={16} />
                    一键复制最新监控ASIN
                  </button>
                </section>

                <section className="sidebar-note-panel page-panel">
                  <div className="panel-title"><AlertTriangle size={16} />缺少 Keepa</div>
                  <strong>{missingKeepaItems.length}</strong>
                  <p className="status-text">区分未出现在本次上传、已读取但价格缺失、已读取但排名缺失。</p>
                  <button className="reset-button" type="button" onClick={() => exportMissingKeepaReport(missingKeepaItems)}>下载分类处理表</button>
                </section>

                <section className="sidebar-info-panel page-panel">
                  <div className="panel-title"><FileSpreadsheet size={16} />同步摘要</div>
                  <div className="sidebar-metric-list">
                    <div className="sidebar-metric-item"><span>待补映射</span><strong>{unmatchedOnlineRows.length}</strong></div>
                    <div className="sidebar-metric-item"><span>完整监控清单</span><strong>{monitorRows.length}</strong></div>
                    <div className="sidebar-metric-item"><span>匹配冲突</span><strong>{ambiguousMonitorRows.length}</strong></div>
                    <div className="sidebar-metric-item"><span>趋势天数</span><strong>{historyDayCount}</strong></div>
                    <div className="sidebar-metric-item"><span>云端同步</span><strong>{formatSyncTime(sharedSyncAt)}</strong></div>
                  </div>
                </section>

                <section className="sidebar-note-panel page-panel">
                  <div className="panel-title"><Database size={16} />监控清单待补映射</div>
                  <strong>{unmatchedOnlineRows.length}</strong>
                  <p className="status-text">按平台 SKU 去重统计；A+B 会先提取 A 重新匹配，清单已有归属不会被覆盖。</p>
                  <button className="reset-button" disabled={!unmatchedOnlineRows.length} type="button" onClick={() => exportUnmatchedOnlineRows(unmatchedOnlineRows)}>导出待补映射SKU</button>
                </section>

                <section className="sidebar-note-panel page-panel">
                  <div className="panel-title"><AlertTriangle size={16} />重复监控关系</div>
                  <strong>{duplicateItems.length}</strong>
                  <p className="status-text">仅同一平台SKU下监控同一ASIN重复时才计为重复，可导出后选择保留或剔除再回传刷新。</p>
                  <div className="upload-stack">
                    <button className="reset-button" disabled={!duplicateItems.length} type="button" onClick={() => void exportMonitorWithDuplicateMarks(combinedMonitorRows)}>导出重复处理表</button>
                    <label className="upload-button secondary-upload-button">
                      <RefreshCw size={16} />
                      <span>上传重复处理结果</span>
                      <input accept=".xlsx,.xls,.csv" type="file" onChange={(event) => { const file = event.currentTarget.files?.[0] ?? null; event.currentTarget.value = ''; void handleDuplicateRefreshUpload(file) }} />
                    </label>
                  </div>
                </section>

                <section className="sidebar-info-panel page-panel">
                  <div className="panel-title"><Upload size={16} />上传后说明</div>
                  <div className="sidebar-upload-summary">
                    <section className={`sidebar-upload-role upload-role-${monitorSyncReport.updatedAt ? 'success' : 'idle'}`}>
                      <div className="sidebar-upload-role-heading"><strong>监控清单回传</strong><span>{monitorSyncActionLabels[monitorSyncReport.action]}</span></div>
                      <p>回传时间：{formatSyncTime(monitorSyncReport.updatedAt)}</p>
                      <p>本次新增 {monitorSyncReport.added.toLocaleString()} 条 · 删除 {monitorSyncReport.deleted.toLocaleString()} 条</p>
                      <p>当前完整清单 {monitorRows.length.toLocaleString()} 条</p>
                      <p>{isSharedStateConfigured ? `云端最近写入：${formatSyncTime(sharedSyncAt)}` : '当前未配置 Supabase，仅保存本地'}</p>
                    </section>
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
                    {(['mapping', 'monitor'] as const).map((kind) => {
                      const report = sourceUploadReports[kind]
                      const statusLabel = report.status === 'success' ? '上传成功' : report.status === 'error' ? '上传失败' : report.status === 'processing' ? '校验中' : '等待上传'
                      return (
                        <section className={`sidebar-upload-role upload-role-${report.status}`} key={kind}>
                          <div className="sidebar-upload-role-heading"><strong>{uploadLabels[kind]}</strong><span>{statusLabel}</span></div>
                          <p>{report.fileName || '尚未选择文件'}</p>
                          {report.status === 'success' ? <p>已解析并写入 {report.imported.toLocaleString()} 条</p> : null}
                          {report.notes.length ? <ul className="report-list compact-report-list">{report.notes.map((note) => <li key={note}>{note}</li>)}</ul> : null}
                          {report.errors.length ? <ul className="report-list report-error compact-report-list">{report.errors.map((error) => <li key={error}>{error}</li>)}</ul> : null}
                        </section>
                      )
                    })}
                  </div>
                </section>
              </section>
            </>
          ) : null}

          {activePage === 'buybox' ? (
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
          ) : null}

          {activePage === 'add-monitor' ? (
            <>
              <section className="editor-panel">
                <div className="section-heading editor-heading">
                  <div><span className="eyebrow">支持直接粘贴线下表格，平台SKU会自动匹配映射</span><h2>在线添加</h2></div>
                  <div className="editor-actions">
                    <button className="icon-command" title="重新匹配映射信息" type="button" onClick={refreshOnlineMappings}><RefreshCw className={isRefreshingOnline ? 'spin-icon' : ''} size={16} /></button>
                    <button className="active-action" type="button" onClick={startBatchOnlineAdd}><FileSpreadsheet size={16} />在线批量添加</button>
                    <button type="button" onClick={saveBatchOnlineRows}><Save size={16} />保存</button>
                    <button type="button" onClick={() => exportOnlineImportTemplate(mappingRows)}><Download size={16} />下载批量导入表</button>
                    <label className="upload-button inline-upload-button">
                      <Upload size={16} />
                      <span>表格导入</span>
                      <input accept=".xlsx,.xls,.csv" type="file" onChange={(event) => { const file = event.currentTarget.files?.[0] ?? null; event.currentTarget.value = ''; void handleOnlineMonitorImportUpload(file) }} />
                    </label>
                  </div>
                </div>
                <div className="mini-table-wrap batch-online-wrap batch-entry-panel">
                  <table className="data-table mini-table batch-entry-table">
                    <colgroup>
                      <col style={{ width: '18%' }} />
                      <col style={{ width: '14%' }} />
                      <col style={{ width: '18%' }} />
                      <col style={{ width: '16%' }} />
                      <col style={{ width: '16%' }} />
                      <col style={{ width: '16%' }} />
                      <col style={{ width: '72px' }} />
                    </colgroup>
                    <thead><tr><th>平台 SKU</th><th>ASIN 分类</th><th>ASIN</th><th>运营</th><th>组别</th><th>账号</th><th>操作</th></tr></thead>
                    <tbody>{batchOnlineRows.map((row, index) => {
                      const mappingResult = resolveMapping(row, mappingIndex)
                      const matched = mappingResult.mapping
                      const needsManual = Boolean(row.sku && !matched)
                      const needsAccount = mappingResult.status === 'ambiguous'
                      return <tr key={`batch-${index}`}>
                        <td><input className="table-input" placeholder="填写平台SKU" value={row.sku} onChange={(event) => updateBatchOnlineCell(index, 'sku', event.target.value)} onPaste={(event) => {
                          const handled = handleBatchPaste(index, 'sku', event.clipboardData.getData('text'))
                          if (handled) event.preventDefault()
                        }} /></td>
                        <td><select className="table-input" value={row.asinType} onChange={(event) => updateBatchOnlineCell(index, 'asinType', event.target.value)} onPaste={(event) => {
                          const handled = handleBatchPaste(index, 'asinType', event.clipboardData.getData('text'))
                          if (handled) event.preventDefault()
                        }}><option value="">请选择</option><option value="竞对ASIN">竞对ASIN</option><option value="KMASIN">KMASIN</option></select></td>
                        <td><input className="table-input" placeholder="填写ASIN，可直接粘贴整列" value={row.asin} onChange={(event) => updateBatchOnlineCell(index, 'asin', event.target.value)} onPaste={(event) => {
                          const handled = handleBatchPaste(index, 'asin', event.clipboardData.getData('text'))
                          if (handled) event.preventDefault()
                        }} /></td>
                        <td><input className={`table-input ${matched ? 'mapped-field' : ''}`} placeholder={needsManual ? '未匹配手动填写' : '自动匹配'} readOnly={Boolean(matched)} value={row.owner} onChange={(event) => updateBatchOnlineCell(index, 'owner', event.target.value)} onPaste={(event) => {
                          const handled = handleBatchPaste(index, 'owner', event.clipboardData.getData('text'))
                          if (handled) event.preventDefault()
                        }} /></td>
                        <td><input className={`table-input ${matched ? 'mapped-field' : ''}`} placeholder={needsManual ? '未匹配手动填写' : '自动匹配'} readOnly={Boolean(matched)} value={row.group} onChange={(event) => updateBatchOnlineCell(index, 'group', event.target.value)} onPaste={(event) => {
                          const handled = handleBatchPaste(index, 'group', event.clipboardData.getData('text'))
                          if (handled) event.preventDefault()
                        }} /></td>
                        <td><input className={`table-input ${matched ? 'mapped-field' : ''}`} placeholder={needsAccount ? '多个账号，请填写账号' : needsManual ? '未匹配手动填写' : '自动匹配'} readOnly={Boolean(matched)} value={row.account} onChange={(event) => updateBatchOnlineCell(index, 'account', event.target.value)} onPaste={(event) => {
                          const handled = handleBatchPaste(index, 'account', event.clipboardData.getData('text'))
                          if (handled) event.preventDefault()
                        }} /></td>
                        <td><button className="row-icon danger" title="删除本行" type="button" onClick={() => removeBatchOnlineLine(index)}><Trash2 size={14} /></button></td>
                      </tr>
                    })}</tbody>
                  </table>
                  <div className="batch-online-actions">
                    <button type="button" onClick={addBatchOnlineLine}><Plus size={16} />新增一行</button>
                  </div>
                </div>
                <div className="source-meta"><span>新增SKU监控 · {new Date().toLocaleDateString('sv-SE')}</span><strong>{todayAddedMonitorRows.length.toLocaleString()} 条</strong></div>
                <div className="mini-table-wrap online-table-wrap"><table className="data-table mini-table"><thead><tr><th>平台 SKU</th><th>ASIN 分类</th><th>ASIN</th><th>账号</th><th>组别</th><th>运营</th><th>匹配状态</th><th>操作</th></tr></thead><tbody>{todayAddedMonitorRows.length ? todayAddedMonitorRows.map((row, index) => {
                  const mappingResult = resolveMapping(row, mappingIndex)
                  const matched = Boolean(mappingResult.mapping)
                  const statusText = mappingResult.status === 'ambiguous' ? '账号冲突' : '手动填写'
                  return <tr key={`${monitorRowKey(row)}-${index}`}><td className="sku-cell">{row.sku}</td><td>{row.asinType}</td><td className="asin-cell">{row.asin}</td><td>{row.account || '-'}</td><td>{row.group || '-'}</td><td>{row.owner || '-'}</td><td>{matched ? <span className="source-tag source-tag-mapped">映射可用</span> : <span className="source-tag source-tag-manual"><AlertTriangle size={13} />{statusText}</span>}</td><td><button className="row-icon" title="编辑" type="button" onClick={() => startOnlineEdit(row)}><Edit3 size={14} /></button><button className="row-icon danger" title="从监控清单删除" type="button" onClick={() => requestMonitorDelete(row)}><Trash2 size={14} /></button></td></tr>
                }) : <tr><td className="empty-cell" colSpan={8}>今日暂无新增SKU监控</td></tr>}</tbody></table></div>
              </section>
            </>
          ) : null}
        </div>
      </section>

      {dataPreviewMode ? (
        <div className="modal-backdrop">
          <section aria-labelledby="data-preview-title" aria-modal="true" className="data-preview-dialog" role="dialog">
            <div className="data-preview-heading">
              <div>
                <span className="eyebrow">固定数据源 · 全量可视化</span>
                <h2 id="data-preview-title">{uploadLabels[dataPreviewMode]}</h2>
                <p>{sourceReports[dataPreviewMode].fileName} · 共 {(dataPreviewMode === 'mapping' ? mappingRows.length : monitorRows.length).toLocaleString()} 条 · {sourceReports[dataPreviewMode].updatedAt}</p>
              </div>
              <button className="row-icon" title="关闭数据预览" type="button" onClick={() => setDataPreviewMode(null)}><CircleX size={18} /></button>
            </div>
            <div className="data-preview-table-wrap">
              {dataPreviewMode === 'mapping' ? (
                <table className="data-table mini-table">
                  <thead><tr><th>序号</th><th>平台 SKU</th><th>系统 SKU</th><th>店铺别名</th><th>运营</th><th>小组</th></tr></thead>
                  <tbody>{mappingRows.map((row, index) => <tr key={`${mappingCompositeKey(row)}-${index}`}><td>{index + 1}</td><td className="sku-cell">{row.sku}</td><td>{row.systemSku}</td><td>{row.account}</td><td>{row.owner}</td><td>{row.group}</td></tr>)}</tbody>
                </table>
              ) : (
                <table className="data-table mini-table">
                  <thead><tr><th>序号</th><th>平台 SKU</th><th>ASIN 分类</th><th>ASIN</th><th>账号</th><th>运营</th><th>组别</th><th>备注</th></tr></thead>
                  <tbody>{monitorRows.map((row, index) => <tr key={`${monitorRowKey(row)}-${index}`}><td>{index + 1}</td><td className="sku-cell">{row.sku}</td><td>{row.asinType}</td><td className="asin-cell">{row.asin}</td><td>{row.account}</td><td>{row.owner}</td><td>{row.group}</td><td>{row.note}</td></tr>)}</tbody>
                </table>
              )}
            </div>
          </section>
        </div>
      ) : null}

      {pendingMonitorDelete ? (
        <div className="modal-backdrop">
          <section aria-describedby="monitor-delete-description" aria-labelledby="monitor-delete-title" aria-modal="true" className="save-snapshot-dialog delete-monitor-dialog" role="dialog">
            <div className="dialog-heading delete-dialog-heading"><Trash2 size={20} /><div><span className="eyebrow">同步修改完整监控清单</span><h2 id="monitor-delete-title">{pendingMonitorDelete.length > 1 ? `是否删除已选 ${pendingMonitorDelete.length} 条数据？` : '确定删除这条监控关系？'}</h2></div></div>
            <p id="monitor-delete-description">
              {pendingMonitorDelete.length > 1
                ? <>删除后，已选的 <strong>{pendingMonitorDelete.length}</strong> 条整行数据将同时从检索结果和 SKU / ASIN 监控清单移除，并自动回传 Supabase 云端。</>
                : <>删除后，<strong>{pendingMonitorDelete[0].sku}</strong> / <strong>{pendingMonitorDelete[0].asin}</strong> 将同时从检索结果和 SKU / ASIN 监控清单移除，并自动回传 Supabase 云端。</>}
            </p>
            <div className="dialog-actions">
              <button autoFocus className="dialog-button dialog-button-cancel" type="button" onClick={() => setPendingMonitorDelete(null)}>{pendingMonitorDelete.length > 1 ? '保留' : '不删除'}</button>
              <button className="dialog-button dialog-button-danger" type="button" onClick={confirmMonitorDelete}><Trash2 size={16} />{pendingMonitorDelete.length > 1 ? '删除' : '确定删除'}</button>
            </div>
          </section>
        </div>
      ) : null}

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
