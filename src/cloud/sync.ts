import {
  downloadGoogleFile,
  googleRevision,
  listGoogleFiles,
  uploadGoogleXlsx,
  createGoogleXlsx,
} from '@/cloud/googleFiles'
import { pullGoogleSheet, pushGoogleRows, pushGoogleSheet } from '@/cloud/googleSheets'
import {
  downloadMicrosoftFile,
  listRecentExcel,
  microsoftRevision,
  uploadMicrosoftXlsx,
  createMicrosoftXlsx,
} from '@/cloud/microsoftFiles'
import { pullMicrosoftWorkbook, pushMicrosoftRows, pushMicrosoftWorkbook } from '@/cloud/microsoftWorkbook'
import { preserveFormSheet } from '@/cloud/formCloud'
import { buildXlsxBlob } from '@/cloud/session'
import { GOOGLE_SHEET_MIME, type CloudBinding, type CloudFileRef, type CloudOpenMode, type CloudProviderId } from '@/cloud/types'
import { cloudBaseline, setCloudBaseline } from '@/cloud/baseline'
import { assertLiveWorkbook, diffLiveWorkbook, snapshotSheetMatrices } from '@/cloud/values'
import type { WorkbookData } from '@/types/excel'

export interface CloudOpenResult {
  binding: CloudBinding
  file?: File
  workbook?: WorkbookData
}

export { listGoogleFiles, listRecentExcel }

export async function openCloudFile(
  provider: CloudProviderId,
  mode: CloudOpenMode,
  file: CloudFileRef,
): Promise<CloudOpenResult> {
  if (mode === 'live') {
    if (provider === 'google' && file.mime !== GOOGLE_SHEET_MIME) {
      throw new Error('Chế độ trực tiếp của Google chỉ mở Google Sheet, không mở file .xlsx.')
    }
    const pulled = provider === 'google' ? await pullGoogleSheet(file) : await pullMicrosoftWorkbook(file)
    return {
      binding: { provider, mode, file, revision: pulled.revision },
      workbook: pulled.workbook,
    }
  }

  const downloaded = provider === 'google' ? await downloadGoogleFile(file) : await downloadMicrosoftFile(file)
  const revision = provider === 'google' ? await googleRevision(file.id) : await microsoftRevision(file)
  return {
    binding: { provider, mode, file, revision },
    file: downloaded,
  }
}

export async function remoteRevision(binding: CloudBinding): Promise<string> {
  if (binding.provider === 'google') return googleRevision(binding.file.id)
  return microsoftRevision(binding.file)
}

export function writesCellValues(binding: CloudBinding): boolean {
  return binding.mode === 'live' || binding.file.mime === GOOGLE_SHEET_MIME
}

export function pendingCloudDelta(sessionId: string, workbook: WorkbookData) {
  return diffLiveWorkbook(workbook, cloudBaseline(sessionId))
}

export function rememberCloudBaseline(sessionId: string, workbook: WorkbookData): void {
  try {
    setCloudBaseline(sessionId, snapshotSheetMatrices(workbook))
  } catch {
    setCloudBaseline(sessionId, new Map())
  }
}

export async function writeCloudWorkbook(
  binding: CloudBinding,
  workbook: WorkbookData,
  sessionId?: string,
  onProgress?: (message: string) => void,
): Promise<string> {
  if (writesCellValues(binding)) {
    assertLiveWorkbook(workbook)
    const delta = sessionId ? diffLiveWorkbook(workbook, cloudBaseline(sessionId)) : { kind: 'full' as const }
    if (delta.kind === 'rows' && delta.writes.length === 0 && delta.clears.length === 0) {
      return binding.revision
    }
    if (delta.kind === 'rows') {
      try {
        const revision =
          binding.provider === 'google'
            ? await pushGoogleRows(binding.file, delta.writes, delta.clears)
            : await pushMicrosoftRows(binding.file, delta.writes, delta.clears)
        if (sessionId) setCloudBaseline(sessionId, snapshotSheetMatrices(workbook))
        return revision
      } catch (err) {
        const message = err instanceof Error ? err.message : ''
        const canRewrite =
          /grid|tried writing|exceeds grid|Unable to parse range/i.test(message) ||
          (binding.provider === 'microsoft' && message.includes('chưa có trên cloud'))
        if (!canRewrite) throw err
      }
    }
    const revision =
      binding.provider === 'google'
        ? await pushGoogleSheet(binding.file, workbook)
        : await pushMicrosoftWorkbook(binding.file, workbook)
    if (sessionId) setCloudBaseline(sessionId, snapshotSheetMatrices(workbook))
    return revision
  }
  onProgress?.('Đang dựng file .xlsx...')
  const built = await buildXlsxBlob(workbook, onProgress)
  const blob = await preserveFormSheet(binding.provider, binding.file, built)
  if (binding.provider === 'google') return uploadGoogleXlsx(binding.file.id, blob)
  return uploadMicrosoftXlsx(binding.file, blob)
}

export async function publishLocalWorkbook(
  provider: CloudProviderId,
  workbook: WorkbookData,
  onProgress?: (message: string) => void,
): Promise<CloudBinding> {
  onProgress?.('Đang dựng file .xlsx...')
  const blob = await buildXlsxBlob(workbook, onProgress)
  onProgress?.('Đang tải lên cloud...')
  const file = provider === 'google'
    ? await createGoogleXlsx(workbook.meta.fileName, blob)
    : await createMicrosoftXlsx(workbook.meta.fileName, blob)
  const revision = provider === 'google' ? await googleRevision(file.id) : await microsoftRevision(file)
  return { provider, mode: 'file', file, revision }
}
