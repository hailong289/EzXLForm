export const DISCORD_FILE_LIMIT = 25 * 1024 * 1024
export const SUPPORT_FILE_MAIL = 'longdh2.dev@gmail.com'

export function supportPayload(input: {
  title: string
  description: string
  email: string
  fileName: string | null
  fileBytes: number
}): { payload: string; attach: boolean } {
  const attach = input.fileBytes > 0 && input.fileBytes <= DISCORD_FILE_LIMIT
  const fileNote = !input.fileName
    ? 'Không có file local'
    : attach
      ? input.fileName
      : `${input.fileName} quá 25 MB. Yêu cầu gửi file tới ${SUPPORT_FILE_MAIL}`
  return {
    attach,
    payload: JSON.stringify({
      embeds: [
        {
          title: input.title.slice(0, 256),
          description: input.description.slice(0, 4000),
          fields: [
            { name: 'Email', value: input.email.slice(0, 256) },
            { name: 'File', value: fileNote.slice(0, 256) },
          ],
        },
      ],
    }),
  }
}
