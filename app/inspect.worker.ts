import { inspectWorkbook } from '../core/workbook.mjs';
import { ZIP_LIMITS } from '../core/zip.mjs';

self.onmessage = async (
  event: MessageEvent<{ file?: File; bytes?: Uint8Array }>,
) => {
  try {
    const { file } = event.data;
    if (
      file &&
      (!/\.xlsx$/i.test(file.name) || file.size > ZIP_LIMITS.archiveBytes)
    )
      throw new Error(
        'Choose an unencrypted .xlsx file no larger than 20 MiB.',
      );
    self.postMessage({ status: 'Reading workbook structure…' });
    const bytes = file
      ? new Uint8Array(await file.arrayBuffer())
      : event.data.bytes;
    if (!bytes) throw new Error('No workbook was selected.');
    self.postMessage({ report: inspectWorkbook(bytes) });
  } catch (error) {
    self.postMessage({
      error:
        error instanceof Error
          ? error.message
          : 'This workbook could not be inspected.',
    });
  }
};
