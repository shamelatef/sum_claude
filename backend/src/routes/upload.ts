import { Router, Request, Response } from 'express';
import multer from 'multer';
import { parseExcelBuffer, extractUniquePMs } from '../services/excelParser';

const router = Router();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 MB
  fileFilter: (_req, file, cb) => {
    const allowed = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'application/octet-stream',
    ];
    const extOk = /\.(xlsx|xls)$/i.test(file.originalname);
    if (allowed.includes(file.mimetype) || extOk) {
      cb(null, true);
    } else {
      cb(new Error('Only .xlsx and .xls files are allowed.'));
    }
  },
});

router.post('/', upload.single('file'), (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const parsed = parseExcelBuffer(req.file.buffer);

    return res.json({
      headers: parsed.headers,
      rows: parsed.rows,
      totalRows: parsed.rows.length,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Failed to parse file.';
    return res.status(500).json({ error: message });
  }
});

router.post('/pms', upload.single('file'), (req: Request, res: Response) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });
    const pmColumn = req.body.pmColumn as string;
    if (!pmColumn) return res.status(400).json({ error: 'pmColumn is required.' });

    const parsed = parseExcelBuffer(req.file.buffer);
    const pms = extractUniquePMs(parsed.rows, pmColumn);
    return res.json({ pms });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Failed to extract PMs.';
    return res.status(500).json({ error: message });
  }
});

export default router;
