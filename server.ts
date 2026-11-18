import express, { Request, Response } from 'express';
import multer from 'multer';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  handleHopDongGao,
  handleHopDongNguyenTacGao,
  handleHopDongXeTai,
} from './src/contractHandlers.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || '0.0.0.0';

app.use((_req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  next();
});

app.get('/api/health', (_req: Request, res: Response) => {
  res.json({ message: 'Hệ thống Backend đang chạy mượt mà trên Node.js!' });
});

app.post(
  '/api/hop-dong-gao',
  upload.fields([
    { name: 'excel_file', maxCount: 1 },
    { name: 'word_template', maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const files = req.files as Record<string, Express.Multer.File[]> | undefined;
      const excelFile = files?.['excel_file']?.[0];
      const wordFile = files?.['word_template']?.[0];

      if (!excelFile || !wordFile) {
        res.status(400).json({ error: 'Thiếu file Excel hoặc file mẫu Word.' });
        return;
      }

      const zipBuffer = await handleHopDongGao(excelFile.buffer, wordFile.buffer);
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename=1_KetQua_HopDongGao.zip');
      res.send(zipBuffer);
    } catch (err) {
      console.error('Error in /api/hop-dong-gao:', err);
      res.status(500).json({ error: err instanceof Error ? err.message : 'Lỗi xử lý hợp đồng gạo.' });
    }
  }
);

app.post(
  '/api/hop-dong-xe-tai',
  upload.fields([
    { name: 'excel_file', maxCount: 1 },
    { name: 'word_template', maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const files = req.files as Record<string, Express.Multer.File[]> | undefined;
      const excelFile = files?.['excel_file']?.[0];
      const wordFile = files?.['word_template']?.[0];

      if (!excelFile || !wordFile) {
        res.status(400).json({ error: 'Thiếu file Excel hoặc file mẫu Word.' });
        return;
      }

      const result = await handleHopDongXeTai(excelFile.buffer, wordFile.buffer);
      if ('error' in result) {
        res.status(400).json(result);
        return;
      }

      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename=2_KetQua_HopDongXeTai.zip');
      res.send(result);
    } catch (err) {
      console.error('Error in /api/hop-dong-xe-tai:', err);
      res.status(500).json({ error: err instanceof Error ? err.message : 'Lỗi xử lý hợp đồng xe tải.' });
    }
  }
);

app.post(
  '/api/hop-dong-nguyen-tac-gao',
  upload.fields([
    { name: 'excel_file', maxCount: 1 },
    { name: 'hdnt_template', maxCount: 1 },
    { name: 'bbgn_template', maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const files = req.files as Record<string, Express.Multer.File[]> | undefined;
      const excelFile = files?.['excel_file']?.[0];
      const hdntFile = files?.['hdnt_template']?.[0];
      const bbgnFile = files?.['bbgn_template']?.[0];

      if (!excelFile || !hdntFile || !bbgnFile) {
        res.status(400).json({ error: 'Thiếu file Excel hoặc file mẫu HĐNT/BBGN.' });
        return;
      }

      const zipBuffer = await handleHopDongNguyenTacGao(
        excelFile.buffer,
        hdntFile.buffer,
        bbgnFile.buffer
      );
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename=3_KetQua_HDNT_Gao.zip');
      res.send(zipBuffer);
    } catch (err) {
      console.error('Error in /api/hop-dong-nguyen-tac-gao:', err);
      res.status(500).json({ error: err instanceof Error ? err.message : 'Lỗi xử lý HĐNT gạo.' });
    }
  }
);

// Serve static directories
const aMainDir = path.join(__dirname, 'a-main');
const toolMainDir = path.join(__dirname, 'tool-main');

app.get('/favicon.ico', (_req: Request, res: Response) => {
  res.status(204).end();
});

app.use((req: Request, res: Response, next) => {
  if (req.path === '/a-main' || req.path === '/a') {
    res.redirect(301, `${req.path}/`);
    return;
  }
  next();
});

app.use('/a-main', express.static(aMainDir));
app.use('/a', express.static(aMainDir));
app.use('/tool-main', express.static(toolMainDir));
app.use('/', express.static(toolMainDir));

app.listen(PORT, HOST, () => {
  console.log(`Server running at http://${HOST}:${PORT}`);
});
