import { Router, Request, Response } from 'express';
import { groupData } from '../services/excelParser';
import { generatePPT } from '../services/pptGenerator';
import { GenerateRequest } from '../types';

const router = Router();

router.post('/', async (req: Request, res: Response) => {
  try {
    const body = req.body as GenerateRequest;
    const { rows, columnMapping, selectedPMs } = body;

    if (!rows || !Array.isArray(rows)) {
      return res.status(400).json({ error: 'rows array is required.' });
    }
    if (!columnMapping || !columnMapping.pmName || !columnMapping.projectName) {
      return res.status(400).json({ error: 'columnMapping is required with pmName and projectName.' });
    }

    const pmGroups = groupData(rows, columnMapping, selectedPMs ?? []);

    if (pmGroups.length === 0) {
      return res.status(400).json({ error: 'No data found for the selected PMs.' });
    }

    const buffer = await generatePPT(pmGroups);

    const filename = `VOIS_Portfolio_${new Date().toISOString().slice(0, 10)}.pptx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Length', buffer.length);
    return res.send(buffer);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Failed to generate PowerPoint.';
    console.error('[generate]', err);
    return res.status(500).json({ error: message });
  }
});

export default router;
