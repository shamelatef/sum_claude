import { RawRow, ColumnMapping } from '../types';
import { groupData } from './excelParser';
import { generatePPT as generatePPTFromGroups } from './pptGenerator';

export { parseExcelFile } from './excelParser';

export async function generatePPT(
  rows: RawRow[],
  columnMapping: ColumnMapping,
  selectedPMs: string[]
): Promise<Blob> {
  const pmGroups = groupData(rows, columnMapping, selectedPMs);
  return generatePPTFromGroups(pmGroups);
}
