import { useState, useCallback } from 'react';
import { parseExcelFile } from '../services/api';
import { RawRow } from '../types';

interface UseFileUploadReturn {
  file: File | null;
  headers: string[];
  rows: RawRow[];
  totalRows: number;
  loading: boolean;
  error: string | null;
  handleFile: (f: File) => Promise<void>;
  reset: () => void;
}

export function useFileUpload(): UseFileUploadReturn {
  const [file, setFile] = useState<File | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<RawRow[]>([]);
  const [totalRows, setTotalRows] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFile = useCallback(async (f: File) => {
    setError(null);
    setLoading(true);
    try {
      const data = await parseExcelFile(f);
      setFile(f);
      setHeaders(data.headers);
      setRows(data.rows);
      setTotalRows(data.totalRows);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Failed to read file.');
    } finally {
      setLoading(false);
    }
  }, []);

  const reset = useCallback(() => {
    setFile(null);
    setHeaders([]);
    setRows([]);
    setTotalRows(0);
    setError(null);
  }, []);

  return { file, headers, rows, totalRows, loading, error, handleFile, reset };
}
