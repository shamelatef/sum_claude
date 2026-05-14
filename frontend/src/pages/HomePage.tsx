import React, { useState, useMemo } from 'react';
import { DropZone } from '../components/DropZone';
import { ExcelPreview } from '../components/ExcelPreview';
import { ColumnMapper } from '../components/ColumnMapper';
import { PMSelector } from '../components/PMSelector';
import { StepIndicator } from '../components/StepIndicator';
import { useFileUpload } from '../hooks/useFileUpload';
import { useColumnMapping } from '../hooks/useColumnMapping';
import { generatePPT } from '../services/api';
import { downloadBlob, formatDate, cleanPMName } from '../utils/helpers';
import { Step, RawRow } from '../types';

export const HomePage: React.FC = () => {
  const { file, headers, rows, totalRows, loading: uploading, error: uploadError, handleFile, reset } = useFileUpload();
  const { mapping, setField, isComplete } = useColumnMapping(headers);

  const [step, setStep] = useState<Step>('upload');
  const [selectedPMs, setSelectedPMs] = useState<string[]>([]);
  const [generating, setGenerating] = useState(false);
  const [genError, setGenError] = useState<string | null>(null);
  const [downloadReady, setDownloadReady] = useState(false);

  const uniquePMs = useMemo<string[]>(() => {
    if (!mapping.pmName || rows.length === 0) return [];
    const names = new Set<string>();
    for (const row of rows) {
      const raw = row[mapping.pmName];
      if (raw !== undefined && raw !== null && raw !== '') {
        const cleaned = cleanPMName(String(raw));
        if (cleaned) names.add(cleaned);
      }
    }
    return Array.from(names).sort((a, b) => a.localeCompare(b));
  }, [rows, mapping.pmName]);

  const handleFileSelected = async (f: File) => {
    reset();
    setStep('upload');
    setSelectedPMs([]);
    setDownloadReady(false);
    await handleFile(f);
    setStep('map');
  };

  const handleMappingConfirm = () => {
    if (!isComplete) return;
    setSelectedPMs([...uniquePMs]);
    setStep('select');
  };

  const handleGenerate = async () => {
    if (selectedPMs.length === 0) return;
    setGenerating(true);
    setGenError(null);
    setDownloadReady(false);
    try {
      const blob = await generatePPT(rows as RawRow[], mapping, selectedPMs);
      downloadBlob(blob, `VOIS_Portfolio_${formatDate()}.pptx`);
      setDownloadReady(true);
      setStep('generate');
    } catch (err: unknown) {
      setGenError(err instanceof Error ? err.message : 'Generation failed. Please try again.');
    } finally {
      setGenerating(false);
    }
  };

  const startOver = () => {
    reset();
    setStep('upload');
    setSelectedPMs([]);
    setDownloadReady(false);
    setGenError(null);
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ background: 'linear-gradient(135deg, #0a0a0a 0%, #1a1a1a 50%, #2a0000 100%)' }}>
      {/* Top nav */}
      <header className="flex items-center justify-between px-8 py-4">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-white/10 backdrop-blur flex items-center justify-center">
            <span className="text-white font-black text-sm tracking-tight">V</span>
          </div>
          <div>
            <p className="text-white font-bold text-sm leading-tight">VOIS</p>
            <p className="text-vois-300 text-[10px] leading-tight">PMO Portal</p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <span className="px-3 py-1 bg-white/10 backdrop-blur text-vois-200 rounded-full text-xs font-medium">
            Portfolio PPT Generator
          </span>
        </div>
      </header>

      {/* Hero */}
      <div className="text-center pt-6 pb-8 px-4">
        <h1 className="text-3xl font-extrabold text-white tracking-tight">
          Portfolio Presentation Generator
        </h1>
        <p className="text-vois-300 mt-2 text-sm max-w-lg mx-auto">
          Upload your Excel data, map columns, select PMs, and export a polished VOIS-branded PowerPoint in seconds.
        </p>
      </div>

      {/* Main card */}
      <div className="flex-1 flex items-start justify-center px-4 pb-12">
        <div className="w-full max-w-3xl">
          {/* Step indicator */}
          <div className="mb-8 px-2">
            <StepIndicator current={step} />
          </div>

          <div className="bg-white rounded-3xl shadow-2xl shadow-vois-900/40 overflow-hidden">

            {/* ── STEP: Upload ── */}
            {step === 'upload' && (
              <div className="p-8">
                <SectionHeader
                  step={1}
                  title="Upload Excel File"
                  subtitle="Start by uploading your project data spreadsheet"
                />
                <DropZone onFile={handleFileSelected} loading={uploading} />
                {uploadError && (
                  <ErrorBanner message={uploadError} />
                )}
              </div>
            )}

            {/* ── STEP: Map ── */}
            {step === 'map' && (
              <div className="p-8 space-y-6">
                <div className="flex items-center justify-between">
                  <SectionHeader
                    step={2}
                    title="Map Excel Columns"
                    subtitle="Tell the system which column maps to each required field"
                  />
                  <button onClick={startOver} className="text-xs text-gray-400 hover:text-vois-600 transition-colors">
                    ← Start over
                  </button>
                </div>

                {file && (
                  <div className="flex items-center gap-2 px-3 py-2 bg-green-50 border border-green-200 rounded-lg">
                    <svg className="w-4 h-4 text-green-500 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    <span className="text-xs text-green-700 font-medium">{file.name}</span>
                    <span className="text-xs text-green-500">· {totalRows} rows loaded</span>
                  </div>
                )}

                <ExcelPreview headers={headers} rows={rows} totalRows={totalRows} />

                <ColumnMapper headers={headers} mapping={mapping} onFieldChange={setField} />

                <div className="flex justify-end pt-2">
                  <button
                    onClick={handleMappingConfirm}
                    disabled={!isComplete}
                    className={`
                      px-6 py-3 rounded-xl font-semibold text-sm transition-all duration-200
                      ${isComplete
                        ? 'bg-vois-700 text-white hover:bg-vois-800 shadow-md hover:shadow-lg hover:shadow-vois-200'
                        : 'bg-gray-100 text-gray-400 cursor-not-allowed'
                      }
                    `}
                  >
                    Continue to PM Selection →
                  </button>
                </div>
              </div>
            )}

            {/* ── STEP: Select PMs ── */}
            {step === 'select' && (
              <div className="p-8 space-y-6">
                <div className="flex items-center justify-between">
                  <SectionHeader
                    step={3}
                    title="Select Project Managers"
                    subtitle="Choose which PMs to include in the generated slides"
                  />
                  <button onClick={() => setStep('map')} className="text-xs text-gray-400 hover:text-vois-600 transition-colors">
                    ← Back
                  </button>
                </div>

                <PMSelector pms={uniquePMs} selected={selectedPMs} onChange={setSelectedPMs} />

                {selectedPMs.length > 0 && (
                  <div className="bg-vois-50 border border-vois-200 rounded-xl p-4">
                    <div className="flex items-center gap-2 text-vois-800">
                      <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd" />
                      </svg>
                      <span className="text-sm font-semibold">Slide Preview</span>
                    </div>
                    <p className="text-xs text-vois-600 mt-1">
                      Will generate: 1 summary slide + {selectedPMs.length} PM slide{selectedPMs.length !== 1 ? 's' : ''} (+ continuation slides as needed)
                    </p>
                  </div>
                )}

                <div className="flex justify-between items-center pt-2">
                  <button
                    onClick={() => setStep('map')}
                    className="px-5 py-2.5 rounded-xl text-sm font-medium text-gray-600 hover:bg-gray-100 transition-colors"
                  >
                    ← Back
                  </button>
                  <button
                    onClick={handleGenerate}
                    disabled={selectedPMs.length === 0 || generating}
                    className={`
                      flex items-center gap-2 px-7 py-3 rounded-xl font-bold text-sm transition-all duration-200
                      ${selectedPMs.length > 0 && !generating
                        ? 'bg-gradient-to-r from-vois-700 to-vois-600 text-white shadow-md hover:shadow-xl hover:shadow-vois-300 hover:scale-[1.02]'
                        : 'bg-gray-100 text-gray-400 cursor-not-allowed'
                      }
                    `}
                  >
                    {generating ? (
                      <>
                        <svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
                        </svg>
                        Generating…
                      </>
                    ) : (
                      <>
                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                            d="M12 10v6m0 0l-3-3m3 3l3-3M3 17V7a2 2 0 012-2h6l2 2h6a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2z" />
                        </svg>
                        Generate PowerPoint
                      </>
                    )}
                  </button>
                </div>

                {genError && <ErrorBanner message={genError} />}
              </div>
            )}

            {/* ── STEP: Done ── */}
            {step === 'generate' && downloadReady && (
              <div className="p-12 text-center space-y-6">
                <div className="w-20 h-20 rounded-full bg-green-100 flex items-center justify-center mx-auto">
                  <svg className="w-10 h-10 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                  </svg>
                </div>
                <div>
                  <h2 className="text-2xl font-bold text-gray-800">Presentation Ready!</h2>
                  <p className="text-gray-500 text-sm mt-1">
                    Your VOIS-branded PowerPoint has been downloaded automatically.
                  </p>
                </div>
                <div className="flex justify-center gap-4">
                  <button
                    onClick={handleGenerate}
                    className="flex items-center gap-2 px-5 py-2.5 rounded-xl border border-vois-300 text-vois-700 font-semibold text-sm hover:bg-vois-50 transition-colors"
                  >
                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                    </svg>
                    Download Again
                  </button>
                  <button
                    onClick={startOver}
                    className="px-5 py-2.5 rounded-xl bg-vois-700 text-white font-semibold text-sm hover:bg-vois-800 transition-colors"
                  >
                    Start New Report
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* Footer */}
          <p className="text-center text-vois-400 text-xs mt-6">
            VOIS Project Management Office · Internal Tool
          </p>
        </div>
      </div>
    </div>
  );
};

// ─── small sub-components ──────────────────────────────────────────────────────

const SectionHeader: React.FC<{ step: number; title: string; subtitle: string }> = ({ step, title, subtitle }) => (
  <div className="flex items-start gap-3 mb-6">
    <div className="w-8 h-8 rounded-full bg-vois-700 text-white flex items-center justify-center text-sm font-bold flex-shrink-0 mt-0.5">
      {step}
    </div>
    <div>
      <h2 className="text-lg font-bold text-gray-900">{title}</h2>
      <p className="text-sm text-gray-500">{subtitle}</p>
    </div>
  </div>
);

const ErrorBanner: React.FC<{ message: string }> = ({ message }) => (
  <div className="flex items-start gap-2 p-3 bg-red-50 border border-red-200 rounded-xl text-sm text-red-700 mt-4">
    <svg className="w-4 h-4 flex-shrink-0 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
      <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
    </svg>
    {message}
  </div>
);
