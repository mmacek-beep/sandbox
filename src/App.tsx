/**
 * Evala Edu - Účetní Sandbox & Trenažér oběhu dokladů pro OA
 * Complete React Application with Microsoft Word (.docx) Integration
 * @license Apache-2.0
 */

import React, { useEffect, useRef, useState } from 'react';
import {
  FileText,
  Download,
  Upload,
  CheckCircle2,
  BookOpen,
  Award,
  Layers,
  Sparkles,
  Info,
  X,
} from 'lucide-react';
import {
  generateDocumentDocx,
  generateTeacherProtocolDocx,
  generateWorksheetDocx,
  generateQuizDocx,
  generateJournalDocx,
  downloadBlob,
} from './utils/docxExport';
import { parseDocxFile, ParsedWordDocument } from './utils/docxImport';

export default function App() {
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const [isWordBannerOpen, setIsWordBannerOpen] = useState(true);
  const [parsedDoc, setParsedDoc] = useState<ParsedWordDocument | null>(null);
  const [toastMessage, setToastMessage] = useState<{ title: string; desc: string; type: 'success' | 'info' | 'error' } | null>(null);

  const showToast = (title: string, desc: string, type: 'success' | 'info' | 'error' = 'success') => {
    setToastMessage({ title, desc, type });
    setTimeout(() => {
      setToastMessage(null);
    }, 4500);
  };

  // Expose docxExport & docxImport to window so iframe can access it directly
  useEffect(() => {
    (window as any).docxExport = {
      generateDocumentDocx,
      generateTeacherProtocolDocx,
      generateWorksheetDocx,
      generateQuizDocx,
      generateJournalDocx,
      downloadBlob,
    };
    (window as any).docxImport = {
      parseDocxFile,
    };

    // Listen for postMessage from the iframe for Word export requests
    const handleMessage = async (event: MessageEvent) => {
      const data = event.data;
      if (!data || typeof data !== 'object') return;

      try {
        if (data.type === 'EXPORT_DOCX_DOCUMENT') {
          showToast('Generuji Word (.docx)', `Připravuji dokument ${data.docType || 'VF'}...`, 'info');
          const blob = await generateDocumentDocx(data.docType || 'VF', data.state);
          downloadBlob(blob, `Doklad_${data.docType || 'VF'}_20260001.docx`);
          showToast('Word stažen!', `Doklad ${data.docType || 'VF'} byl úspěšně uložen jako .docx.`, 'success');
        } else if (data.type === 'EXPORT_DOCX_PROTOCOL') {
          showToast('Generuji Word (.docx)', 'Připravuji učitelský protokol...', 'info');
          const blob = await generateTeacherProtocolDocx(data.state);
          const student = (data.state?.studentName || 'Student').replace(/[^a-zA-Z0-9]/g, '_');
          downloadBlob(blob, `Vyhodnoceni_Zkousky_${student}.docx`);
          showToast('Protokol stažen!', 'Učitelský protokol byl úspěšně uložen jako .docx.', 'success');
        } else if (data.type === 'EXPORT_DOCX_WORKSHEET') {
          showToast('Generuji Word (.docx)', 'Připravuji pracovní list pro studenty...', 'info');
          const blob = await generateWorksheetDocx(data.state);
          downloadBlob(blob, 'Pracovni_list_ucetnictvi_OA.docx');
          showToast('Pracovní list stažen!', 'Pracovní list byl úspěšně uložen jako .docx.', 'success');
        } else if (data.type === 'EXPORT_DOCX_QUIZ') {
          showToast('Generuji Word (.docx)', 'Připravuji písemný test...', 'info');
          const blob = await generateQuizDocx();
          downloadBlob(blob, 'Pisemny_test_DPH_ucetnictvi_OA.docx');
          showToast('Test stažen!', 'Písemný test byl úspěšně uložen jako .docx.', 'success');
        } else if (data.type === 'EXPORT_DOCX_JOURNAL') {
          showToast('Generuji Word (.docx)', 'Připravuji účetní deník...', 'info');
          const blob = await generateJournalDocx(data.journal || [], data.title || 'Účetní deník');
          downloadBlob(blob, 'Ucetni_denik_Evala_Edu.docx');
          showToast('Deník stažen!', 'Účetní deník byl úspěšně uložen jako .docx.', 'success');
        }
      } catch (err: any) {
        console.error('Word export error:', err);
        showToast('Chyba Wordu', err?.message || 'Chyba při exportu do .docx', 'error');
      }
    };

    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  // Pass docxExport into iframe when iframe loads
  const handleIframeLoad = () => {
    try {
      if (iframeRef.current && iframeRef.current.contentWindow) {
        (iframeRef.current.contentWindow as any).docxExport = (window as any).docxExport;
        (iframeRef.current.contentWindow as any).docxImport = (window as any).docxImport;
      }
    } catch (e) {
      // Ignored if cross-origin, but here it's same-origin
    }
  };

  // Helper to get state from iframe
  const getIframeState = () => {
    try {
      if (iframeRef.current && iframeRef.current.contentWindow) {
        return (iframeRef.current.contentWindow as any).state || {};
      }
    } catch (e) {}
    return {
      caseTitle: 'Dodávka a implementace ERP systému Evala',
      customBaseAmount: 100000,
      vatRate: 21,
      advancePercent: 100,
      role: 'supplier',
    };
  };

  const handleDownloadWorksheet = async () => {
    try {
      showToast('Word export', 'Připravuji pracovní list pro studenty...', 'info');
      const state = getIframeState();
      const blob = await generateWorksheetDocx(state);
      downloadBlob(blob, 'Pracovni_list_ucetnictvi_OA.docx');
      showToast('Hotovo!', 'Pracovní list byl stažen jako .docx', 'success');
    } catch (err: any) {
      showToast('Chyba', err.message, 'error');
    }
  };

  const handleDownloadQuiz = async () => {
    try {
      showToast('Word export', 'Připravuji test z DPH a účtování...', 'info');
      const blob = await generateQuizDocx();
      downloadBlob(blob, 'Pisemny_test_DPH_ucetnictvi_OA.docx');
      showToast('Hotovo!', 'Písemný test byl stažen jako .docx', 'success');
    } catch (err: any) {
      showToast('Chyba', err.message, 'error');
    }
  };

  const handleDownloadProtocol = async () => {
    try {
      showToast('Word export', 'Připravuji učitelský protokol...', 'info');
      const state = getIframeState();
      const blob = await generateTeacherProtocolDocx(state);
      const student = (state.studentName || 'Student').replace(/[^a-zA-Z0-9]/g, '_');
      downloadBlob(blob, `Vyhodnoceni_Zkousky_${student}.docx`);
      showToast('Hotovo!', 'Učitelský protokol byl stažen jako .docx', 'success');
    } catch (err: any) {
      showToast('Chyba', err.message, 'error');
    }
  };

  const handleDownloadDocument = async (type: string = 'VF') => {
    try {
      showToast('Word export', `Připravuji doklad ${type}...`, 'info');
      const state = getIframeState();
      const blob = await generateDocumentDocx(type, state);
      downloadBlob(blob, `Doklad_${type}_20260001.docx`);
      showToast('Hotovo!', `Doklad ${type} byl stažen jako .docx`, 'success');
    } catch (err: any) {
      showToast('Chyba', err.message, 'error');
    }
  };

  const handleWordFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      showToast('Čtení Wordu (.docx)', `Zpracovávám dokument ${file.name}...`, 'info');
      const parsed = await parseDocxFile(file);
      setParsedDoc(parsed);

      // If iframe exists, optionally load extracted data into state
      try {
        const win = iframeRef.current?.contentWindow as any;
        if (win && win.state) {
          if (parsed.foundIco) {
            win.state.partner = win.state.partner || {};
            win.state.partner.ico = parsed.foundIco;
            if (parsed.foundDic) win.state.partner.dic = parsed.foundDic;
          }
          if (parsed.foundAmount) {
            win.state.customBaseAmount = parsed.foundAmount;
          }
          if (parsed.detectedTitle) {
            win.state.caseTitle = parsed.detectedTitle;
          }
          if (win.renderApp) {
            win.renderApp();
          }
        }
      } catch (e) {}

      showToast('Word načten!', `Dokument "${file.name}" byl úspěšně digitalizován.`, 'success');
    } catch (err: any) {
      showToast('Chyba čtení', err.message || 'Nepodařilo se přečíst Word soubor', 'error');
    }
  };

  return (
    <div className="flex flex-col h-screen w-screen bg-slate-950 font-sans text-slate-100 overflow-hidden">
      {/* WORD INTEGRATION ANNOUNCEMENT & QUICK TOOLBAR */}
      {isWordBannerOpen && (
        <div className="bg-gradient-to-r from-blue-950 via-slate-900 to-indigo-950 border-b border-blue-800/50 px-4 py-2 flex items-center justify-between flex-wrap gap-2 text-xs shrink-0 z-30 shadow-md">
          <div className="flex items-center space-x-2">
            <span className="flex h-2 w-2 relative">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-blue-500"></span>
            </span>
            <div className="flex items-center gap-1.5 font-bold text-white">
              <FileText className="w-4 h-4 text-blue-400" />
              <span>Plná podpora Microsoft Word (.docx) aktivována:</span>
            </div>
            <span className="text-slate-300 hidden md:inline">
              Všechny účetní doklady, učitelské protokoly, zadání písemek a cvičení lze exportovat i importovat ve Wordu!
            </span>
          </div>

          <div className="flex items-center gap-1.5 flex-wrap">
            <button
              onClick={handleDownloadWorksheet}
              className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-bold flex items-center gap-1 shadow-sm transition-all"
              title="Stáhnout pracovní list cvičení pro studenty do Wordu (.docx)"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Pracovní list (.docx)</span>
            </button>

            <button
              onClick={handleDownloadQuiz}
              className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg font-bold flex items-center gap-1 shadow-sm transition-all"
              title="Stáhnout test ze znalostí DPH a účtování do Wordu (.docx)"
            >
              <Award className="w-3.5 h-3.5" />
              <span>Test / Písemka (.docx)</span>
            </button>

            <button
              onClick={handleDownloadProtocol}
              className="px-2.5 py-1 bg-teal-600 hover:bg-teal-500 text-white rounded-lg font-bold flex items-center gap-1 shadow-sm transition-all"
              title="Stáhnout učitelský protokol hodnocení zkoušky do Wordu (.docx)"
            >
              <BookOpen className="w-3.5 h-3.5" />
              <span>Protokol (.docx)</span>
            </button>

            <button
              onClick={() => handleDownloadDocument('VF')}
              className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-blue-300 border border-slate-700 rounded-lg font-bold flex items-center gap-1 shadow-sm transition-all"
              title="Stáhnout vyúčtovací fakturu do Wordu (.docx)"
            >
              <FileText className="w-3.5 h-3.5" />
              <span>Faktura (.docx)</span>
            </button>

            {/* Word Upload Button */}
            <label className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-amber-300 border border-amber-500/40 rounded-lg font-bold flex items-center gap-1 shadow-sm cursor-pointer transition-all">
              <Upload className="w-3.5 h-3.5" />
              <span>Nahrát Word (.docx)</span>
              <input
                type="file"
                accept=".docx,.doc"
                onChange={handleWordFileUpload}
                className="hidden"
              />
            </label>

            <button
              onClick={() => setIsWordBannerOpen(false)}
              className="p-1 text-slate-400 hover:text-white rounded hover:bg-slate-800"
              title="Skrýt lištu"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* PARSED WORD FILE MODAL POPUP (IF UPLOADED) */}
      {parsedDoc && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-lg w-full p-6 text-white space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2 text-blue-400 font-bold text-sm">
                <FileText className="w-5 h-5" />
                <span>Nalezená data z Word dokumentu (.docx)</span>
              </div>
              <button
                onClick={() => setParsedDoc(null)}
                className="text-slate-400 hover:text-white"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-2 text-xs font-mono">
              <div className="p-2.5 bg-slate-950 rounded-lg border border-slate-800">
                <span className="text-slate-400 block text-[10px] font-sans">Detekovaný název / Předmět:</span>
                <strong className="text-teal-300">{parsedDoc.detectedTitle || 'Dokument'}</strong>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div className="p-2 bg-slate-950 rounded-lg border border-slate-800">
                  <span className="text-slate-400 block text-[10px] font-sans">Nalezené IČO:</span>
                  <strong className="text-white">{parsedDoc.foundIco || 'Nenalezeno'}</strong>
                </div>
                <div className="p-2 bg-slate-950 rounded-lg border border-slate-800">
                  <span className="text-slate-400 block text-[10px] font-sans">Nalezené DIČ:</span>
                  <strong className="text-white">{parsedDoc.foundDic || 'Nenalezeno'}</strong>
                </div>
              </div>
              {parsedDoc.foundAmount && (
                <div className="p-2 bg-slate-950 rounded-lg border border-slate-800">
                  <span className="text-slate-400 block text-[10px] font-sans">Nalezená částka:</span>
                  <strong className="text-emerald-400 text-sm">{parsedDoc.foundAmount.toLocaleString('cs-CZ')} Kč</strong>
                </div>
              )}
              <div className="p-2.5 bg-slate-950 rounded-lg border border-slate-800 max-h-36 overflow-y-auto text-[11px] text-slate-300 font-sans leading-relaxed">
                <span className="text-slate-500 block text-[10px] uppercase font-bold mb-1">Náhled extrahovaného textu:</span>
                {parsedDoc.rawText.slice(0, 400)}...
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setParsedDoc(null)}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg text-xs"
              >
                Rozumím, zavřít
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MAIN APPLICATION CONTAINER (EMBEDS FULL EVALA EDU WORKFLOW) */}
      <div className="flex-1 w-full h-full relative overflow-hidden">
        <iframe
          ref={iframeRef}
          src="/assets/aistudio/evala_trenazer_oa_f9b36e8f.html"
          onLoad={handleIframeLoad}
          className="w-full h-full border-0"
          title="Evala Edu Účetní Trenažér pro OA"
        />
      </div>

      {/* TOAST NOTIFICATION */}
      {toastMessage && (
        <div className="fixed bottom-4 right-4 z-50 pointer-events-none">
          <div
            className={`p-4 rounded-xl shadow-2xl border text-xs flex items-start gap-3 pointer-events-auto max-w-sm ${
              toastMessage.type === 'success'
                ? 'bg-emerald-950 border-emerald-500/80 text-emerald-100'
                : toastMessage.type === 'error'
                ? 'bg-rose-950 border-rose-500/80 text-rose-100'
                : 'bg-blue-950 border-blue-500/80 text-blue-100'
            }`}
          >
            <div className="mt-0.5">
              {toastMessage.type === 'success' ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
              ) : (
                <Info className="w-4 h-4 text-blue-400" />
              )}
            </div>
            <div className="flex-1">
              <strong className="block font-bold text-sm mb-0.5">{toastMessage.title}</strong>
              <p className="opacity-90">{toastMessage.desc}</p>
            </div>
            <button
              onClick={() => setToastMessage(null)}
              className="text-slate-400 hover:text-white"
            >
              ✕
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
