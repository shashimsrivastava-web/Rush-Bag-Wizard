'use client';

import React, { useEffect, useState, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { X, Camera, Sparkles, Volume2, Info, CheckCircle2, RotateCcw } from 'lucide-react';
import { getActiveScanner, BarcodeScannerProvider } from '../lib/scanner';

interface ScannerModalProps {
  isOpen: boolean;
  onClose: () => void;
  onScanSuccess: (barcode: string) => void;
  title?: string;
  isContinuous?: boolean;
  continuousCount?: number;
  onFinishContinuous?: () => void;
}

export default function ScannerModal({
  isOpen,
  onClose,
  onScanSuccess,
  title = 'Scan Baggage Barcode',
  isContinuous = false,
  continuousCount = 0,
  onFinishContinuous
}: ScannerModalProps) {
  const [error, setError] = useState<string | null>(null);
  const [successFlash, setSuccessFlash] = useState<string | null>(null);
  const [simulatedInput, setSimulatedInput] = useState('');
  const [providerName] = useState(() => getActiveScanner().name);
  const [cameraActive, setCameraActive] = useState(false);
  const [isInitializing, setIsInitializing] = useState(false);
  const scannerRef = useRef<BarcodeScannerProvider | null>(null);

  const handleScanMatch = React.useCallback((barcode: string) => {
    const cleanBarcode = barcode.trim().toUpperCase();
    if (!cleanBarcode) return;

    setSuccessFlash(cleanBarcode);
    onScanSuccess(cleanBarcode);

    // Flash animation trigger
    setTimeout(() => {
      setSuccessFlash(null);
      if (!isContinuous) {
        onClose();
      }
    }, 1200);
  }, [onScanSuccess, isContinuous, onClose]);

  useEffect(() => {
    if (!isOpen) {
      if (scannerRef.current) {
        scannerRef.current.stopScanning().catch(console.error);
        setCameraActive(false);
      }
      return;
    }

    // Initialize scan
    const provider = getActiveScanner();
    scannerRef.current = provider;

    const startCamera = async () => {
      setIsInitializing(true);
      setError(null);
      try {
        await provider.init();
        await provider.startScanning(
          'scanner-reader-container',
          (text) => {
            // Success
            handleScanMatch(text);
          },
          (err) => {
            // Scan warning or frame miss (usually safe to ignore unless critical)
          }
        );
        setCameraActive(true);
      } catch (err: any) {
        console.error('Camera initialization failed', err);
        setError(
          'Unable to access the camera stream. This is common in secured iframes. Use the simulation console below to test real scanning logic.'
        );
        setCameraActive(false);
      } finally {
        setIsInitializing(false);
      }
    };

    // Delay slightly to ensure element has rendered
    const timer = setTimeout(() => {
      startCamera();
    }, 400);

    return () => {
      clearTimeout(timer);
      if (provider) {
        provider.stopScanning().catch(console.error);
      }
    };
  }, [isOpen, handleScanMatch]);

  const triggerSimulatedScan = (e: React.FormEvent) => {
    e.preventDefault();
    if (simulatedInput.trim()) {
      handleScanMatch(simulatedInput.trim());
      setSimulatedInput('');
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-950/90 backdrop-blur-md p-4">
      <motion.div
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className="w-full max-w-lg bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-2xl relative flex flex-col"
        style={{ maxHeight: '90vh' }}
      >
        {/* Success Flash Overlay */}
        <AnimatePresence>
          {successFlash && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-emerald-500/20 backdrop-blur-xs flex flex-col items-center justify-center z-50 text-center p-6 animate-pulse"
            >
              <div className="bg-emerald-600 text-white p-4 rounded-full shadow-lg mb-3">
                <CheckCircle2 className="w-12 h-12" />
              </div>
              <p className="text-white text-lg font-bold font-mono tracking-wider">
                {successFlash}
              </p>
              <p className="text-emerald-300 text-xs mt-1 uppercase font-semibold tracking-widest">
                Scanned Successfully
              </p>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Header */}
        <div className="p-4 border-b border-slate-800/80 bg-slate-900/55 flex justify-between items-center">
          <div className="flex items-center gap-2">
            <Camera className="w-5 h-5 text-indigo-400 animate-pulse" />
            <div>
              <h3 className="font-bold text-slate-200 text-sm">{title}</h3>
              <p className="text-[10px] text-indigo-300 font-mono">
                SDK: {providerName} (Swappable Bridge Pattern)
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-slate-200 cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Camera Scanner Container */}
        <div className="relative aspect-video w-full bg-slate-950 flex flex-col items-center justify-center overflow-hidden border-b border-slate-800">
          <div id="scanner-reader-container" className="w-full h-full absolute inset-0 z-10" />

          {/* Holographic Laser Scanner effect overlay */}
          {cameraActive && !successFlash && (
            <div className="absolute inset-0 z-20 pointer-events-none flex flex-col items-center justify-center">
              {/* Green active scanning frame */}
              <div className="w-[85%] h-[45%] border-2 border-indigo-500/60 rounded-lg relative shadow-[0_0_15px_rgba(99,102,241,0.2)]">
                {/* Scanning Laser */}
                <div className="absolute left-0 right-0 h-[2px] bg-indigo-400 shadow-[0_0_8px_rgba(129,140,248,1)] top-0 animate-[scan_2.5s_ease-in-out_infinite]" />
                
                {/* Corner markers */}
                <span className="absolute top-0 left-0 w-4 h-4 border-t-4 border-l-4 border-indigo-400 -mt-[3px] -ml-[3px]" />
                <span className="absolute top-0 right-0 w-4 h-4 border-t-4 border-r-4 border-indigo-400 -mt-[3px] -mr-[3px]" />
                <span className="absolute bottom-0 left-0 w-4 h-4 border-b-4 border-l-4 border-indigo-400 -mb-[3px] -ml-[3px]" />
                <span className="absolute bottom-0 right-0 w-4 h-4 border-b-4 border-r-4 border-indigo-400 -mb-[3px] -mr-[3px]" />
              </div>
              <p className="text-[10px] text-slate-400 uppercase tracking-widest mt-4 font-semibold text-center px-4 bg-slate-950/80 py-1 rounded">
                Align barcode within wide reticle area
              </p>
            </div>
          )}

          {/* Initializing / Offline indicator */}
          {isInitializing && (
            <div className="absolute inset-0 bg-slate-950/90 flex flex-col items-center justify-center z-30 space-y-3">
              <div className="w-8 h-8 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin" />
              <p className="text-slate-400 text-xs">Requesting camera permissions...</p>
            </div>
          )}

          {/* Fallback & Info Box */}
          {error && !cameraActive && (
            <div className="absolute inset-0 bg-slate-950 flex flex-col items-center justify-center z-30 p-6 text-center space-y-3">
              <div className="p-3 bg-amber-500/10 text-amber-400 rounded-full">
                <Info className="w-8 h-8" />
              </div>
              <p className="text-slate-300 text-xs font-semibold max-w-sm">
                Sandbox Environment Frame Restriction
              </p>
              <p className="text-slate-500 text-[10px] leading-relaxed max-w-xs">
                To run native haptic barcode scanning on mobile devices, open the app in a new tab. In the meantime, use the simulator below.
              </p>
            </div>
          )}
        </div>

        {/* Continuous Scanning Statistics (Bulk mode) */}
        {isContinuous && (
          <div className="bg-indigo-950/40 border-b border-indigo-900/40 px-4 py-2.5 flex justify-between items-center text-xs">
            <div className="flex items-center gap-2">
              <span className="w-2.5 h-2.5 bg-emerald-500 rounded-full animate-ping" />
              <span className="font-semibold text-indigo-300 uppercase tracking-wider text-[10px]">
                Continuous Scan Active
              </span>
            </div>
            <div className="font-mono bg-indigo-950 border border-indigo-800 text-indigo-200 px-2 py-0.5 rounded font-bold">
              Bags Scanned: {continuousCount}
            </div>
          </div>
        )}

        {/* Simulated Manual Scanner Barcode Entry Console */}
        <div className="p-4 bg-slate-950 border-t border-slate-800/60 flex flex-col space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider flex items-center gap-1">
              <Sparkles className="w-3 h-3 text-indigo-400" />
              Barcode Scanner Simulator Console
            </span>
            <span className="text-[9px] text-slate-500">
              For testing ITF, 1D plates & prefix formats
            </span>
          </div>

          <form onSubmit={triggerSimulatedScan} className="flex gap-2">
            <input
              type="text"
              value={simulatedInput}
              onChange={(e) => setSimulatedInput(e.target.value)}
              placeholder="e.g. 0220123456 or LH123456"
              className="flex-1 bg-slate-900 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded-lg text-xs font-mono tracking-wider placeholder-slate-600"
            />
            <button
              type="submit"
              className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-bold transition shadow-lg shadow-indigo-600/10 cursor-pointer"
            >
              Simulate Scan
            </button>
          </form>

          {/* Quick presets for testers */}
          <div className="flex flex-wrap gap-1.5 pt-1">
            <span className="text-[9px] text-slate-600 self-center">Presets:</span>
            {[
              { val: '0220102943', name: 'LH Plate (0220)' },
              { val: 'LH102943', name: 'LH Prefix (LH)' },
              { val: '0724987654', name: 'LX Plate (0724)' },
              { val: 'LX987654', name: 'LX Prefix (LX)' },
              { val: '0098555444', name: 'AI Plate (0098)' }
            ].map((preset) => (
              <button
                key={preset.val}
                type="button"
                onClick={() => setSimulatedInput(preset.val)}
                className="text-[9px] bg-slate-900 hover:bg-slate-800 border border-slate-800 text-slate-400 hover:text-slate-200 px-2 py-0.5 rounded transition cursor-pointer"
              >
                {preset.name}
              </button>
            ))}
          </div>
        </div>

        {/* Footer controls */}
        <div className="p-4 bg-slate-900 border-t border-slate-800/80 flex justify-end gap-3">
          {isContinuous && onFinishContinuous ? (
            <button
              type="button"
              onClick={onFinishContinuous}
              className="w-full sm:w-auto px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold transition shadow-lg shadow-indigo-600/15 text-center cursor-pointer"
            >
              Finish & Review ({continuousCount} Bags)
            </button>
          ) : (
            <button
              type="button"
              onClick={onClose}
              className="w-full sm:w-auto px-5 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition text-center cursor-pointer"
            >
              Close Scanner
            </button>
          )}
        </div>
      </motion.div>
    </div>
  );
}

// Inject custom CSS for barcode laser scanning animation
if (typeof document !== 'undefined') {
  const style = document.createElement('style');
  style.innerHTML = `
    @keyframes scan {
      0% { top: 0%; opacity: 0.8; }
      50% { top: 100%; opacity: 1; }
      100% { top: 0%; opacity: 0.8; }
    }
  `;
  document.head.appendChild(style);
}
