'use client';

import React, { useEffect, useState, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { X, Camera, Sparkles, CheckCircle2, Info, AlertTriangle, Flashlight } from 'lucide-react';
import { getActiveScanner, IBaggageScanner, ScanResult } from '../lib/scanner';

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
  const [simulatedInput, setSimulatedInput] = useState('');
  const [providerName] = useState(() => getActiveScanner().name);
  const [cameraActive, setCameraActive] = useState(false);
  const [isInitializing, setIsInitializing] = useState(false);
  const [scanStatus, setScanStatus] = useState<'idle' | 'scanning' | 'detecting' | 'success'>('idle');
  const [torchEnabled, setTorchEnabled] = useState(false);
  const [pendingConfirmation, setPendingConfirmation] = useState<ScanResult | null>(null);
  const scannerRef = useRef<IBaggageScanner | null>(null);

  const toggleTorch = async () => {
    if (scannerRef.current?.toggleTorch) {
      const nextState = !torchEnabled;
      await scannerRef.current.toggleTorch(nextState);
      setTorchEnabled(nextState);
    }
  };

  const handleScanMatch = React.useCallback((result: ScanResult) => {
    if (pendingConfirmation) return; // Ignore if already waiting for confirmation
    
    setPendingConfirmation(result);
    setScanStatus('success');
  }, [pendingConfirmation]);

  const confirmScan = () => {
    if (!pendingConfirmation) return;
    
    const result = pendingConfirmation;
    onScanSuccess(result.tagNumber);
    setPendingConfirmation(null);

    if (!isContinuous) {
      onClose();
    } else {
      setScanStatus('scanning');
    }
  };

  useEffect(() => {
    const provider = getActiveScanner();
    scannerRef.current = provider;

    if (!isOpen) {
      provider.stopScanning().catch(console.error);
      return;
    }

    const startCamera = async () => {
      setIsInitializing(true);
      setError(null);
      setScanStatus('idle');
      try {
        await provider.init();
        await provider.startScanning(
          'scanner-reader-container',
          (result) => {
            handleScanMatch(result);
          },
          (err) => {
            if (typeof err === 'string' && err.includes('Invalid')) {
              setError(err);
              setTimeout(() => setError(null), 3000);
            }
          }
        );
        setCameraActive(true);
        setScanStatus('scanning');
      } catch (err: any) {
        console.error('Camera initialization failed', err);
        setError('Unable to access camera. Please check permissions.');
        setCameraActive(false);
      } finally {
        setIsInitializing(false);
      }
    };

    const timer = setTimeout(startCamera, 400);

    return () => {
      clearTimeout(timer);
      if (provider) provider.stopScanning().catch(console.error);
    };
  }, [isOpen, handleScanMatch]);

  const triggerSimulatedScan = (e: React.FormEvent) => {
    e.preventDefault();
    if (simulatedInput.trim()) {
      handleScanMatch({
        tagNumber: simulatedInput.trim().toUpperCase(),
        barcodeFormat: 'SIMULATED',
        scannerUsed: 'SIMULATOR',
        confidenceScore: 1.0,
        scanTime: 0
      });
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
        {/* Confirmation Overlay */}
        <AnimatePresence>
          {pendingConfirmation && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-slate-950/95 backdrop-blur-xl flex flex-col items-center justify-center z-[100] text-center p-8"
            >
              <motion.div 
                initial={{ scale: 0.8, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                className="bg-indigo-600 text-white p-6 rounded-full shadow-2xl mb-6 border-4 border-indigo-400/30"
              >
                <CheckCircle2 className="w-20 h-20" />
              </motion.div>
              
              <h2 className="text-white text-2xl font-black mb-2 uppercase tracking-tight">
                Baggage Scanned
              </h2>
              
              <div className="bg-slate-900 border-2 border-indigo-500/50 px-8 py-4 rounded-2xl mb-8 shadow-inner">
                <p className="text-indigo-400 text-[10px] uppercase font-black tracking-[0.2em] mb-1">
                  Tag Number
                </p>
                <p className="text-white text-4xl font-black font-mono tracking-widest">
                  {pendingConfirmation.tagNumber}
                </p>
              </div>

              <button
                onClick={confirmScan}
                className="w-full max-w-xs py-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded-2xl text-lg font-black uppercase tracking-widest transition-all shadow-2xl shadow-indigo-600/30 active:scale-95 flex items-center justify-center gap-3"
              >
                Confirm OK
                <CheckCircle2 className="w-6 h-6" />
              </button>
              
              <p className="mt-6 text-slate-500 text-[10px] uppercase font-bold tracking-widest">
                Verify tag before processing
              </p>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Header */}
        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: -20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="absolute top-20 left-1/2 -translate-x-1/2 z-[60] bg-red-600 text-white px-4 py-2 rounded-full shadow-xl flex items-center gap-2 text-xs font-bold whitespace-nowrap"
            >
              <AlertTriangle className="w-4 h-4" />
              {error}
            </motion.div>
          )}
        </AnimatePresence>

        {/* Header */}
        <div className="p-4 border-b border-slate-800/80 bg-slate-900/55 flex justify-between items-center">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-indigo-500/10 flex items-center justify-center border border-indigo-500/20">
              <Camera className={`w-4 h-4 text-indigo-400 ${scanStatus === 'scanning' ? 'animate-pulse' : ''}`} />
            </div>
            <div>
              <h3 className="font-bold text-slate-200 text-sm tracking-tight">{title}</h3>
              <div className="flex items-center gap-2">
                <span className="text-[9px] text-indigo-400 font-mono font-bold uppercase">
                  IATA-Grade Engine Active
                </span>
                <span className="w-1 h-1 rounded-full bg-slate-700" />
                <span className="text-[9px] text-slate-500 font-mono uppercase">
                  {providerName}
                </span>
              </div>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-slate-200 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Camera Container */}
        <div className="relative aspect-video w-full bg-slate-950 overflow-hidden border-b border-slate-800 group">
          <div id="scanner-reader-container" className="w-full h-full absolute inset-0 z-10 grayscale-[0.2]" />

          {/* Torch Toggle Button */}
          {cameraActive && (
            <button
              onClick={toggleTorch}
              className={`absolute top-4 right-4 z-30 p-3 rounded-full backdrop-blur-md transition-all border ${
                torchEnabled 
                  ? 'bg-amber-500/20 border-amber-500/50 text-amber-400 shadow-[0_0_15px_rgba(245,158,11,0.3)]' 
                  : 'bg-slate-950/40 border-slate-700/50 text-slate-400 hover:bg-slate-900/60'
              }`}
            >
              <Flashlight className={`w-5 h-5 ${torchEnabled ? 'fill-current' : ''}`} />
            </button>
          )}

          {/* Reticle & Guidance Overlay */}
          {cameraActive && !pendingConfirmation && (
            <div className="absolute inset-0 z-20 pointer-events-none flex flex-col items-center justify-center">
              {/* Wide IATA Reticle */}
              <div className="w-[85%] h-[40%] border border-white/20 rounded-2xl relative">
                {/* Active Corners */}
                <span className="absolute -top-[2px] -left-[2px] w-8 h-8 border-t-4 border-l-4 border-indigo-500 rounded-tl-xl" />
                <span className="absolute -top-[2px] -right-[2px] w-8 h-8 border-t-4 border-r-4 border-indigo-500 rounded-tr-xl" />
                <span className="absolute -bottom-[2px] -left-[2px] w-8 h-8 border-b-4 border-l-4 border-indigo-500 rounded-bl-xl" />
                <span className="absolute -bottom-[2px] -right-[2px] w-8 h-8 border-b-4 border-r-4 border-indigo-500 rounded-br-xl" />
                
                {/* Horizontal Scan Laser */}
                <motion.div 
                  animate={{ top: ['10%', '90%', '10%'] }}
                  transition={{ duration: 2.5, repeat: Infinity, ease: "easeInOut" }}
                  className="absolute left-0 right-0 h-[2px] bg-gradient-to-r from-transparent via-indigo-400 to-transparent shadow-[0_0_15px_rgba(129,140,248,0.8)] opacity-60" 
                />

                {/* HUD Elements */}
                <div className="absolute top-2 left-2 text-[8px] font-mono text-indigo-400/60 uppercase font-bold">
                  ITF / C128 Priority
                </div>
                <div className="absolute bottom-2 right-2 text-[8px] font-mono text-indigo-400/60 uppercase font-bold">
                  Autofocus Active
                </div>
              </div>

              {/* Dynamic Status Text */}
              <div className="mt-8 bg-slate-950/80 backdrop-blur-md px-6 py-2 rounded-full border border-slate-800 flex items-center gap-3">
                <div className="w-2 h-2 rounded-full bg-indigo-500 animate-ping" />
                <span className="text-[11px] text-slate-100 font-bold uppercase tracking-widest">
                  {scanStatus === 'scanning' ? 'Scanning...' : 'Hold Steady'}
                </span>
              </div>
              
              <p className="text-[9px] text-slate-500 uppercase tracking-widest mt-4 font-bold text-center px-6">
                Align 10-digit barcode within reticle
              </p>
            </div>
          )}

          {/* Initializing State */}
          {isInitializing && (
            <div className="absolute inset-0 bg-slate-950 flex flex-col items-center justify-center z-30 space-y-4">
              <div className="relative">
                <div className="w-12 h-12 border-2 border-indigo-500/20 rounded-full" />
                <div className="w-12 h-12 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin absolute inset-0" />
              </div>
              <div className="text-center">
                <p className="text-slate-100 text-sm font-bold tracking-tight">Initializing Camera</p>
                <p className="text-slate-500 text-[10px] uppercase font-bold tracking-widest mt-1">Requesting permissions...</p>
              </div>
            </div>
          )}

          {/* Error Fallback */}
          {!cameraActive && !isInitializing && (
            <div className="absolute inset-0 bg-slate-950 flex flex-col items-center justify-center z-30 p-8 text-center">
              <div className="w-16 h-16 bg-slate-900 rounded-3xl flex items-center justify-center mb-6 border border-slate-800">
                <Info className="w-8 h-8 text-slate-500" />
              </div>
              <h4 className="text-slate-100 font-bold mb-2">Camera Access Restricted</h4>
              <p className="text-slate-500 text-xs leading-relaxed max-w-xs mx-auto mb-6">
                The browser has blocked camera access in this frame. Open the app in a new tab for native mobile scanning, or use the simulator below.
              </p>
              <button 
                onClick={() => window.location.reload()}
                className="text-[10px] text-indigo-400 font-bold uppercase tracking-widest border border-indigo-500/30 px-4 py-2 rounded-lg hover:bg-indigo-500/5 transition-colors"
              >
                Retry Initialization
              </button>
            </div>
          )}
        </div>

        {/* Continuous Scanning Stats */}
        {isContinuous && (
          <div className="bg-indigo-950/40 border-b border-indigo-900/40 px-5 py-3 flex justify-between items-center">
            <div className="flex items-center gap-3">
              <div className="relative">
                <div className="w-2 h-2 bg-emerald-500 rounded-full animate-ping" />
                <div className="w-2 h-2 bg-emerald-500 rounded-full absolute inset-0" />
              </div>
              <span className="font-black text-indigo-300 uppercase tracking-tighter text-xs">
                Continuous Ops Active
              </span>
            </div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] text-slate-500 font-bold uppercase">Bags Logged:</span>
              <div className="font-mono bg-indigo-500 text-white px-3 py-1 rounded-lg font-black text-sm shadow-lg shadow-indigo-600/20">
                {continuousCount}
              </div>
            </div>
          </div>
        )}

        {/* Simulator Console */}
        <div className="p-5 bg-slate-950 border-t border-slate-800/40">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2">
              <div className="p-1 bg-indigo-500/10 rounded border border-indigo-500/20">
                <Sparkles className="w-3 h-3 text-indigo-400" />
              </div>
              <span className="text-[10px] uppercase font-black text-slate-400 tracking-wider">
                Baggage Tag Simulator
              </span>
            </div>
            <span className="text-[9px] text-slate-600 font-bold uppercase">Test IATA Formats</span>
          </div>

          <form onSubmit={triggerSimulatedScan} className="flex gap-2 mb-4">
            <input
              type="text"
              value={simulatedInput}
              onChange={(e) => setSimulatedInput(e.target.value)}
              placeholder="e.g. 0220123456"
              className="flex-1 bg-slate-900 border border-slate-800 focus:border-indigo-500 text-slate-200 px-4 py-2.5 rounded-xl text-xs font-mono tracking-widest placeholder-slate-700 outline-none transition-all"
            />
            <button
              type="submit"
              className="px-6 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-black uppercase tracking-tighter transition shadow-xl shadow-indigo-600/10 active:scale-95"
            >
              Simulate
            </button>
          </form>

          <div className="flex flex-wrap gap-1.5">
            {[
              { val: '0220102943', name: 'LH ITF' },
              { val: 'LH102943', name: 'LH C128' },
              { val: '0724987654', name: 'LX ITF' },
              { val: 'LX987654', name: 'LX C128' }
            ].map((preset) => (
              <button
                key={preset.val}
                type="button"
                onClick={() => setSimulatedInput(preset.val)}
                className="text-[9px] bg-slate-900 hover:bg-slate-800 border border-slate-800 text-slate-500 hover:text-slate-200 px-3 py-1.5 rounded-lg transition-all font-bold uppercase"
              >
                {preset.name}
              </button>
            ))}
          </div>
        </div>

        {/* Action Bar */}
        <div className="p-5 bg-slate-900 border-t border-slate-800/60 flex flex-col sm:flex-row gap-3">
          {isContinuous && onFinishContinuous ? (
            <button
              type="button"
              onClick={onFinishContinuous}
              className="flex-1 px-6 py-3.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-2xl text-xs font-black uppercase tracking-widest transition shadow-2xl shadow-indigo-600/20 text-center active:scale-[0.98]"
            >
              Complete Batch ({continuousCount} Bags)
            </button>
          ) : (
            <button
              type="button"
              onClick={onClose}
              className="flex-1 px-6 py-3.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-2xl text-xs font-black uppercase tracking-widest transition text-center"
            >
              Cancel Operation
            </button>
          )}
        </div>
      </motion.div>
    </div>
  );
}
