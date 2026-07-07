'use client';

import React, { useEffect, useState, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  X, 
  Camera, 
  RefreshCw, 
  AlertCircle, 
  Upload, 
  Check, 
  Zap, 
  ZapOff, 
  RotateCw, 
  Sparkles, 
  Sliders, 
  Edit, 
  Plus, 
  Trash2, 
  Eye, 
  ZoomIn, 
  ZoomOut, 
  Maximize2, 
  Layers, 
  ClipboardList, 
  Settings, 
  CheckCircle2, 
  AlertTriangle, 
  FileSpreadsheet, 
  Loader2,
  Trash
} from 'lucide-react';

interface CapturedPage {
  id: string;
  originalUrl: string;
  processedUrl: string;
  rotation: number; // 0, 90, 180, 270
  contrastBoosted: boolean;
  sharpened: boolean;
}

interface TableRow {
  flightNo: string;
  receivedAt: string;
  originalTag: string;
  name: string;
  weight: string;
  rushTag: string;
  remarks: string;
  confidence?: 'high' | 'medium' | 'low';
  lowConfidenceFields?: string[];
}

interface OcrCameraModalProps {
  isOpen: boolean;
  onClose: () => void;
  onImportText: (text: string) => void;
  onImportTable: (data: any[]) => void;
  mode: 'altea' | 'table';
}

export default function OcrCameraModal({ isOpen, onClose, onImportText, onImportTable, mode }: OcrCameraModalProps) {
  // Navigation / Workspace States
  const [viewMode, setViewMode] = useState<'capturing' | 'processing' | 'reviewing'>('capturing');
  
  // Camera Stream States
  const [stream, setStream] = useState<MediaStream | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [facingMode, setFacingMode] = useState<'environment' | 'user'>('environment');
  const [devices, setDevices] = useState<MediaDeviceInfo[]>([]);
  const [selectedDeviceId, setSelectedDeviceId] = useState<string>('');
  const [isInitializing, setIsInitializing] = useState(false);
  const [torchSupported, setTorchSupported] = useState(false);
  const [torchOn, setTorchOn] = useState(false);
  const [cameraTrigger, setCameraTrigger] = useState(0);

  // Multi-Image Queue States
  const [capturedPages, setCapturedPages] = useState<CapturedPage[]>([]);
  const [activePageIndex, setActivePageIndex] = useState<number>(0);

  // OCR Sequential Queue Progress States
  const [processingProgress, setProcessingProgress] = useState(0);
  const [processingStage, setProcessingStage] = useState('');

  // OCR Raw Extracted / Editable Review States
  const [ocrTextResult, setOcrTextResult] = useState('');
  const [ocrTableRows, setOcrTableRows] = useState<TableRow[]>([]);

  // Interactive Zoom / Pan State
  const [zoom, setZoom] = useState(1);
  const [pan, setPan] = useState({ x: 0, y: 0 });
  const [isDragging, setIsDragging] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });

  // Camera Refs
  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Stop camera stream
  const stopCamera = () => {
    if (stream) {
      stream.getTracks().forEach((track) => track.stop());
      setTimeout(() => {
        setStream(null);
        setTorchOn(false);
      }, 0);
    } else {
      setTimeout(() => {
        setTorchOn(false);
      }, 0);
    }
  };

  // Get available video devices
  const getDevices = async () => {
    try {
      const allDevices = await navigator.mediaDevices.enumerateDevices();
      const videoDevices = allDevices.filter((device) => device.kind === 'videoinput');
      setDevices(videoDevices);
    } catch (err) {
      console.error('Error listing devices:', err);
    }
  };

  // Camera stream lifecycle
  useEffect(() => {
    let active = true;

    const startCamera = async () => {
      if (!isOpen || viewMode !== 'capturing') return;
      setIsInitializing(true);
      setError(null);

      try {
        await getDevices();

        // Maximize resolution and options for enterprise quality
        const constraints: MediaStreamConstraints = {
          video: selectedDeviceId 
            ? { 
                deviceId: { exact: selectedDeviceId },
                width: { ideal: 4096, min: 1920 },
                height: { ideal: 3072, min: 1080 }
              }
            : { 
                facingMode: facingMode,
                width: { ideal: 4096, min: 1920 },
                height: { ideal: 3072, min: 1080 }
              }
        };

        const mediaStream = await navigator.mediaDevices.getUserMedia(constraints);
        
        if (active) {
          setStream(mediaStream);
          
          // Check for torch support
          const videoTrack = mediaStream.getVideoTracks()[0];
          if (videoTrack) {
            const capabilities = videoTrack.getCapabilities?.();
            // @ts-ignore
            if (capabilities && 'torch' in capabilities) {
              setTorchSupported(true);
            } else {
              setTorchSupported(false);
            }
          }
        } else {
          mediaStream.getTracks().forEach((track) => track.stop());
        }
      } catch (err: any) {
        console.error('Error accessing high-resolution camera:', err);
        if (active) {
          // Fallback to standard video stream
          try {
            const fallbackStream = await navigator.mediaDevices.getUserMedia({ 
              video: { width: { ideal: 1920 }, height: { ideal: 1080 } } 
            });
            if (active) {
              setStream(fallbackStream);
            } else {
              fallbackStream.getTracks().forEach((track) => track.stop());
            }
          } catch (fallbackErr) {
            setError('Camera access denied or unavailable. Please upload an image file instead.');
          }
        }
      } finally {
        if (active) {
          setIsInitializing(false);
        }
      }
    };

    startCamera();

    return () => {
      active = false;
      if (stream) {
        stream.getTracks().forEach((track) => track.stop());
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, facingMode, selectedDeviceId, cameraTrigger, viewMode]);

  // Bind stream to video element
  useEffect(() => {
    if (videoRef.current && stream) {
      videoRef.current.srcObject = stream;
    }
  }, [stream]);

  // Reset local states on close/open
  useEffect(() => {
    if (isOpen) {
      const timer = setTimeout(() => {
        setViewMode('capturing');
        setCapturedPages([]);
        setActivePageIndex(0);
        setOcrTextResult('');
        setOcrTableRows([]);
        setZoom(1);
        setPan({ x: 0, y: 0 });
      }, 0);
      return () => clearTimeout(timer);
    } else {
      stopCamera();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen]);

  // Toggle front/rear
  const switchCamera = () => {
    if (devices.length > 1) {
      const currentIndex = devices.findIndex(d => d.deviceId === selectedDeviceId);
      const nextIndex = (currentIndex + 1) % devices.length;
      setSelectedDeviceId(devices[nextIndex].deviceId);
    } else {
      setFacingMode((prev) => (prev === 'environment' ? 'user' : 'environment'));
    }
  };

  // Toggle Torch
  const toggleTorch = async () => {
    if (!stream || !torchSupported) return;
    const videoTrack = stream.getVideoTracks()[0];
    if (videoTrack) {
      try {
        const nextState = !torchOn;
        await videoTrack.applyConstraints({
          // @ts-ignore
          advanced: [{ torch: nextState }]
        });
        setTorchOn(nextState);
      } catch (err) {
        console.error('Failed to toggle torch:', err);
      }
    }
  };

  // Rotate camera or captured view
  const rotatePageClockwise = (pageIndex: number) => {
    const updated = [...capturedPages];
    const page = updated[pageIndex];
    const nextRotation = (page.rotation + 90) % 360;
    page.rotation = nextRotation;

    applyCanvasFilters(
      page.originalUrl,
      nextRotation,
      page.contrastBoosted,
      page.sharpened,
      (newUrl) => {
        page.processedUrl = newUrl;
        setCapturedPages(updated);
      }
    );
  };

  // Contrast enhancement filter
  const toggleContrastBooster = (pageIndex: number) => {
    const updated = [...capturedPages];
    const page = updated[pageIndex];
    const nextVal = !page.contrastBoosted;
    page.contrastBoosted = nextVal;

    applyCanvasFilters(
      page.originalUrl,
      page.rotation,
      nextVal,
      page.sharpened,
      (newUrl) => {
        page.processedUrl = newUrl;
        setCapturedPages(updated);
      }
    );
  };

  // Sharpening 3x3 filter
  const toggleTextSharpening = (pageIndex: number) => {
    const updated = [...capturedPages];
    const page = updated[pageIndex];
    const nextVal = !page.sharpened;
    page.sharpened = nextVal;

    applyCanvasFilters(
      page.originalUrl,
      page.rotation,
      page.contrastBoosted,
      nextVal,
      (newUrl) => {
        page.processedUrl = newUrl;
        setCapturedPages(updated);
      }
    );
  };

  // Apply visual enhancements client-side on high-resolution image
  const applyCanvasFilters = (
    originalUrl: string,
    rotation: number,
    boostContrast: boolean,
    sharpen: boolean,
    onComplete: (newUrl: string) => void
  ) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        onComplete(originalUrl);
        return;
      }

      const angle = (rotation * Math.PI) / 180;
      const is90or270 = (rotation / 90) % 2 !== 0;

      canvas.width = is90or270 ? img.height : img.width;
      canvas.height = is90or270 ? img.width : img.height;

      // Adjust coordinates and draw
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate(angle);
      ctx.drawImage(img, -img.width / 2, -img.height / 2);
      ctx.setTransform(1, 0, 0, 1, 0, 0); // reset transform

      // Pixel adjustment
      if (boostContrast || sharpen) {
        try {
          const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
          const data = imgData.data;

          if (boostContrast) {
            const factor = (259 * (80 + 255)) / (255 * (259 - 80));
            for (let i = 0; i < data.length; i += 4) {
              data[i]     = factor * (data[i] - 128) + 128;     // R
              data[i + 1] = factor * (data[i + 1] - 128) + 128; // G
              data[i + 2] = factor * (data[i + 2] - 128) + 128; // B
            }
          }

          ctx.putImageData(imgData, 0, 0);

          if (sharpen) {
            const sharpenData = ctx.createImageData(canvas.width, canvas.height);
            const sData = sharpenData.data;
            const w = canvas.width;
            const h = canvas.height;

            // Copy alpha/edges
            const currentImgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const cData = currentImgData.data;
            for (let i = 0; i < cData.length; i++) {
              sData[i] = cData[i];
            }

            // High precision sharpening kernel
            const kernel = [
               0, -1,  0,
              -1,  5, -1,
               0, -1,  0
            ];

            for (let y = 1; y < h - 1; y++) {
              for (let x = 1; x < w - 1; x++) {
                const idx = (y * w + x) * 4;
                let r = 0, g = 0, b = 0;

                for (let ky = -1; ky <= 1; ky++) {
                  for (let kx = -1; kx <= 1; kx++) {
                    const pIdx = ((y + ky) * w + (x + kx)) * 4;
                    const weight = kernel[(ky + 1) * 3 + (kx + 1)];
                    r += cData[pIdx] * weight;
                    g += cData[pIdx + 1] * weight;
                    b += cData[pIdx + 2] * weight;
                  }
                }

                sData[idx] = Math.min(255, Math.max(0, r));
                sData[idx + 1] = Math.min(255, Math.max(0, g));
                sData[idx + 2] = Math.min(255, Math.max(0, b));
              }
            }
            ctx.putImageData(sharpenData, 0, 0);
          }
        } catch (err) {
          console.error('Error applying visual filters:', err);
        }
      }

      onComplete(canvas.toDataURL('image/jpeg', 0.9));
    };
    img.src = originalUrl;
  };

  // Capture still photograph
  const capturePhoto = () => {
    if (videoRef.current && canvasRef.current) {
      const video = videoRef.current;
      const canvas = canvasRef.current;
      const ctx = canvas.getContext('2d');

      if (ctx) {
        // Full resolution native capture
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        
        if (facingMode === 'user') {
          ctx.translate(canvas.width, 0);
          ctx.scale(-1, 1);
        }
        
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        
        if (facingMode === 'user') {
          ctx.setTransform(1, 0, 0, 1, 0, 0);
        }

        const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
        
        const newPage: CapturedPage = {
          id: `page_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
          originalUrl: dataUrl,
          processedUrl: dataUrl,
          rotation: 0,
          contrastBoosted: false,
          sharpened: false
        };

        setCapturedPages((prev) => [...prev, newPage]);
        setActivePageIndex(capturedPages.length); // switch focus to the new page
      }
    }
  };

  // Fallback Device File Selection
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result as string;
        const newPage: CapturedPage = {
          id: `page_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
          originalUrl: dataUrl,
          processedUrl: dataUrl,
          rotation: 0,
          contrastBoosted: false,
          sharpened: false
        };
        setCapturedPages((prev) => [...prev, newPage]);
        setActivePageIndex(capturedPages.length);
      };
      reader.readAsDataURL(file);
    }
  };

  // Deletes a page from the sequential queue
  const deletePage = (index: number) => {
    const updated = capturedPages.filter((_, idx) => idx !== index);
    setCapturedPages(updated);
    if (activePageIndex >= updated.length && updated.length > 0) {
      setActivePageIndex(updated.length - 1);
    }
  };

  // Reset the entire queue and go back to camera
  const restartCapturing = () => {
    setCapturedPages([]);
    setActivePageIndex(0);
    setViewMode('capturing');
  };

  // Execute sequential OCR pipeline for all captured pages
  const handleRunSequentialOcr = async () => {
    if (capturedPages.length === 0) return;

    setViewMode('processing');
    setProcessingProgress(5);
    setProcessingStage('Pre-processing pages and enhancing text borders...');

    stopCamera();

    let mergedText = '';
    const mergedRows: TableRow[] = [];

    try {
      for (let i = 0; i < capturedPages.length; i++) {
        const pageNum = i + 1;
        const page = capturedPages[i];
        const progressStart = Math.round((i / capturedPages.length) * 100);
        
        setProcessingProgress(progressStart + 10);
        setProcessingStage(`Scanning Page ${pageNum} of ${capturedPages.length} with Gemini OCR Engine...`);

        const response = await fetch('/api/ocr', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ image: page.processedUrl, mode }),
        });

        const result = await response.json();
        if (result.error) {
          throw new Error(`Page ${pageNum} failed: ${result.error}`);
        }

        if (mode === 'altea') {
          const text = result.text || '';
          mergedText += (mergedText ? '\n\n' : '') + `--- PAGE ${pageNum} ---\n` + text;
        } else {
          if (result.data && Array.isArray(result.data)) {
            const mapped = result.data.map((item: any) => ({
              flightNo: item.flightNo || '',
              receivedAt: item.receivedAt || '',
              originalTag: item.originalTag || '',
              name: item.name || '',
              weight: item.weight || '',
              rushTag: item.rushTag || '',
              remarks: item.remarks || '',
              confidence: item.confidence || 'high',
              lowConfidenceFields: item.lowConfidenceFields || []
            }));
            mergedRows.push(...mapped);
          }
        }
      }

      setProcessingProgress(95);
      setProcessingStage('Assembling data grid and computing confidence scores...');

      // Transition to Review Workspace
      setOcrTextResult(mergedText);
      setOcrTableRows(mergedRows.length > 0 ? mergedRows : [
        { flightNo: '', receivedAt: '', originalTag: '', name: '', weight: '', rushTag: '', remarks: '', confidence: 'low', lowConfidenceFields: ['originalTag'] }
      ]);
      
      setViewMode('reviewing');
      setZoom(1);
      setPan({ x: 0, y: 0 });
    } catch (err: any) {
      console.error('OCR Processing error:', err);
      setError(`OCR Processing failed: ${err.message || 'Unknown error'}`);
      setViewMode('capturing');
    }
  };

  // Submit verified data to parent workflow
  const commitImport = () => {
    if (mode === 'altea') {
      onImportText(ocrTextResult);
    } else {
      // Map back to format expected by parent
      onImportTable(ocrTableRows);
    }
    onClose();
  };

  // Interactive Zoom / Pan Handlers
  const handleMouseDown = (e: React.MouseEvent) => {
    if (zoom === 1) return;
    setIsDragging(true);
    setDragStart({ x: e.clientX - pan.x, y: e.clientY - pan.y });
  };

  const handleMouseMove = (e: React.MouseEvent) => {
    if (!isDragging) return;
    setPan({
      x: e.clientX - dragStart.x,
      y: e.clientY - dragStart.y
    });
  };

  const handleMouseUp = () => {
    setIsDragging(false);
  };

  const handleZoomIn = () => {
    setZoom((prev) => Math.min(4, prev + 0.5));
  };

  const handleZoomOut = () => {
    setZoom((prev) => {
      const next = Math.max(1, prev - 0.5);
      if (next === 1) setPan({ x: 0, y: 0 });
      return next;
    });
  };

  const handleZoomReset = () => {
    setZoom(1);
    setPan({ x: 0, y: 0 });
  };

  // Row operations for table reviews
  const updateTableCell = (index: number, field: keyof TableRow, value: string) => {
    const updated = [...ocrTableRows];
    // @ts-ignore
    updated[index][field] = value;
    
    // Auto-resolve low confidence flag on edit
    if (updated[index].lowConfidenceFields?.includes(field)) {
      updated[index].lowConfidenceFields = updated[index].lowConfidenceFields?.filter(f => f !== field);
      if ((updated[index].lowConfidenceFields?.length || 0) === 0) {
        updated[index].confidence = 'high';
      }
    }
    setOcrTableRows(updated);
  };

  const addTableRow = () => {
    setOcrTableRows((prev) => [
      ...prev,
      { flightNo: '', receivedAt: '', originalTag: '', name: '', weight: '', rushTag: '', remarks: '', confidence: 'high' }
    ]);
  };

  const deleteTableRow = (index: number) => {
    setOcrTableRows((prev) => prev.filter((_, idx) => idx !== index));
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-slate-950/98 backdrop-blur-xl z-[150] flex flex-col select-none overflow-hidden text-slate-200">
      
      {/* HEADER SECTION */}
      <div className="w-full bg-slate-900/60 border-b border-slate-800/80 px-6 py-4 flex items-center justify-between z-10 shrink-0">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-indigo-500/10 border border-indigo-500/20 rounded-xl">
            <Camera className="w-5 h-5 text-indigo-400 animate-pulse" />
          </div>
          <div>
            <h3 className="text-white font-bold text-sm tracking-wide flex items-center gap-2">
              Enterprise Camera OCR Engine
              <span className="text-[10px] font-mono bg-indigo-500/20 text-indigo-300 px-2 py-0.5 rounded uppercase font-bold">
                v3.5 Multi-Page
              </span>
            </h3>
            <p className="text-[11px] text-slate-400">
              {viewMode === 'capturing' && 'Capture spreadsheet worksheets, printed lists, or cryptic screen terminals'}
              {viewMode === 'processing' && 'Running intelligent document enhancement and sequence processing...'}
              {viewMode === 'reviewing' && 'Cross-check, edit, and verify OCR output side-by-side with original photos'}
            </p>
          </div>
        </div>
        <button
          onClick={onClose}
          className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white rounded-xl transition cursor-pointer"
        >
          <X className="w-4 h-4" />
        </button>
      </div>

      {/* VIEWPORT AREA */}
      <div className="flex-1 overflow-hidden relative flex">
        
        {/* VIEWMODE: CAPTURING (Camera Stream / Thumbnail Queue) */}
        {viewMode === 'capturing' && (
          <div className="flex-1 flex flex-col md:flex-row h-full">
            
            {/* Live Frame Pane */}
            <div className="flex-1 bg-slate-950 flex flex-col items-center justify-center p-4 relative">
              
              {/* Guidance HUD at top */}
              <div className="absolute top-6 left-1/2 -translate-x-1/2 bg-slate-900/90 border border-slate-800/80 backdrop-blur px-4 py-2.5 rounded-2xl flex items-center gap-2.5 z-20 text-xs shadow-xl max-w-md pointer-events-none">
                <Sparkles className="w-4 h-4 text-amber-400 shrink-0 animate-bounce" />
                <span className="text-[11px] text-slate-300 leading-tight">
                  {mode === 'altea' 
                    ? "Tip: Point directly at the monitor. Keep console lines straight and glare minimal."
                    : "Tip: Align Excel sheets with guidance lines. Hold device level for micro-fonts."
                  }
                </span>
              </div>

              {/* Viewport Frame */}
              <div className="relative w-full max-w-lg aspect-[3/4] max-h-[58vh] bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden flex items-center justify-center shadow-2xl">
                {isInitializing && (
                  <div className="absolute inset-0 flex flex-col items-center justify-center bg-slate-950/80 backdrop-blur-sm z-20">
                    <Loader2 className="w-8 h-8 text-indigo-400 animate-spin mb-3" />
                    <span className="text-xs text-indigo-300 font-mono">Calibrating native focal lens...</span>
                  </div>
                )}

                {error ? (
                  <div className="p-8 text-center max-w-sm flex flex-col items-center gap-4">
                    <AlertCircle className="w-12 h-12 text-rose-500 animate-pulse" />
                    <p className="text-slate-300 text-sm font-medium leading-relaxed">{error}</p>
                    <button
                      onClick={() => fileInputRef.current?.click()}
                      className="bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold px-5 py-3 rounded-xl flex items-center gap-2 mx-auto transition shadow-lg shadow-indigo-600/20"
                    >
                      <Upload className="w-4 h-4" />
                      Select Photo from Files
                    </button>
                  </div>
                ) : (
                  <>
                    <video
                      ref={videoRef}
                      autoPlay
                      playsInline
                      muted
                      className={`w-full h-full object-cover ${facingMode === 'user' ? 'scale-x-[-1]' : ''}`}
                    />
                    
                    {/* Live Guidance Framing HUD */}
                    <div className="absolute inset-0 border-[28px] border-slate-950/40 pointer-events-none flex items-center justify-center">
                      <div className="w-full h-full border border-dashed border-indigo-400/40 rounded-xl relative">
                        {/* Shimmer Scan Line */}
                        <div className="absolute left-0 right-0 h-0.5 bg-indigo-400/30 top-1/2 -translate-y-1/2 animate-pulse" />
                        
                        {/* High tech corner scopes */}
                        <div className="absolute top-0 left-0 w-6 h-6 border-t-2 border-l-2 border-indigo-400" />
                        <div className="absolute top-0 right-0 w-6 h-6 border-t-2 border-r-2 border-indigo-400" />
                        <div className="absolute bottom-0 left-0 w-6 h-6 border-b-2 border-l-2 border-indigo-400" />
                        <div className="absolute bottom-0 right-0 w-6 h-6 border-b-2 border-r-2 border-indigo-400" />
                        
                        {/* Grid Guides */}
                        <div className="absolute inset-0 grid grid-cols-3 grid-rows-3 opacity-20">
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                          <div className="border border-slate-700"></div>
                        </div>

                        <div className="absolute bottom-4 left-1/2 -translate-x-1/2 bg-slate-950/80 px-3 py-1 rounded-full border border-slate-800 text-[9px] text-indigo-300 font-mono tracking-wider uppercase">
                          {mode === 'altea' ? 'Terminal Screen Target' : 'Document grid bounds'}
                        </div>
                      </div>
                    </div>
                  </>
                )}
              </div>

              {/* Shutter / Controls Panel */}
              <div className="w-full max-w-lg mt-5 flex items-center justify-between px-6 z-10 shrink-0">
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="p-3.5 bg-slate-900 border border-slate-800/80 text-slate-300 hover:text-white rounded-2xl transition shadow-lg hover:border-indigo-500/20"
                  title="Upload Image File"
                >
                  <Upload className="w-5 h-5" />
                </button>

                {/* Shutter Trigger */}
                <button
                  onClick={capturePhoto}
                  disabled={isInitializing || !!error}
                  className={`w-16 h-16 rounded-full border-4 border-slate-950 flex items-center justify-center p-1 transition shadow-2xl ${
                    isInitializing || !!error 
                      ? 'bg-slate-800 opacity-40 cursor-not-allowed' 
                      : 'bg-indigo-500 hover:scale-105 active:scale-95'
                  }`}
                >
                  <div className="w-full h-full rounded-full bg-white hover:bg-slate-100 transition" />
                </button>

                <div className="flex gap-2">
                  {torchSupported && (
                    <button
                      onClick={toggleTorch}
                      className={`p-3.5 rounded-2xl transition border ${
                        torchOn 
                          ? 'bg-amber-500/20 text-amber-400 border-amber-500/30 shadow-lg' 
                          : 'bg-slate-900 border-slate-800/80 text-slate-300 hover:text-white'
                      }`}
                      title="Toggle Torch Flash"
                    >
                      {torchOn ? <Zap className="w-5 h-5" /> : <ZapOff className="w-5 h-5" />}
                    </button>
                  )}
                  
                  <button
                    onClick={switchCamera}
                    disabled={devices.length <= 1 && facingMode === 'environment'}
                    className="p-3.5 bg-slate-900 border border-slate-800/80 text-slate-300 hover:text-white rounded-2xl transition disabled:opacity-40 shadow-lg"
                    title="Toggle Lens"
                  >
                    <RefreshCw className="w-5 h-5" />
                  </button>
                </div>
              </div>
            </div>

            {/* Queue Deck Side panel (Desktop: side, Mobile: bottom or collapsible) */}
            <div className="w-full md:w-80 bg-slate-900 border-t md:border-t-0 md:border-l border-slate-800/80 p-5 flex flex-col justify-between shrink-0">
              <div className="space-y-4 flex-1 overflow-y-auto">
                <div className="flex items-center justify-between">
                  <h4 className="text-[11px] font-bold text-slate-400 uppercase tracking-widest flex items-center gap-2">
                    <Layers className="w-4 h-4 text-indigo-400" />
                    Multi-Page Queue ({capturedPages.length})
                  </h4>
                  {capturedPages.length > 0 && (
                    <button 
                      onClick={restartCapturing}
                      className="text-[10px] text-rose-400 hover:text-rose-300 font-medium underline"
                    >
                      Clear All
                    </button>
                  )}
                </div>

                {capturedPages.length === 0 ? (
                  <div className="border border-dashed border-slate-800 rounded-2xl p-8 text-center text-slate-500 flex flex-col items-center justify-center h-[240px]">
                    <Camera className="w-8 h-8 text-slate-600 mb-2.5" />
                    <p className="text-xs">Queue is empty</p>
                    <p className="text-[10px] text-slate-600 mt-1">Snap some sheets to queue multiple pages</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-2 gap-3 max-h-[300px] md:max-h-none overflow-y-auto pr-1">
                    {capturedPages.map((page, idx) => (
                      <div 
                        key={page.id}
                        className={`relative rounded-xl border overflow-hidden aspect-[3/4] cursor-pointer group bg-slate-950 transition-all ${
                          idx === activePageIndex 
                            ? 'border-indigo-500 ring-2 ring-indigo-500/20 scale-[0.98]' 
                            : 'border-slate-800 hover:border-slate-700'
                        }`}
                        onClick={() => setActivePageIndex(idx)}
                      >
                        {/* eslint-disable-next-line @next/next/no-img-element */}
                        <img 
                          src={page.processedUrl} 
                          alt={`Page ${idx + 1}`} 
                          className="w-full h-full object-cover" 
                        />
                        <div className="absolute top-2 left-2 bg-slate-950/80 px-2 py-0.5 rounded-lg text-[9px] font-mono border border-slate-800 text-slate-300">
                          P. {idx + 1}
                        </div>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            deletePage(idx);
                          }}
                          className="absolute bottom-2 right-2 p-1.5 bg-rose-950/90 text-rose-400 hover:text-rose-200 rounded-lg opacity-0 group-hover:opacity-100 transition border border-rose-800/40"
                          title="Delete Page"
                        >
                          <Trash className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Progress and Run Button */}
              <div className="pt-4 border-t border-slate-800/60 mt-4 shrink-0">
                <button
                  onClick={handleRunSequentialOcr}
                  disabled={capturedPages.length === 0}
                  className="w-full py-4 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold rounded-2xl transition flex items-center justify-center gap-2.5 shadow-lg shadow-indigo-600/20 cursor-pointer animate-pulse"
                >
                  <Sparkles className="w-4 h-4 text-indigo-200" />
                  Analyze {capturedPages.length} {capturedPages.length === 1 ? 'Page' : 'Pages'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* VIEWMODE: PROCESSING (Loader with animations) */}
        {viewMode === 'processing' && (
          <div className="flex-1 bg-slate-950 flex flex-col items-center justify-center p-8 relative">
            <div className="w-full max-w-md space-y-6 text-center z-10">
              <div className="relative w-24 h-24 mx-auto">
                <motion.div 
                  className="absolute inset-0 rounded-full border-4 border-dashed border-indigo-500/20"
                  animate={{ rotate: 360 }}
                  transition={{ repeat: Infinity, duration: 12, ease: 'linear' }}
                />
                <motion.div 
                  className="absolute inset-2 rounded-full border-4 border-t-indigo-500 border-r-transparent border-b-transparent border-l-transparent"
                  animate={{ rotate: -360 }}
                  transition={{ repeat: Infinity, duration: 2, ease: 'linear' }}
                />
                <div className="absolute inset-4 bg-slate-900 rounded-full flex items-center justify-center border border-slate-800">
                  <Loader2 className="w-7 h-7 text-indigo-400 animate-spin" />
                </div>
              </div>

              <div className="space-y-2">
                <h4 className="font-bold text-slate-100 text-base">{processingStage}</h4>
                <p className="text-xs text-slate-500 font-mono">Running asynchronous multi-threaded cloud OCR</p>
              </div>

              {/* Progress Bar */}
              <div className="space-y-1.5">
                <div className="w-full bg-slate-900 rounded-full h-2.5 overflow-hidden border border-slate-800/50">
                  <motion.div 
                    className="bg-indigo-500 h-full rounded-full"
                    animate={{ width: `${processingProgress}%` }}
                    transition={{ duration: 0.3 }}
                  />
                </div>
                <div className="flex justify-between text-[10px] font-mono text-slate-500">
                  <span>SYSTEM_OCR_EXEC</span>
                  <span>{processingProgress}%</span>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* VIEWMODE: REVIEWING (Split Workspace) */}
        {viewMode === 'reviewing' && (
          <div className="flex-1 flex flex-col lg:flex-row overflow-hidden bg-slate-950">
            
            {/* Left Column: Image Reviewer with Enhancement Controls */}
            <div className="w-full lg:w-[45%] border-r border-slate-800 flex flex-col justify-between overflow-hidden bg-slate-900/20 p-5">
              
              {/* Image Controls toolbar */}
              <div className="flex items-center justify-between bg-slate-950 p-2.5 rounded-2xl border border-slate-800/80 mb-4 text-xs shrink-0">
                <div className="flex items-center gap-1">
                  <button
                    onClick={() => rotatePageClockwise(activePageIndex)}
                    className="p-2 hover:bg-slate-800 hover:text-white rounded-lg text-slate-400 transition"
                    title="Rotate clockwise"
                  >
                    <RotateCw className="w-4 h-4" />
                  </button>

                  <div className="h-4 w-px bg-slate-800 mx-1" />

                  <button
                    onClick={() => toggleContrastBooster(activePageIndex)}
                    className={`px-3 py-1.5 rounded-lg font-semibold flex items-center gap-1.5 transition ${
                      capturedPages[activePageIndex]?.contrastBoosted 
                        ? 'bg-indigo-500/20 text-indigo-300 border border-indigo-500/30' 
                        : 'text-slate-400 hover:text-white hover:bg-slate-800'
                    }`}
                    title="Enhance light/dark contrasts"
                  >
                    <Sliders className="w-3.5 h-3.5" />
                    Contrast Boost
                  </button>

                  <button
                    onClick={() => toggleTextSharpening(activePageIndex)}
                    className={`px-3 py-1.5 rounded-lg font-semibold flex items-center gap-1.5 transition ${
                      capturedPages[activePageIndex]?.sharpened 
                        ? 'bg-amber-500/20 text-amber-300 border border-amber-500/30' 
                        : 'text-slate-400 hover:text-white hover:bg-slate-800'
                    }`}
                    title="Sharpen micro fonts using 3x3 filter"
                  >
                    <Sparkles className="w-3.5 h-3.5 text-amber-400" />
                    Sharpen
                  </button>
                </div>

                {/* Zoom tools */}
                <div className="flex items-center gap-1.5 bg-slate-900 p-0.5 rounded-lg">
                  <button
                    onClick={handleZoomOut}
                    className="p-1 hover:bg-slate-800 text-slate-400 hover:text-white rounded transition"
                  >
                    <ZoomOut className="w-3.5 h-3.5" />
                  </button>
                  <span className="text-[10px] font-mono px-1 font-semibold text-slate-400">
                    {zoom}x
                  </span>
                  <button
                    onClick={handleZoomIn}
                    className="p-1 hover:bg-slate-800 text-slate-400 hover:text-white rounded transition"
                  >
                    <ZoomIn className="w-3.5 h-3.5" />
                  </button>
                  {zoom > 1 && (
                    <button
                      onClick={handleZoomReset}
                      className="text-[9px] text-indigo-400 hover:text-indigo-300 underline font-mono px-1.5"
                    >
                      Reset
                    </button>
                  )}
                </div>
              </div>

              {/* Pan-Zoom Interactive Image Stage */}
              <div 
                className={`flex-1 border border-slate-800 bg-slate-950/80 rounded-3xl overflow-hidden relative flex items-center justify-center shadow-inner ${
                  zoom > 1 ? 'cursor-grab active:cursor-grabbing' : ''
                }`}
                onMouseDown={handleMouseDown}
                onMouseMove={handleMouseMove}
                onMouseUp={handleMouseUp}
                onMouseLeave={handleMouseUp}
              >
                <div
                  style={{
                    transform: `translate(${pan.x}px, ${pan.y}px) scale(${zoom})`,
                    transition: isDragging ? 'none' : 'transform 0.2s ease-out'
                  }}
                  className="w-full h-full flex items-center justify-center p-4 pointer-events-none"
                >
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img
                    src={capturedPages[activePageIndex]?.processedUrl}
                    alt="Active Review page"
                    className="max-w-full max-h-full object-contain shadow-2xl rounded"
                    draggable={false}
                  />
                </div>

                {zoom > 1 && (
                  <div className="absolute bottom-4 right-4 bg-slate-950/80 px-3 py-1.5 rounded-lg border border-slate-800 text-[9px] text-slate-400 tracking-wider flex items-center gap-1">
                    <Maximize2 className="w-3 h-3" />
                    Drag to pan image details
                  </div>
                )}
              </div>

              {/* Page Selector Strip under Image */}
              {capturedPages.length > 1 && (
                <div className="mt-4 flex gap-2 shrink-0 overflow-x-auto py-1">
                  {capturedPages.map((page, idx) => (
                    <button
                      key={page.id}
                      onClick={() => {
                        setActivePageIndex(idx);
                        handleZoomReset();
                      }}
                      className={`relative rounded-lg overflow-hidden border aspect-[3/4] w-12 h-16 shrink-0 bg-slate-950 transition-all ${
                        idx === activePageIndex 
                          ? 'border-indigo-500 scale-[0.98] ring-2 ring-indigo-500/10' 
                          : 'border-slate-800 opacity-60 hover:opacity-100'
                      }`}
                    >
                      {/* eslint-disable-next-line @next/next/no-img-element */}
                      <img src={page.processedUrl} alt="" className="w-full h-full object-cover" />
                      <div className="absolute inset-0 bg-black/20 flex items-center justify-center text-[9px] font-mono font-bold text-white shadow-inner">
                        {idx + 1}
                      </div>
                    </button>
                  ))}
                </div>
              )}
            </div>

            {/* Right Column: Editable data tables / Verification Workspace */}
            <div className="flex-1 flex flex-col justify-between overflow-hidden bg-slate-950 p-5 border-t lg:border-t-0 border-slate-800">
              
              {/* Interactive Inspector Header */}
              <div className="flex items-center justify-between pb-4 border-b border-slate-800/80 shrink-0">
                <div className="flex items-center gap-2">
                  <div className="p-1.5 bg-emerald-500/10 border border-emerald-500/20 rounded-lg text-emerald-400">
                    <CheckCircle2 className="w-4 h-4" />
                  </div>
                  <div>
                    <h4 className="font-bold text-xs text-slate-100 uppercase tracking-wider">
                      Verify & Edit Extracted Fields
                    </h4>
                    <p className="text-[10px] text-slate-400">Check values before final importation</p>
                  </div>
                </div>

                {mode === 'table' && (
                  <button
                    onClick={addTableRow}
                    className="px-3 py-1.5 bg-indigo-600/10 border border-indigo-500/20 hover:bg-indigo-600/20 text-indigo-400 text-[10px] font-bold rounded-lg transition flex items-center gap-1 cursor-pointer"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    Add Row
                  </button>
                )}
              </div>

              {/* Main verification scroll container */}
              <div className="flex-1 overflow-y-auto py-4">
                
                {mode === 'table' ? (
                  /* TABULAR REVIEWING WORKSPACE (BDO/SBH) */
                  <div className="space-y-4">
                    {ocrTableRows.length === 0 ? (
                      <div className="p-12 text-center text-slate-500">
                        <p className="text-xs">No rows extracted</p>
                      </div>
                    ) : (
                      <div className="overflow-x-auto border border-slate-800/80 rounded-2xl bg-slate-900/10">
                        <table className="w-full text-xs text-left border-collapse">
                          <thead className="bg-slate-900/60 text-slate-400 font-mono text-[9px] uppercase tracking-wider border-b border-slate-800">
                            <tr>
                              <th className="px-3 py-3 w-[10%]">Flight</th>
                              <th className="px-3 py-3 w-[12%]">Date</th>
                              <th className="px-3 py-3 w-[22%]">Baggage Tag</th>
                              <th className="px-3 py-3 w-[24%]">Passenger Name</th>
                              <th className="px-3 py-3 w-[10%]">Wt</th>
                              <th className="px-3 py-3 w-[12%]">Status</th>
                              <th className="px-3 py-3 w-[10%] text-center">Action</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-800">
                            {ocrTableRows.map((row, idx) => {
                              const isLowConfidence = row.confidence === 'low' || (row.lowConfidenceFields?.length || 0) > 0;
                              return (
                                <tr 
                                  key={idx}
                                  className={`transition-colors ${
                                    isLowConfidence 
                                      ? 'bg-rose-500/5 hover:bg-rose-500/10' 
                                      : 'hover:bg-slate-900/40'
                                  }`}
                                >
                                  {/* Flight No */}
                                  <td className="px-2 py-2">
                                    <input 
                                      type="text"
                                      value={row.flightNo}
                                      onChange={(e) => updateTableCell(idx, 'flightNo', e.target.value)}
                                      className={`w-full bg-slate-950/60 border rounded px-2 py-1 font-mono text-xs uppercase text-slate-300 outline-none ${
                                        row.lowConfidenceFields?.includes('flightNo') ? 'border-rose-500/50 focus:border-rose-500' : 'border-slate-800 focus:border-indigo-500'
                                      }`}
                                    />
                                  </td>
                                  
                                  {/* Date */}
                                  <td className="px-2 py-2">
                                    <input 
                                      type="text"
                                      value={row.receivedAt}
                                      onChange={(e) => updateTableCell(idx, 'receivedAt', e.target.value)}
                                      className="w-full bg-slate-950/60 border border-slate-800 focus:border-indigo-500 rounded px-2 py-1 text-xs text-slate-300 outline-none"
                                    />
                                  </td>

                                  {/* Baggage Tag */}
                                  <td className="px-2 py-2">
                                    <div className="relative">
                                      <input 
                                        type="text"
                                        value={row.originalTag}
                                        onChange={(e) => updateTableCell(idx, 'originalTag', e.target.value)}
                                        className={`w-full bg-slate-950/60 border rounded pl-2 pr-6 py-1 font-mono text-xs text-slate-300 outline-none ${
                                          row.lowConfidenceFields?.includes('originalTag') || row.originalTag.length < 10
                                            ? 'border-rose-500/50 focus:border-rose-500' 
                                            : 'border-slate-800 focus:border-indigo-500'
                                        }`}
                                      />
                                      {isLowConfidence && row.lowConfidenceFields?.includes('originalTag') && (
                                        <AlertTriangle className="absolute right-2 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-rose-500 animate-pulse" />
                                      )}
                                    </div>
                                  </td>

                                  {/* Passenger Name */}
                                  <td className="px-2 py-2">
                                    <input 
                                      type="text"
                                      value={row.name}
                                      onChange={(e) => updateTableCell(idx, 'name', e.target.value)}
                                      className={`w-full bg-slate-950/60 border rounded px-2 py-1 text-xs text-slate-300 outline-none ${
                                        row.lowConfidenceFields?.includes('name') ? 'border-rose-500/50 focus:border-rose-500' : 'border-slate-800 focus:border-indigo-500'
                                      }`}
                                    />
                                  </td>

                                  {/* Weight */}
                                  <td className="px-2 py-2">
                                    <input 
                                      type="text"
                                      value={row.weight}
                                      onChange={(e) => updateTableCell(idx, 'weight', e.target.value)}
                                      className="w-full bg-slate-950/60 border border-slate-800 focus:border-indigo-500 rounded px-2 py-1 font-mono text-xs text-slate-300 outline-none"
                                    />
                                  </td>

                                  {/* Status */}
                                  <td className="px-2 py-2">
                                    <input 
                                      type="text"
                                      value={row.rushTag}
                                      onChange={(e) => updateTableCell(idx, 'rushTag', e.target.value)}
                                      className="w-full bg-slate-950/60 border border-slate-800 focus:border-indigo-500 rounded px-2 py-1 text-xs text-slate-300 outline-none"
                                    />
                                  </td>

                                  {/* Delete */}
                                  <td className="px-2 py-2 text-center">
                                    <button
                                      onClick={() => deleteTableRow(idx)}
                                      className="p-1.5 hover:bg-slate-800 text-slate-500 hover:text-rose-400 rounded transition cursor-pointer"
                                      title="Remove Row"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" />
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                ) : (
                  /* ALTEA PLAIN TEXT CONSOLE EDITOR */
                  <div className="w-full h-full min-h-[350px] bg-slate-950 border border-slate-800 rounded-2xl overflow-hidden flex flex-col">
                    <textarea
                      value={ocrTextResult}
                      onChange={(e) => setOcrTextResult(e.target.value)}
                      className="flex-1 p-4 bg-transparent border-none focus:ring-0 text-slate-200 font-mono text-xs leading-relaxed outline-none resize-none h-[400px]"
                      placeholder="Fine-tune Altea output lines..."
                    />
                  </div>
                )}
              </div>

              {/* Bottom footer buttons */}
              <div className="pt-4 border-t border-slate-800/80 flex flex-col sm:flex-row gap-3 shrink-0">
                <button
                  onClick={restartCapturing}
                  className="flex-1 bg-slate-900 hover:bg-slate-800 text-slate-400 hover:text-slate-200 font-bold py-3.5 rounded-2xl border border-slate-800 transition flex items-center justify-center gap-2 cursor-pointer"
                >
                  <RefreshCw className="w-4 h-4" />
                  Discard & Snap Again
                </button>
                <button
                  onClick={commitImport}
                  className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white font-bold py-3.5 rounded-2xl transition flex items-center justify-center gap-2 shadow-lg shadow-indigo-600/20 cursor-pointer"
                >
                  <CheckCircle2 className="w-5 h-5 text-indigo-100" />
                  Confirm & Complete Import
                </button>
              </div>

            </div>

          </div>
        )}

      </div>

      {/* HIDDEN HELPERS */}
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={handleFileChange}
      />
      <canvas ref={canvasRef} className="hidden" />

    </div>
  );
}
