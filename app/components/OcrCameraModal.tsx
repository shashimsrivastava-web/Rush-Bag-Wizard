'use client';

import React, { useEffect, useState, useRef } from 'react';
import { motion } from 'motion/react';
import { 
  X, Camera, RefreshCw, AlertCircle, Upload, Check, Zap, ZapOff,
  Bug, Terminal, FileText, CheckCircle2, ShieldAlert, Image as ImageIcon, Info, Play, Loader2, BarChart2
} from 'lucide-react';

interface OcrCameraModalProps {
  isOpen: boolean;
  onClose: () => void;
  onCapture: (base64Image: string) => void;
  mode: 'altea' | 'table';
}

function applyConvolution(canvas: HTMLCanvasElement, kernel: number[], mix: number = 1.0) {
  const ctx = canvas.getContext('2d');
  if (!ctx) return;
  const width = canvas.width;
  const height = canvas.height;
  const imageData = ctx.getImageData(0, 0, width, height);
  const pixels = imageData.data;
  const side = Math.round(Math.sqrt(kernel.length));
  const halfSide = Math.floor(side / 2);
  
  // Output image data
  const output = ctx.createImageData(width, height);
  const outPixels = output.data;
  
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const sy = y;
      const sx = x;
      const dstOff = (y * width + x) * 4;
      
      let r = 0, g = 0, b = 0;
      for (let cy = 0; cy < side; cy++) {
        for (let cx = 0; cx < side; cx++) {
          const scy = sy + cy - halfSide;
          const scx = sx + cx - halfSide;
          if (scy >= 0 && scy < height && scx >= 0 && scx < width) {
            const srcOff = (scy * width + scx) * 4;
            const wt = kernel[cy * side + cx];
            r += pixels[srcOff] * wt;
            g += pixels[srcOff + 1] * wt;
            b += pixels[srcOff + 2] * wt;
          }
        }
      }
      
      outPixels[dstOff] = Math.min(255, Math.max(0, r * mix + pixels[dstOff] * (1 - mix)));
      outPixels[dstOff + 1] = Math.min(255, Math.max(0, g * mix + pixels[dstOff + 1] * (1 - mix)));
      outPixels[dstOff + 2] = Math.min(255, Math.max(0, b * mix + pixels[dstOff + 2] * (1 - mix)));
      outPixels[dstOff + 3] = pixels[dstOff + 3]; // Alpha
    }
  }
  ctx.putImageData(output, 0, 0);
}

export default function OcrCameraModal({ isOpen, onClose, onCapture, mode }: OcrCameraModalProps) {
  // Camera & Device states
  const [stream, setStream] = useState<MediaStream | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [facingMode, setFacingMode] = useState<'environment' | 'user'>('environment');
  const [devices, setDevices] = useState<MediaDeviceInfo[]>([]);
  const [selectedDeviceId, setSelectedDeviceId] = useState<string>('');
  const [isInitializing, setIsInitializing] = useState(false);
  const [torchSupported, setTorchSupported] = useState(false);
  const [torchOn, setTorchOn] = useState(false);
  const [cameraTrigger, setCameraTrigger] = useState(0);

  // Captured / Preprocessed Image states
  const [capturedImage, setCapturedImage] = useState<string | null>(null);
  const [stageOriginal, setStageOriginal] = useState<string | null>(null);
  const [stageAutoEnhance, setStageAutoEnhance] = useState<string | null>(null);
  const [stageContrast, setStageContrast] = useState<string | null>(null);
  const [stageSharpen, setStageSharpen] = useState<string | null>(null);
  const [stageOcrInput, setStageOcrInput] = useState<string | null>(null);
  const [selectedPreviewStage, setSelectedPreviewStage] = useState<'original' | 'auto_enhance' | 'contrast' | 'sharpen' | 'ocr_input'>('original');

  // Diagnostics Dashboard States
  const [isDebugMode, setIsDebugMode] = useState(false);
  const [logs, setLogs] = useState<string[]>([]);
  const [dimensions, setDimensions] = useState<{ width: number; height: number } | null>(null);
  const [charCount, setCharCount] = useState<number>(0);
  const [ocrDuration, setOcrDuration] = useState<number>(0);
  const [lineCount, setLineCount] = useState<number>(0);
  const [rawOcrText, setRawOcrText] = useState<string>('');
  const [isDiagnosticOcrRunning, setIsDiagnosticOcrRunning] = useState(false);
  const [diagnosticError, setDiagnosticError] = useState<string | null>(null);
  const [stages, setStages] = useState<Record<string, 'pending' | 'running' | 'success' | 'failed'>>({
    cameraCapture: 'pending',
    imageLoaded: 'pending',
    imageDimensions: 'pending',
    imageEnhancement: 'pending',
    ocrStarted: 'pending',
    ocrFinished: 'pending',
    parserStarted: 'pending',
    parserFinished: 'pending',
  });

  const [systemDiagnostics, setSystemDiagnostics] = useState<{
    userAgent: string;
    secureContext: boolean;
    cameraSupported: boolean;
    canvasSupported: boolean;
  }>({
    userAgent: '',
    secureContext: false,
    cameraSupported: false,
    canvasSupported: false,
  });

  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Helper log function
  const addLog = (msg: string) => {
    setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);
  };

  // Stop camera stream
  const stopCamera = () => {
    if (stream) {
      stream.getTracks().forEach((track) => track.stop());
      setStream(null);
    }
    setTorchOn(false);
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

  // System & Browser diagnostics checking
  useEffect(() => {
    if (typeof window !== 'undefined' && isOpen) {
      setTimeout(() => {
        setSystemDiagnostics({
          userAgent: navigator.userAgent,
          secureContext: window.isSecureContext,
          cameraSupported: !!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia),
          canvasSupported: !!document.createElement('canvas').getContext,
        });
      }, 0);
    }
  }, [isOpen]);

  // Main camera stream lifecycle effect
  useEffect(() => {
    let active = true;

    const startCamera = async () => {
      if (!isOpen) return;
      setIsInitializing(true);
      setError(null);

      try {
        await getDevices();

        const constraints: MediaStreamConstraints = {
          video: selectedDeviceId 
            ? { deviceId: { exact: selectedDeviceId } }
            : { facingMode: facingMode }
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
        console.error('Error accessing camera:', err);
        if (active) {
          if (facingMode === 'environment' && !selectedDeviceId) {
            try {
              const fallbackStream = await navigator.mediaDevices.getUserMedia({ video: true });
              if (active) {
                setStream(fallbackStream);
              } else {
                fallbackStream.getTracks().forEach((track) => track.stop());
              }
            } catch (fallbackErr) {
              setError('Camera access denied or unavailable. Please upload a file instead.');
            }
          } else {
            setError('Camera access denied or unavailable. Please upload a file instead.');
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
  }, [isOpen, facingMode, selectedDeviceId, cameraTrigger]);

  // Bind media stream to video tag
  useEffect(() => {
    if (videoRef.current && stream) {
      videoRef.current.srcObject = stream;
    }
  }, [stream]);

  // Reset states on open/close
  useEffect(() => {
    if (isOpen) {
      const timer = setTimeout(() => {
        setCapturedImage(null);
        setStageOriginal(null);
        setStageAutoEnhance(null);
        setStageContrast(null);
        setStageSharpen(null);
        setStageOcrInput(null);
        setRawOcrText('');
        setDiagnosticError(null);
        setLogs([]);
        setStages({
          cameraCapture: 'pending',
          imageLoaded: 'pending',
          imageDimensions: 'pending',
          imageEnhancement: 'pending',
          ocrStarted: 'pending',
          ocrFinished: 'pending',
          parserStarted: 'pending',
          parserFinished: 'pending',
        });
        setSelectedPreviewStage('original');
      }, 0);
      return () => clearTimeout(timer);
    } else {
      setTimeout(() => {
        stopCamera();
      }, 0);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen]);

  // Preprocessing pipeline
  const runPreprocessing = (base64: string) => {
    addLog("Starting image load stage...");
    setStages(prev => ({ ...prev, imageLoaded: 'running' }));
    
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      addLog(`Image loaded successfully. Dimensions: ${img.width}x${img.height}`);
      setDimensions({ width: img.width, height: img.height });
      setStages(prev => ({ 
        ...prev, 
        imageLoaded: 'success',
        imageDimensions: 'success',
        imageEnhancement: 'running'
      }));

      // Create intermediate canvas for processing
      const canvas = document.createElement('canvas');
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        addLog("[ERROR] Failed to obtain 2D rendering context.");
        setStages(prev => ({ ...prev, imageEnhancement: 'failed' }));
        return;
      }

      // Stage 1: Original
      setStageOriginal(base64);

      // Stage 2: Grayscale / Auto-enhance
      ctx.drawImage(img, 0, 0);
      ctx.filter = "grayscale(100%) brightness(1.1) contrast(1.1)";
      ctx.drawImage(canvas, 0, 0);
      const stage2Url = canvas.toDataURL('image/jpeg', 0.9);
      setStageAutoEnhance(stage2Url);
      addLog("Auto Crop & Grayscale Enhancement completed.");

      // Stage 3: Contrast Enhancement
      ctx.filter = "contrast(1.5) brightness(1.15)";
      ctx.drawImage(canvas, 0, 0);
      const stage3Url = canvas.toDataURL('image/jpeg', 0.9);
      setStageContrast(stage3Url);
      addLog("Contrast Enhancement (150% boost) completed.");

      // Stage 4: Sharpen
      ctx.filter = "none"; // Reset filter
      const contrastImg = new Image();
      contrastImg.onload = () => {
        ctx.drawImage(contrastImg, 0, 0);
        applyConvolution(canvas, [
           0, -1,  0,
          -1,  5, -1,
           0, -1,  0
        ]);
        const stage4Url = canvas.toDataURL('image/jpeg', 0.9);
        setStageSharpen(stage4Url);
        addLog("Sharpen Convolution Filtering (high-frequency pass) completed.");

        // Stage 5: Final OCR Input Image
        setStageOcrInput(stage4Url);
        setStages(prev => ({ ...prev, imageEnhancement: 'success' }));
        addLog("Preprocessing stages successfully compiled.");
      };
      contrastImg.src = stage3Url;
    };
    img.onerror = () => {
      addLog("[ERROR] Image loading failed. The base64 source may be corrupt.");
      setStages(prev => ({ ...prev, imageLoaded: 'failed', imageEnhancement: 'failed' }));
    };
    img.src = base64;
  };

  // Capture still photo
  const capturePhoto = () => {
    if (videoRef.current && canvasRef.current) {
      const video = videoRef.current;
      const canvas = canvasRef.current;
      const ctx = canvas.getContext('2d');

      if (ctx) {
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
        setCapturedImage(dataUrl);
        setStages(prev => ({ ...prev, cameraCapture: 'success' }));
        addLog("Camera capture successful.");
        stopCamera();

        if (isDebugMode) {
          runPreprocessing(dataUrl);
        }
      }
    }
  };

  // Handle manual file selection (as fallback)
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result as string;
        setCapturedImage(dataUrl);
        setStages(prev => ({ ...prev, cameraCapture: 'success' }));
        addLog("File imported successfully from local device.");
        
        if (isDebugMode) {
          runPreprocessing(dataUrl);
        }
      };
      reader.readAsDataURL(file);
    }
  };

  // Switch between front and back camera
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

  const handleConfirm = () => {
    if (capturedImage) {
      const finalImage = isDebugMode && stageOcrInput ? stageOcrInput : capturedImage;
      onCapture(finalImage);
      onClose();
    }
  };

  // Trigger diagnostic OCR call directly inside the modal
  const runOcrDiagnostics = async (testOnly: boolean) => {
    if (!capturedImage) return;
    setIsDiagnosticOcrRunning(true);
    setDiagnosticError(null);
    setRawOcrText('');
    
    addLog(`[OCR] Initiating OCR request to Gemini 3.5 Flash (Mode: ${testOnly ? 'ocr_only' : mode})...`);
    
    setStages(prev => ({ 
      ...prev, 
      ocrStarted: 'running', 
      ocrFinished: 'pending', 
      parserStarted: 'pending', 
      parserFinished: 'pending' 
    }));

    const startTime = Date.now();
    try {
      const imgToSend = stageOcrInput || capturedImage;
      const response = await fetch('/api/ocr', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ image: imgToSend, mode: testOnly ? 'ocr_only' : mode }),
      });
      
      const duration = ((Date.now() - startTime) / 1000).toFixed(2);
      setOcrDuration(Number(duration));

      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }

      const result = await response.json();
      
      if (result.error) {
        throw new Error(result.error);
      }

      setStages(prev => ({ ...prev, ocrFinished: 'success' }));
      addLog(`[OCR] Response received in ${duration} seconds.`);
      
      let rawText = '';
      if (testOnly) {
        rawText = result.text || '';
        setRawOcrText(rawText);
        setCharCount(rawText.length);
        setLineCount(rawText ? rawText.split('\n').length : 0);
        
        if (!rawText || rawText.trim().length === 0) {
          addLog('[OCR] WARNING: OCR returned 0 characters.');
          setStages(prev => ({ ...prev, ocrFinished: 'failed' }));
        } else {
          addLog(`[OCR] Success: Extracted ${rawText.length} characters.`);
        }
      } else {
        if (result.data) {
          rawText = JSON.stringify(result.data, null, 2);
          setRawOcrText(rawText);
          setCharCount(rawText.length);
          setLineCount(result.data.length || 0);
          
          addLog(`[OCR] Success: Parsed ${result.data.length} structured records.`);
          setStages(prev => ({ ...prev, parserStarted: 'success', parserFinished: 'success' }));
        } else if (result.text) {
          rawText = result.text || '';
          setRawOcrText(rawText);
          setCharCount(rawText.length);
          setLineCount(rawText ? rawText.split('\n').length : 0);
          
          addLog(`[OCR] Partial success: Got text, but parsing failed.`);
          setStages(prev => ({ ...prev, parserStarted: 'failed', parserFinished: 'failed' }));
        } else {
          throw new Error('Response payload has empty text and empty structured data.');
        }
      }
    } catch (err: any) {
      addLog(`[ERROR] OCR pipeline failed: ${err.message || err}`);
      setDiagnosticError(err.message || 'Unknown OCR Error');
      setStages(prev => ({ ...prev, ocrFinished: 'failed', parserFinished: 'failed' }));
    } finally {
      setIsDiagnosticOcrRunning(false);
    }
  };

  const renderStatusIcon = (status: 'pending' | 'running' | 'success' | 'failed') => {
    switch (status) {
      case 'success':
        return <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />;
      case 'failed':
        return <AlertCircle className="w-4 h-4 text-rose-500 shrink-0" />;
      case 'running':
        return <Loader2 className="w-4 h-4 text-indigo-400 animate-spin shrink-0" />;
      default:
        return <div className="w-3.5 h-3.5 border border-slate-700 rounded-full shrink-0" />;
    }
  };

  if (!isOpen) return null;

  return (
    <div className={`fixed inset-0 bg-slate-950/98 backdrop-blur-md z-[150] flex flex-col items-center p-4 md:p-6 overflow-y-auto select-none ${isDebugMode ? 'justify-start' : 'justify-between'}`}>
      {/* Header */}
      <div className={`w-full flex items-center justify-between z-10 shrink-0 mb-4 ${isDebugMode ? 'max-w-4xl' : 'max-w-lg'}`}>
        <div>
          <h3 className="text-white font-bold text-lg flex items-center gap-2">
            <Camera className="w-5 h-5 text-indigo-400" />
            {mode === 'altea' ? 'Capture Altea Screen' : 'Capture Excel / Baggage Sheet'}
          </h3>
          <p className="text-xs text-slate-400">
            {capturedImage ? 'Review your photo before analyzing' : 'Ensure all columns and tags are clearly visible'}
          </p>
        </div>
        <button
          onClick={onClose}
          className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-full transition"
        >
          <X className="w-5 h-5" />
        </button>
      </div>

      {/* Main Container */}
      <div className={`w-full flex flex-col items-center gap-4 ${isDebugMode ? 'max-w-4xl' : 'max-w-lg'}`}>
        
        {/* Toggle standard / debug view */}
        {capturedImage && (
          <div className="flex border border-slate-800 bg-slate-900/60 p-1 rounded-xl w-full self-stretch shrink-0">
            <button
              onClick={() => setIsDebugMode(false)}
              className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition ${
                !isDebugMode ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <Check className="w-3.5 h-3.5" />
              Standard Import View
            </button>
            <button
              onClick={() => {
                setIsDebugMode(true);
                if (capturedImage && !stageOcrInput) {
                  runPreprocessing(capturedImage);
                }
              }}
              className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition ${
                isDebugMode ? 'bg-indigo-600/20 text-indigo-400 border border-indigo-500/30' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <Bug className="w-3.5 h-3.5" />
              Developer Diagnostics Mode
            </button>
          </div>
        )}

        {/* Dynamic Display Layout */}
        {!isDebugMode ? (
          /* STANDARD CAMERA VIEW OR IMAGE PREVIEW */
          <div className="relative w-full aspect-[3/4] max-h-[60vh] bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden flex items-center justify-center shadow-2xl">
            {!capturedImage ? (
              <>
                {isInitializing && (
                  <div className="absolute inset-0 flex flex-col items-center justify-center bg-slate-950/60 backdrop-blur-sm z-20">
                    <Loader2 className="w-10 h-10 text-indigo-500 animate-spin mb-3" />
                    <span className="text-xs text-indigo-400 font-mono">Initializing camera...</span>
                  </div>
                )}

                {error ? (
                  <div className="p-6 text-center max-w-sm flex flex-col items-center gap-3">
                    <AlertCircle className="w-12 h-12 text-rose-500 animate-bounce" />
                    <p className="text-slate-300 text-sm font-medium">{error}</p>
                    <button
                      onClick={() => fileInputRef.current?.click()}
                      className="mt-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold px-4 py-2.5 rounded-xl flex items-center gap-2 mx-auto transition"
                    >
                      <Upload className="w-4 h-4" />
                      Select Photo from Device
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
                    
                    {/* Visual guideline overlay */}
                    <div className="absolute inset-0 border-[24px] border-slate-950/40 pointer-events-none flex items-center justify-center">
                      <div className="w-full h-full border border-dashed border-indigo-400/60 rounded-lg relative">
                        <div className="absolute left-0 right-0 h-0.5 bg-gradient-to-r from-transparent via-indigo-400 to-transparent top-1/2 -translate-y-1/2 animate-pulse" />
                        <div className="absolute top-0 left-0 w-4 h-4 border-t-2 border-l-2 border-indigo-400" />
                        <div className="absolute top-0 right-0 w-4 h-4 border-t-2 border-r-2 border-indigo-400" />
                        <div className="absolute bottom-0 left-0 w-4 h-4 border-b-2 border-l-2 border-indigo-400" />
                        <div className="absolute bottom-0 right-0 w-4 h-4 border-b-2 border-r-2 border-indigo-400" />
                        
                        <div className="absolute bottom-4 left-1/2 -translate-x-1/2 bg-slate-950/80 px-3 py-1 rounded-full border border-slate-800 text-[10px] text-indigo-300 font-mono tracking-wider uppercase">
                          {mode === 'altea' ? 'Altea Screen Cryptic View' : 'Baggage Column Grid'}
                        </div>
                      </div>
                    </div>
                  </>
                )}
              </>
            ) : (
              // eslint-disable-next-line @next/next/no-img-element
              <img src={capturedImage} alt="Captured preview" className="w-full h-full object-contain" />
            )}
          </div>
        ) : (
          /* DEVELOPER DIAGNOSTICS DETAILED SPLIT BENTO GRID */
          <div className="w-full grid grid-cols-1 md:grid-cols-2 gap-4 text-left animate-in fade-in duration-200">
            
            {/* Left Column: Image Review & Raw Output */}
            <div className="flex flex-col gap-4">
              
              {/* Active Image View Card */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 flex flex-col gap-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-mono font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-1.5">
                    <ImageIcon className="w-4 h-4" />
                    Interactive Image Previews
                  </span>
                  <span className="text-[10px] bg-indigo-500/10 text-indigo-300 px-2.5 py-0.5 rounded-full font-mono border border-indigo-500/20">
                    {dimensions ? `${dimensions.width} x ${dimensions.height} px` : 'Calculating...'}
                  </span>
                </div>

                <div className="relative aspect-[4/3] bg-slate-950 border border-slate-800 rounded-xl overflow-hidden flex items-center justify-center">
                  {selectedPreviewStage === 'original' && stageOriginal && (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img src={stageOriginal} alt="Original captured" className="w-full h-full object-contain" />
                  )}
                  {selectedPreviewStage === 'auto_enhance' && stageAutoEnhance && (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img src={stageAutoEnhance} alt="Grayscale autoenhance" className="w-full h-full object-contain" />
                  )}
                  {selectedPreviewStage === 'contrast' && stageContrast && (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img src={stageContrast} alt="High contrast" className="w-full h-full object-contain" />
                  )}
                  {selectedPreviewStage === 'sharpen' && stageSharpen && (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img src={stageSharpen} alt="Sharpened" className="w-full h-full object-contain" />
                  )}
                  {selectedPreviewStage === 'ocr_input' && stageOcrInput && (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img src={stageOcrInput} alt="Final OCR input" className="w-full h-full object-contain" />
                  )}
                  {!stageOriginal && (
                    <div className="flex flex-col items-center justify-center text-slate-500 text-xs">
                      <Loader2 className="w-6 h-6 animate-spin mb-2" />
                      Loading Preprocessing Pipelines...
                    </div>
                  )}
                </div>

                {/* Preprocessing Step Selectors */}
                <div className="grid grid-cols-5 gap-1.5">
                  {[
                    { id: 'original', label: '1. Orig', img: stageOriginal },
                    { id: 'auto_enhance', label: '2. Gray', img: stageAutoEnhance },
                    { id: 'contrast', label: '3. Cont', img: stageContrast },
                    { id: 'sharpen', label: '4. Sharp', img: stageSharpen },
                    { id: 'ocr_input', label: '5. OCR', img: stageOcrInput },
                  ].map((btn) => (
                    <button
                      key={btn.id}
                      onClick={() => setSelectedPreviewStage(btn.id as any)}
                      disabled={!btn.img}
                      className={`py-2 px-1 rounded-lg border text-[10px] font-mono transition flex flex-col items-center gap-1 ${
                        selectedPreviewStage === btn.id
                          ? 'bg-indigo-600 border-indigo-500 text-white shadow-md'
                          : btn.img
                            ? 'bg-slate-950 border-slate-800 text-slate-300 hover:bg-slate-800'
                            : 'bg-slate-950/40 border-slate-900/60 text-slate-600 cursor-not-allowed'
                      }`}
                    >
                      <span>{btn.label}</span>
                      <div className="w-full h-4 relative rounded overflow-hidden bg-slate-900 border border-slate-800">
                        {btn.img ? (
                          // eslint-disable-next-line @next/next/no-img-element
                          <img src={btn.img} alt="Thumbnail" className="w-full h-full object-cover" />
                        ) : (
                          <div className="w-full h-full bg-slate-950 flex items-center justify-center">
                            <span className="text-[6px]">...</span>
                          </div>
                        )}
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Console & Event Terminal Logs */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 flex flex-col gap-2 shrink-0">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-mono font-bold uppercase tracking-wider text-emerald-400 flex items-center gap-1.5">
                    <Terminal className="w-4 h-4 animate-pulse" />
                    Diagnostics Log Terminal
                  </span>
                  <button
                    onClick={() => setLogs([])}
                    className="text-[9px] font-mono text-slate-500 hover:text-slate-300 transition"
                  >
                    Clear Logs
                  </button>
                </div>
                <div className="bg-slate-950 border border-slate-800 p-3 rounded-xl h-44 overflow-y-auto font-mono text-[10px] leading-relaxed flex flex-col gap-1 select-text scrollbar-thin">
                  {logs.length === 0 ? (
                    <span className="text-slate-600">No events logged. Tap &quot;Run Diagnostics&quot; below to trigger logs.</span>
                  ) : (
                    logs.map((log, idx) => {
                      let color = 'text-slate-300';
                      if (log.includes('[ERROR]')) color = 'text-rose-400 font-semibold';
                      else if (log.includes('[INFO]')) color = 'text-cyan-400';
                      else if (log.includes('[OCR]')) color = 'text-indigo-400';
                      else if (log.includes('Success')) color = 'text-emerald-400';
                      return <div key={idx} className={color}>{log}</div>;
                    })
                  )}
                </div>
              </div>

            </div>

            {/* Right Column: Checklist, Statistics, Test Controls & Raw OCR text */}
            <div className="flex flex-col gap-4">
              
              {/* Stages Checklist & Diagnostics */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 flex flex-col gap-3">
                <span className="text-xs font-mono font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-1.5">
                  <CheckCircle2 className="w-4 h-4" />
                  OCR Pipeline Checklist
                </span>
                <div className="grid grid-cols-2 gap-x-4 gap-y-2.5">
                  {[
                    { key: 'cameraCapture', label: 'Camera Capture' },
                    { key: 'imageLoaded', label: 'Image Loaded' },
                    { key: 'imageDimensions', label: 'Image Dimensions' },
                    { key: 'imageEnhancement', label: 'Image Enhancement' },
                    { key: 'ocrStarted', label: 'OCR API Request' },
                    { key: 'ocrFinished', label: 'OCR Completion' },
                    { key: 'parserStarted', label: 'Parser Started' },
                    { key: 'parserFinished', label: 'Parser Verification' },
                  ].map((stage) => (
                    <div key={stage.key} className="flex items-center gap-2.5 text-xs text-slate-300">
                      {renderStatusIcon(stages[stage.key])}
                      <span className="font-medium truncate">{stage.label}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Statistics & Compatibility */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 grid grid-cols-2 gap-4">
                {/* Stats Table */}
                <div className="flex flex-col gap-2">
                  <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-1">
                    <BarChart2 className="w-3.5 h-3.5" />
                    OCR Stats
                  </span>
                  <div className="flex flex-col gap-1.5 text-[11px] font-mono">
                    <div className="flex justify-between border-b border-slate-800/40 pb-1">
                      <span className="text-slate-500">Duration:</span>
                      <span className="text-slate-200">{ocrDuration > 0 ? `${ocrDuration}s` : 'N/A'}</span>
                    </div>
                    <div className="flex justify-between border-b border-slate-800/40 pb-1">
                      <span className="text-slate-500">Characters:</span>
                      <span className="text-slate-200">{charCount > 0 ? charCount : 'N/A'}</span>
                    </div>
                    <div className="flex justify-between border-b border-slate-800/40 pb-1">
                      <span className="text-slate-500">Detected Lines:</span>
                      <span className="text-slate-200">{lineCount > 0 ? lineCount : 'N/A'}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-slate-500">Confidence:</span>
                      <span className="text-indigo-400 font-bold">{charCount > 0 ? 'High' : 'N/A'}</span>
                    </div>
                  </div>
                </div>

                {/* Compatibility Checklist */}
                <div className="flex flex-col gap-2">
                  <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-1">
                    <Info className="w-3.5 h-3.5" />
                    System Check
                  </span>
                  <div className="flex flex-col gap-1 text-[10px] font-mono">
                    <div className="flex items-center gap-1.5 text-slate-300">
                      <div className={`w-1.5 h-1.5 rounded-full ${systemDiagnostics.secureContext ? 'bg-emerald-500' : 'bg-rose-500'}`} />
                      <span>HTTPS Secure: {systemDiagnostics.secureContext ? 'YES' : 'NO'}</span>
                    </div>
                    <div className="flex items-center gap-1.5 text-slate-300">
                      <div className={`w-1.5 h-1.5 rounded-full ${systemDiagnostics.cameraSupported ? 'bg-emerald-500' : 'bg-rose-500'}`} />
                      <span>Camera API: {systemDiagnostics.cameraSupported ? 'OK' : 'N/A'}</span>
                    </div>
                    <div className="flex items-center gap-1.5 text-slate-300">
                      <div className={`w-1.5 h-1.5 rounded-full ${systemDiagnostics.canvasSupported ? 'bg-emerald-500' : 'bg-rose-500'}`} />
                      <span>Canvas GPU: {systemDiagnostics.canvasSupported ? 'OK' : 'N/A'}</span>
                    </div>
                    <div className="text-[8px] text-slate-500 truncate mt-1" title={systemDiagnostics.userAgent}>
                      UA: {systemDiagnostics.userAgent.split(' ')[0]}
                    </div>
                  </div>
                </div>
              </div>

              {/* Raw OCR Text output console */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 flex flex-col gap-2 relative">
                <span className="text-xs font-mono font-bold uppercase tracking-wider text-amber-400 flex items-center gap-1.5">
                  <FileText className="w-4 h-4" />
                  ===== RAW OCR OUTPUT =====
                </span>
                
                <textarea
                  readOnly
                  placeholder="Raw OCR output text from the Gemini model will appear here after diagnostic trigger..."
                  value={diagnosticError ? `[ERROR] OCR pipeline execution failed:\n${diagnosticError}` : rawOcrText || (isDiagnosticOcrRunning ? "Running scan, please wait..." : "")}
                  className="w-full h-36 bg-slate-950 border border-slate-800 rounded-xl p-3 font-mono text-[10px] text-slate-300 placeholder-slate-600 focus:outline-none resize-none select-text animate-pulse"
                />
                
                {rawOcrText === '' && !isDiagnosticOcrRunning && !diagnosticError && (
                  <div className="absolute inset-0 top-8 bg-slate-900/40 backdrop-blur-[1px] rounded-b-2xl flex items-center justify-center pointer-events-none">
                    <span className="text-[11px] font-mono text-slate-500 italic bg-slate-950 border border-slate-800/60 px-3 py-1 rounded-lg">OCR idle. Run test triggers below.</span>
                  </div>
                )}
              </div>

            </div>

          </div>
        )}
      </div>

      {/* Footer controls */}
      <div className={`w-full z-10 flex flex-col gap-4 mt-6 shrink-0 ${isDebugMode ? 'max-w-4xl' : 'max-w-lg'}`}>
        
        {/* Secure context / compatibility alert banner */}
        {!systemDiagnostics.secureContext && !capturedImage && (
          <div className="bg-rose-500/10 border border-rose-500/20 text-rose-300 p-3 rounded-2xl text-[11px] flex gap-2.5 items-center">
            <ShieldAlert className="w-5 h-5 text-rose-400 shrink-0" />
            <div>
              <span className="font-bold">Browser Limitation Detected</span>: Cameras require an HTTPS Secure Context to stream. Please upload an image file from your device instead.
            </div>
          </div>
        )}

        {!capturedImage ? (
          <div className="flex items-center justify-between px-4">
            {/* File fallback selector */}
            <button
              onClick={() => fileInputRef.current?.click()}
              className="p-3 bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white rounded-full transition"
              title="Upload File Instead"
            >
              <Upload className="w-5 h-5" />
            </button>

            {/* Shutter Button */}
            <button
              onClick={capturePhoto}
              disabled={isInitializing || !!error}
              className={`w-16 h-16 rounded-full border-4 border-white flex items-center justify-center transition p-1 ${
                isInitializing || !!error ? 'opacity-50 cursor-not-allowed' : 'hover:scale-105 active:scale-95 bg-indigo-500'
              }`}
            >
              <div className="w-full h-full rounded-full bg-white transition hover:bg-indigo-100" />
            </button>

            {/* Toggle front/rear camera */}
            <div className="flex gap-2">
              {torchSupported && (
                <button
                  onClick={toggleTorch}
                  className={`p-3 rounded-full transition ${
                    torchOn ? 'bg-amber-500/20 text-amber-400 border border-amber-500/30' : 'bg-slate-800/80 text-slate-300 hover:text-white'
                  }`}
                  title="Toggle Flashlight"
                >
                  {torchOn ? <Zap className="w-5 h-5" /> : <ZapOff className="w-5 h-5" />}
                </button>
              )}
              
              <button
                onClick={switchCamera}
                disabled={devices.length <= 1 && facingMode === 'environment'}
                className="p-3 bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white rounded-full transition disabled:opacity-40"
                title="Switch Camera"
              >
                <RefreshCw className="w-5 h-5" />
              </button>
            </div>
          </div>
        ) : (
          /* ACTION CONTROL CONTAINER */
          <div className="flex flex-col gap-3">
            {isDebugMode && (
              <div className="grid grid-cols-2 gap-3 shrink-0">
                <button
                  onClick={() => runOcrDiagnostics(true)}
                  disabled={isDiagnosticOcrRunning}
                  className="bg-amber-600 hover:bg-amber-500 disabled:bg-slate-800 text-white font-mono text-xs font-bold py-3 px-4 rounded-xl transition flex items-center justify-center gap-2 shadow-lg shadow-amber-600/10"
                >
                  {isDiagnosticOcrRunning ? (
                    <Loader2 className="w-4 h-4 animate-spin" />
                  ) : (
                    <Play className="w-4 h-4" />
                  )}
                  Test OCR Only (No Parse)
                </button>
                <button
                  onClick={() => runOcrDiagnostics(false)}
                  disabled={isDiagnosticOcrRunning}
                  className="bg-indigo-600 hover:bg-indigo-500 disabled:bg-slate-800 text-white font-mono text-xs font-bold py-3 px-4 rounded-xl transition flex items-center justify-center gap-2 shadow-lg shadow-indigo-600/10"
                >
                  {isDiagnosticOcrRunning ? (
                    <Loader2 className="w-4 h-4 animate-spin" />
                  ) : (
                    <Bug className="w-4 h-4" />
                  )}
                  Run Full Analysis
                </button>
              </div>
            )}

            <div className="flex gap-3 px-2">
              <button
                onClick={() => {
                  setCapturedImage(null);
                  setStageOriginal(null);
                  setStageAutoEnhance(null);
                  setStageContrast(null);
                  setStageSharpen(null);
                  setStageOcrInput(null);
                  setRawOcrText('');
                  setDiagnosticError(null);
                  setStages({
                    cameraCapture: 'pending',
                    imageLoaded: 'pending',
                    imageDimensions: 'pending',
                    imageEnhancement: 'pending',
                    ocrStarted: 'pending',
                    ocrFinished: 'pending',
                    parserStarted: 'pending',
                    parserFinished: 'pending',
                  });
                  setCameraTrigger((prev) => prev + 1);
                }}
                className="flex-1 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold py-3.5 rounded-2xl transition flex items-center justify-center gap-2"
              >
                <RefreshCw className="w-4 h-4" />
                Retake / New Photo
              </button>
              <button
                onClick={handleConfirm}
                className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white font-bold py-3.5 rounded-2xl transition shadow-lg shadow-indigo-600/20 flex items-center justify-center gap-2 group"
              >
                <Check className="w-5 h-5 group-hover:scale-125 transition-transform" />
                {isDebugMode ? 'Use Enhanced Image' : 'Analyze Photo'}
              </button>
            </div>
          </div>
        )}

        {/* Hidden inputs */}
        <input
          ref={fileInputRef}
          type="file"
          accept="image/*"
          className="hidden"
          onChange={handleFileChange}
        />
        <canvas ref={canvasRef} className="hidden" />
      </div>
    </div>
  );
}
