'use client';

import React, { useEffect, useState, useRef } from 'react';
import { motion } from 'motion/react';
import { X, Camera, RefreshCw, AlertCircle, Upload, Check, Zap, ZapOff } from 'lucide-react';

interface OcrCameraModalProps {
  isOpen: boolean;
  onClose: () => void;
  onCapture: (base64Image: string) => void;
  mode: 'altea' | 'table';
}

export default function OcrCameraModal({ isOpen, onClose, onCapture, mode }: OcrCameraModalProps) {
  const [stream, setStream] = useState<MediaStream | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [facingMode, setFacingMode] = useState<'environment' | 'user'>('environment');
  const [capturedImage, setCapturedImage] = useState<string | null>(null);
  const [devices, setDevices] = useState<MediaDeviceInfo[]>([]);
  const [selectedDeviceId, setSelectedDeviceId] = useState<string>('');
  const [isInitializing, setIsInitializing] = useState(false);
  const [torchSupported, setTorchSupported] = useState(false);
  const [torchOn, setTorchOn] = useState(false);
  const [cameraTrigger, setCameraTrigger] = useState(0);

  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

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
          // Clean up if component updated or unmounted during initialization
          mediaStream.getTracks().forEach((track) => track.stop());
        }
      } catch (err: any) {
        console.error('Error accessing camera:', err);
        if (active) {
          // Fallback if environment failed
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

  // Bind the media stream to the video element as soon as it's active
  useEffect(() => {
    if (videoRef.current && stream) {
      videoRef.current.srcObject = stream;
    }
  }, [stream]);

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

  // Capture still photo
  const capturePhoto = () => {
    if (videoRef.current && canvasRef.current) {
      const video = videoRef.current;
      const canvas = canvasRef.current;
      const ctx = canvas.getContext('2d');

      if (ctx) {
        // High quality full-resolution capture
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        
        // Mirror front camera if needed
        if (facingMode === 'user') {
          ctx.translate(canvas.width, 0);
          ctx.scale(-1, 1);
        }
        
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        
        // Reset scale
        if (facingMode === 'user') {
          ctx.setTransform(1, 0, 0, 1, 0, 0);
        }

        const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
        setCapturedImage(dataUrl);
        stopCamera();
      }
    }
  };

  // Handle manual file selection (as fallback)
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = () => {
        setCapturedImage(reader.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleConfirm = () => {
    if (capturedImage) {
      onCapture(capturedImage);
      onClose();
    }
  };

  const renderCapturedImage = () => {
    if (!capturedImage) return null;
    // eslint-disable-next-line @next/next/no-img-element
    return <img src={capturedImage} alt="Captured preview" className="w-full h-full object-contain" />;
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-slate-950/95 backdrop-blur-md z-[150] flex flex-col items-center justify-between p-4 md:p-6 select-none overflow-hidden">
      {/* Header */}
      <div className="w-full max-w-lg flex items-center justify-between z-10">
        <div>
          <h3 className="text-white font-bold text-lg flex items-center gap-2">
            <Camera className="w-5 h-5 text-indigo-400 animate-pulse" />
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

      {/* Main viewport */}
      <div className="relative w-full max-w-lg aspect-[3/4] max-h-[60vh] bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden flex items-center justify-center my-4 shadow-2xl">
        {!capturedImage ? (
          <>
            {isInitializing && (
              <div className="absolute inset-0 flex flex-col items-center justify-center bg-slate-950/60 backdrop-blur-sm z-20">
                <motion.div
                  animate={{ rotate: 360 }}
                  transition={{ repeat: Infinity, duration: 1.5, ease: 'linear' }}
                  className="w-10 h-10 border-2 border-indigo-500/20 border-t-indigo-500 rounded-full mb-3"
                />
                <span className="text-xs text-indigo-400 font-mono">Initializing camera...</span>
              </div>
            )}

            {error ? (
              <div className="p-6 text-center max-w-sm flex flex-col items-center gap-3">
                <AlertCircle className="w-12 h-12 text-rose-500" />
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
                    {/* Scanner line animation */}
                    <div className="absolute left-0 right-0 h-0.5 bg-gradient-to-r from-transparent via-indigo-400 to-transparent top-1/2 -translate-y-1/2 animate-pulse" />
                    {/* Corner accents */}
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
          renderCapturedImage()
        )}
      </div>

      {/* Footer controls */}
      <div className="w-full max-w-lg z-10 flex flex-col gap-4">
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
          <div className="flex gap-3 px-2">
            <button
              onClick={() => {
                setCapturedImage(null);
                setCameraTrigger((prev) => prev + 1);
              }}
              className="flex-1 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold py-3.5 rounded-2xl transition flex items-center justify-center gap-2"
            >
              <RefreshCw className="w-4 h-4" />
              Retake Photo
            </button>
            <button
              onClick={handleConfirm}
              className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white font-bold py-3.5 rounded-2xl transition shadow-lg shadow-indigo-600/20 flex items-center justify-center gap-2 group"
            >
              <Check className="w-5 h-5 group-hover:scale-125 transition-transform" />
              Analyze Photo
            </button>
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
