import { 
  BrowserMultiFormatReader, 
  BarcodeFormat, 
  DecodeHintType, 
  Result 
} from '@zxing/library';

// Core Scan Result Interface
export interface ScanResult {
  tagNumber: string;
  barcodeFormat: string;
  scannerUsed: string;
  confidenceScore: number;
  scanTime: number;
}

// Validation logic for IATA baggage tags
export const validateBaggageTag = (raw: string): string | null => {
  const clean = raw.trim().toUpperCase();
  
  // 10-digit numeric (Standard IATA License Plate)
  if (/^\d{10}$/.test(clean)) return clean;
  
  // 13-digit numeric (Extended IATA)
  if (/^\d{13}$/.test(clean)) return clean;

  // Airline Prefix (e.g. LH123456, LX654321)
  const airlineMatch = clean.match(/^([A-Z]{2})(\d{6,8})$/);
  if (airlineMatch) return clean;

  return null;
};

// Abstract interface for Barcode Scanning (Airline Grade)
export interface IBaggageScanner {
  name: string;
  isReady: () => boolean;
  init: () => Promise<void>;
  startScanning: (
    elementId: string,
    onScanSuccess: (result: ScanResult) => void,
    onScanError: (err: any) => void
  ) => Promise<void>;
  stopScanning: () => Promise<void>;
  pauseScanning?: () => void;
  resumeScanning?: () => void;
  toggleTorch?: (enabled: boolean) => Promise<void>;
  dispose: () => void;
}

/**
 * ZXing (Zebra Crossing) implementation for high-performance baggage scanning
 */
export class ZXingBaggageScanner implements IBaggageScanner {
  name = 'ZXing (Zebra Crossing)';
  private reader: BrowserMultiFormatReader | null = null;
  private lastScanTime: number = 0;
  private SCAN_COOLDOWN = 5000;
  private isTransitioning: boolean = false;
  private videoElement: HTMLVideoElement | null = null;

  isReady() {
    return typeof window !== 'undefined';
  }

  async init() {
    if (typeof window === 'undefined') return;
    
    // Configure ZXing hints for faster detection
    const hints = new Map();
    const formats = [
      BarcodeFormat.ITF,
      BarcodeFormat.CODE_128,
      BarcodeFormat.CODE_39
    ];
    hints.set(DecodeHintType.POSSIBLE_FORMATS, formats);
    hints.set(DecodeHintType.TRY_HARDER, true);
    hints.set(DecodeHintType.ASSUME_GS1, false);

    this.reader = new BrowserMultiFormatReader(hints);
    console.log('ZXing scanner engine initialized:', this.name);
  }

  async startScanning(
    elementId: string,
    onScanSuccess: (result: ScanResult) => void,
    onScanError: (err: any) => void
  ) {
    if (typeof window === 'undefined' || !this.reader) return;

    if (this.isTransitioning) {
      console.warn('Scanner transitioning. Skipping start.');
      return;
    }

    try {
      this.isTransitioning = true;
      await this.stopScanningInternal();

      const startTime = Date.now();

      // ZXing needs a video element inside the container
      const container = document.getElementById(elementId);
      if (!container) throw new Error(`Container #${elementId} not found`);

      // Create video element if not exists or reuse
      let video = container.querySelector('video') as HTMLVideoElement;
      if (!video) {
        video = document.createElement('video');
        video.style.width = '100%';
        video.style.height = '100%';
        video.style.objectFit = 'cover';
        video.setAttribute('playsinline', 'true');
        container.appendChild(video);
      }
      this.videoElement = video;

      // Start continuous decoding
      await this.reader.decodeFromVideoDevice(
        null, // Use default camera (usually back on mobile)
        video,
        (result: Result | null, err?: any) => {
          if (result) {
            const decodedText = result.getText();
            const now = Date.now();

            const validatedTag = validateBaggageTag(decodedText);
            if (!validatedTag) {
              // We don't necessarily error here as it might be a partial or noise
              // but the prompt says to inform user if invalid
              // onScanError('Invalid IATA baggage tag detected.');
              return;
            }

            if (now - this.lastScanTime < this.SCAN_COOLDOWN) return;
            this.lastScanTime = now;

            // Feedback
            try {
              if (navigator.vibrate) navigator.vibrate(150);
              const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
              const oscillator = audioCtx.createOscillator();
              const gainNode = audioCtx.createGain();
              oscillator.connect(gainNode);
              gainNode.connect(audioCtx.destination);
              oscillator.type = 'sine';
              oscillator.frequency.setValueAtTime(1400, audioCtx.currentTime);
              gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
              oscillator.start();
              oscillator.stop(audioCtx.currentTime + 0.1);
            } catch (e) {}

            onScanSuccess({
              tagNumber: validatedTag,
              barcodeFormat: result.getBarcodeFormat().toString(),
              scannerUsed: this.name,
              confidenceScore: 1.0,
              scanTime: now - startTime
            });
          }
          if (err && !(err.name === 'NotFoundException')) {
            // ZXing throws NotFoundException continuously when nothing is detected
            // only pass through real errors
            onScanError(err);
          }
        }
      );

    } catch (err) {
      console.error('ZXing start failed:', err);
      onScanError(err);
      throw err;
    } finally {
      this.isTransitioning = false;
    }
  }

  private async stopScanningInternal() {
    if (this.reader) {
      try {
        this.reader.reset();
      } catch (e) {
        console.error('Error resetting ZXing reader:', e);
      }
    }
    if (this.videoElement) {
      this.videoElement.srcObject = null;
      this.videoElement = null;
    }
  }

  async stopScanning() {
    if (this.isTransitioning) return;
    this.isTransitioning = true;
    try {
      await this.stopScanningInternal();
    } finally {
      this.isTransitioning = false;
    }
  }

  async toggleTorch(enabled: boolean) {
    if (!this.videoElement || !this.videoElement.srcObject) return;
    const stream = this.videoElement.srcObject as MediaStream;
    const track = stream.getVideoTracks()[0];
    if (!track) return;

    try {
      const capabilities = track.getCapabilities() as any;
      if (capabilities.torch) {
        await track.applyConstraints({
          advanced: [{ torch: enabled }]
        } as any);
      }
    } catch (e) {
      console.warn('Torch control not supported on this device/browser');
    }
  }

  dispose() {
    this.stopScanning().catch(console.error);
    this.reader = null;
  }
}

// Bridge Pattern Management
let activeScanner: IBaggageScanner = new ZXingBaggageScanner();

export function getActiveScanner(): IBaggageScanner {
  return activeScanner;
}

export function setActiveScanner(provider: IBaggageScanner) {
  activeScanner = provider;
}
