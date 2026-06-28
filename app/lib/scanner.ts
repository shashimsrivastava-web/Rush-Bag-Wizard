import { Html5Qrcode, Html5QrcodeSupportedFormats } from 'html5-qrcode';

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
}

// Html5Qrcode implementation optimized for IATA Baggage Tags
export class Html5BaggageScanner implements IBaggageScanner {
  name = 'IATA-Optimized Html5Qrcode';
  private html5QrcodeInstance: Html5Qrcode | null = null;
  private lastScanTime: number = 0;
  private SCAN_COOLDOWN = 5000;
  private isTransitioning: boolean = false;

  isReady() {
    return typeof window !== 'undefined';
  }

  async init() {
    if (typeof window === 'undefined') return;
    console.log('Airline-Grade scanner SDK initialized:', this.name);
  }

  async startScanning(
    elementId: string,
    onScanSuccess: (result: ScanResult) => void,
    onScanError: (err: any) => void
  ) {
    if (typeof window === 'undefined') return;
    
    if (this.isTransitioning) {
      console.warn('Scanner is already transitioning states. Ignoring start request.');
      return;
    }

    try {
      this.isTransitioning = true;
      // Ensure cleanup
      await this.stopScanningInternal();

      const html5Qrcode = new Html5Qrcode(elementId);
      this.html5QrcodeInstance = html5Qrcode;

      const startTime = Date.now();

      const config = {
        fps: 30,
        qrbox: (width: number, height: number) => {
          const boxWidth = Math.min(width * 0.9, 400);
          const boxHeight = Math.min(height * 0.3, 150);
          return { width: boxWidth, height: boxHeight };
        },
        aspectRatio: 1.777778,
        formatsToSupport: [
          Html5QrcodeSupportedFormats.ITF,
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.CODE_39
        ]
      };

      await html5Qrcode.start(
        { facingMode: 'environment' },
        config,
        (decodedText, decodedResult) => {
          const now = Date.now();
          
          const validatedTag = validateBaggageTag(decodedText);
          if (!validatedTag) {
            onScanError('Invalid IATA baggage tag detected. Please rescan.');
            return;
          }

          if (now - this.lastScanTime < this.SCAN_COOLDOWN) return;
          this.lastScanTime = now;

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
            barcodeFormat: decodedResult.result.format?.formatName || 'UNKNOWN',
            scannerUsed: this.name,
            confidenceScore: 0.95,
            scanTime: now - startTime
          });
        },
        (errorMessage) => {
          onScanError(errorMessage);
        }
      );
    } catch (err) {
      console.error('Failed to start scanning:', err);
      throw err;
    } finally {
      this.isTransitioning = false;
    }
  }

  private async stopScanningInternal() {
    if (this.html5QrcodeInstance) {
      try {
        if (this.html5QrcodeInstance.isScanning) {
          await this.html5QrcodeInstance.stop();
        }
      } catch (e) {
        // Only log if it's not the transition error we're trying to avoid
        if (e instanceof Error && !e.message.includes('transition')) {
          console.error('Error stopping scanner:', e);
        }
      } finally {
        this.html5QrcodeInstance = null;
      }
    }
  }

  async stopScanning() {
    if (this.isTransitioning) {
      console.warn('Scanner is already transitioning states. Ignoring stop request.');
      return;
    }
    this.isTransitioning = true;
    try {
      await this.stopScanningInternal();
    } finally {
      this.isTransitioning = false;
    }
  }
}

// Bridge Pattern Management
let activeScanner: IBaggageScanner = new Html5BaggageScanner();

export function getActiveScanner(): IBaggageScanner {
  return activeScanner;
}

export function setActiveScanner(provider: IBaggageScanner) {
  activeScanner = provider;
}
