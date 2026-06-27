// Abstract interface for Barcode Scanning
export interface BarcodeScannerProvider {
  name: string;
  isReady: () => boolean;
  init: () => Promise<void>;
  startScanning: (
    elementId: string,
    onScanSuccess: (text: string) => void,
    onScanError: (err: any) => void
  ) => Promise<void>;
  stopScanning: () => Promise<void>;
}

// Html5Qrcode implementation of BarcodeScannerProvider
export class Html5QrcodeScannerProvider implements BarcodeScannerProvider {
  name = 'Html5Qrcode';
  private html5QrcodeInstance: any = null;

  isReady() {
    return typeof window !== 'undefined';
  }

  async init() {
    if (typeof window === 'undefined') return;
    // html5-qrcode is dynamically imported to prevent SSR issues
    const { Html5Qrcode } = await import('html5-qrcode');
    console.log('Barcode scanner SDK initialized successfully:', this.name);
  }

  async startScanning(
    elementId: string,
    onScanSuccess: (text: string) => void,
    onScanError: (err: any) => void
  ) {
    if (typeof window === 'undefined') return;
    const { Html5Qrcode } = await import('html5-qrcode');

    // Make sure we stop any active scanning first
    await this.stopScanning();

    const html5Qrcode = new Html5Qrcode(elementId);
    this.html5QrcodeInstance = html5Qrcode;

    const config = {
      fps: 15,
      qrbox: (width: number, height: number) => {
        // Dynamic scan box for baggage barcodes
        return {
          width: Math.min(width * 0.85, 300),
          height: Math.min(height * 0.4, 120), // Wide rectangle for 1D baggage barcodes!
        };
      },
      aspectRatio: 1.777778, // 16:9 widescreen
    };

    // Use environment-preferred camera (back/environment camera)
    await html5Qrcode.start(
      { facingMode: 'environment' },
      config,
      (decodedText) => {
        // Trigger visual and audible feedback (beeps / vibrations)
        try {
          if (typeof window !== 'undefined' && window.navigator && window.navigator.vibrate) {
            window.navigator.vibrate(100);
          }
          // Simple synth audio beep for confirmation
          const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
          const oscillator = audioCtx.createOscillator();
          const gainNode = audioCtx.createGain();
          oscillator.connect(gainNode);
          gainNode.connect(audioCtx.destination);
          oscillator.type = 'sine';
          oscillator.frequency.setValueAtTime(1200, audioCtx.currentTime); // high pitched clean beep
          gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
          oscillator.start();
          oscillator.stop(audioCtx.currentTime + 0.08);
        } catch (e) {
          console.log('Audio/haptic feedback not supported or blocked by permissions');
        }

        onScanSuccess(decodedText);
      },
      (errorMessage) => {
        // Verbose error logging can be noisy, so we delegate to error callback
        onScanError(errorMessage);
      }
    );
  }

  async stopScanning() {
    if (this.html5QrcodeInstance) {
      try {
        if (this.html5QrcodeInstance.isScanning) {
          await this.html5QrcodeInstance.stop();
        }
      } catch (e) {
        console.error('Error stopping scanner:', e);
      } finally {
        this.html5QrcodeInstance = null;
      }
    }
  }
}

// Global active scanner instance that can be swapped out
let activeScanner: BarcodeScannerProvider = new Html5QrcodeScannerProvider();

export function getActiveScanner(): BarcodeScannerProvider {
  return activeScanner;
}

export function setActiveScanner(provider: BarcodeScannerProvider) {
  activeScanner = provider;
}
