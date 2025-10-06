import { spawn } from 'child_process';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export interface PowerShellResult {
  success: boolean;
  data?: any;
  error?: string;
}

export class PowerShellBridge {
  private scriptPath: string;

  constructor() {
    // Path to the PowerShell script
    this.scriptPath = path.join(__dirname, '../../scripts/outlook-calendar.ps1');
  }

  async executeScript(action: string, params: Record<string, any> = {}): Promise<PowerShellResult> {
    return new Promise((resolve, reject) => {
      // Build PowerShell command arguments
      const args = [
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', this.scriptPath,
        '-Action', action
      ];

      // Add parameters
      for (const [key, value] of Object.entries(params)) {
        if (value !== undefined && value !== null && value !== '') {
          args.push(`-${key}`);
          args.push(String(value));
        }
      }

      const powershell = spawn('powershell.exe', args, {
        stdio: ['ignore', 'pipe', 'pipe']
      });

      let stdout = '';
      let stderr = '';

      powershell.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      powershell.stderr.on('data', (data) => {
        stderr += data.toString();
      });

      powershell.on('close', (code) => {
        if (code !== 0) {
          resolve({
            success: false,
            error: stderr || 'PowerShell script failed'
          });
          return;
        }

        try {
          // Parse JSON output from PowerShell
          const result = JSON.parse(stdout.trim());

          // Check if result contains an error
          if (result.error) {
            resolve({
              success: false,
              error: result.error
            });
          } else {
            resolve({
              success: true,
              data: result
            });
          }
        } catch (parseError) {
          resolve({
            success: false,
            error: `Failed to parse PowerShell output: ${parseError}`
          });
        }
      });

      powershell.on('error', (error) => {
        resolve({
          success: false,
          error: `Failed to execute PowerShell: ${error.message}`
        });
      });
    });
  }
}
