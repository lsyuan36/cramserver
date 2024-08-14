import { spawn } from 'child_process';
import path from 'path';

// Define the path to the Python script
const pythonScriptPath = path.join(__dirname, 'file_process.py');

export async function processFile(inputData: string): Promise<void> {
    return new Promise((resolve, reject) => {
        // Spawn a new child process to run the Python script
        const pythonProcess = spawn('./venv/bin/python3', [pythonScriptPath]);

        // Handle stdout data from the Python script
        pythonProcess.stdout.on('data', (data) => {
            console.log(`Python script output: ${data}`);
        });

        // Handle stderr data from the Python script
        pythonProcess.stderr.on('data', (data) => {
            console.error(`Python script error: ${data}`);
        });

        // Handle the exit event of the Python script
        pythonProcess.on('exit', (code) => {
            if (code !== 0) {
                reject(new Error(`Python script exited with code ${code}`));
            } else {
                resolve();
            }
        });

        // Write the input data to the Python script's stdin
        pythonProcess.stdin.write(inputData);
        pythonProcess.stdin.end();
    });
}
