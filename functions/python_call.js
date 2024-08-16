"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.processFile = processFile;
exports.salary = salary;
const child_process_1 = require("child_process");
const path_1 = __importDefault(require("path"));
// Define the path to the Python script
const pythonScriptPath = path_1.default.join(__dirname, 'file_process.py');
const salary_path = path_1.default.join(__dirname, 'salary_recalculation.py');
function processFile(inputData) {
    return __awaiter(this, void 0, void 0, function* () {
        return new Promise((resolve, reject) => {
            // Spawn a new child process to run the Python script
            const pythonProcess = (0, child_process_1.spawn)('./venv/bin/python3', [pythonScriptPath]);
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
                }
                else {
                    resolve();
                }
            });
            // Write the input data to the Python script's stdin
            pythonProcess.stdin.write(inputData);
            pythonProcess.stdin.end();
        });
    });
}
function salary() {
    return __awaiter(this, void 0, void 0, function* () {
        return new Promise((resolve, reject) => {
            // Spawn a new child process to run the Python script
            const Process = (0, child_process_1.spawn)('./venv/bin/python3', [salary_path]);
            // Handle stdout data from the Python script
            Process.stdout.on('data', (data) => {
                console.log(`Python script output: ${data}`);
            });
            // Handle stderr data from the Python script
            Process.stderr.on('data', (data) => {
                console.error(`Python script error: ${data}`);
            });
            // Handle the exit event of the Python script
            Process.on('exit', (code) => {
                if (code !== 0) {
                    reject(new Error(`Python script exited with code ${code}`));
                }
                else {
                    resolve();
                }
            });
            Process.stdin.end();
        });
    });
}
