const { Worker } = require('worker_threads');
const path = require('path');

const processLargeExport = (filters, callback) => {
    return new Promise((resolve, reject) => {
        const worker = new Worker(path.resolve(__dirname, './exportWorker.js'), {
            workerData: { filters }
        });

        worker.on('message', (message) => {
            if (message.progress) {
                callback(message.progress); // Update progress
            } else if (message.result) {
                resolve(message.result);
            }
        });

        worker.on('error', reject);
        worker.on('exit', (code) => {
            if (code !== 0) reject(new Error(`Worker stopped with exit code ${code}`));
        });
    });
};

module.exports = { processLargeExport };