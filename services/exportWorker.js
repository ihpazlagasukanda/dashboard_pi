const { parentPort, workerData } = require('worker_threads');
const { generatePetaniReport } = require('./excelService');

(async () => {
    try {
        const { filters } = workerData;
        const result = await generatePetaniReport(filters);
        parentPort.postMessage({ result });
    } catch (error) {
        throw error;
    }
})();