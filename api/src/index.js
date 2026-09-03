// Register all Azure Functions.
// Load each module independently so one bad entry point cannot prevent
// every HTTP function (including /api/health) from registering.

const modules = [
  './health',
  './clients-read',
  './get-scan-queue',
  './accession-status',
  './clients-write',
  './users-manage',
  './reject-sample',
  './get-rejections',
  './cache-results',
  './approve-scan',
  './mark-scan-processed',
  './test-types',
  './inventory-read',
  './inventory-write',
  './setup-lists',
  './scan-folder',
  './generate-report',
  './prepare-report',
  './export-pdf',
  './patch-report-cell',
  './control-sheet',
  './update-sample',
  './send-report',
  './render-report-pdf',
  './update-inventory',
  './sync-to-sheets',
  './auth-role',
  './approve-scan-debug',
  './ms-token-exchange',
  './labid-search',
  './list-columns',
  './import-icpms',
  './import-control',
  './import-radon',
  './import-acid',
  // './import-bacteria',
  // './import-ph',
  './sync-results-cache',
  './billing-read',
  './billing-update',
];

for (const modulePath of modules) {
  try {
    require(modulePath);
  } catch (err) {
    console.error(`[function-entry] Failed loading ${modulePath}:`, err?.stack || err);
  }
}
