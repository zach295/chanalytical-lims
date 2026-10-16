'use strict';

// Reporting convention: As(III) below its reporting limit contributes zero
// to calculated As(V), while the original As(III) result remains unchanged.
// Keep the existing whole-ug/L rounding and nonnegative floor for other inputs.
function calculateArsenicV(totalRaw, arsenicIIIRaw, asIIIReportingLimit = 1) {
  const totalText = String(totalRaw ?? '').trim();
  const iiiText = String(arsenicIIIRaw ?? '').trim();
  if (!totalText || !iiiText) return null;
  const total = Number(totalText);
  if (!Number.isFinite(total)) return null;

  let adjustedIII;
  if (iiiText.startsWith('<')) {
    const limitText = iiiText.slice(1).trim();
    const indicatedLimit = Number(limitText);
    if (!limitText || !Number.isFinite(indicatedLimit) ||
        indicatedLimit > asIIIReportingLimit) return null;
    adjustedIII = 0;
  } else {
    const measuredIII = Number(iiiText);
    if (!Number.isFinite(measuredIII)) return null;
    adjustedIII = measuredIII < asIIIReportingLimit ? 0 : measuredIII;
  }

  return Math.max(0, Math.round(total) - Math.round(adjustedIII));
}

module.exports = { calculateArsenicV };
