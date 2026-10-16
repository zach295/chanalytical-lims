const assert = require('node:assert/strict');
const { calculateArsenicV } = require('../shared/arsenic-speciation');

// Below the 1 ug/L As(III) reporting limit, use zero in As(V) calculation.
assert.equal(calculateArsenicV('2', '0.7'), 2);
assert.equal(calculateArsenicV('2', '0.99'), 2);
assert.equal(calculateArsenicV('2', '<1'), 2);
assert.equal(calculateArsenicV('2', '0'), 2);

// At or above reporting limit, retain existing whole-number calculation.
assert.equal(calculateArsenicV('2', '1'), 1);
assert.equal(calculateArsenicV('3', '1.4'), 2);
assert.equal(calculateArsenicV('2', '1.6'), 0);

// Missing or invalid data must not produce a fabricated result.
assert.equal(calculateArsenicV('2', ''), null);
assert.equal(calculateArsenicV('', '0.7'), null);
assert.equal(calculateArsenicV('2', '<2'), null);
assert.equal(calculateArsenicV('2', 'ND'), null);
console.log('Arsenic speciation calculation checks passed');
