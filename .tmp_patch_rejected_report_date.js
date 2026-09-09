const fs = require('fs');
const path = 'api/src/approve-scan.js';
let s = fs.readFileSync(path, 'utf8');

const old1 = `    fields["Client Code"]        = params.clientCode || '';
    // Report Date left blank — filled in when report is actually sent
    fields["Location"]           = params.location || '';`;
const new1 = `    fields["Client Code"]        = params.clientCode || '';
    // Normal samples get Report Date when the report is actually sent.
    // Rejected samples are never reported, so preserve the legacy billing behavior:
    // assign the next business day after Date Rec'd at approval time.
    if (params.isRejected) {
      const next = nextBusinessDay(params.receivedDate);
      if (next) {
        const [m, d, y] = next.split('/');
        fields["Report Date"] = m + '/' + d + '/' + String(y).slice(-2);
      } else {
        fields["Report Date"] = '';
      }
    } else {
      fields["Report Date"] = '';
    }
    fields["Location"]           = params.location || '';`;

if (!s.includes(old1)) throw new Error('RTB report-date anchor not found');
s = s.replace(old1, new1);

const old2 = `            testName:        rowTest,
          }, context, rtbCache).catch(e => ({ success:false, error:e.message }));`;
const new2 = `            testName:        rowTest,
            isRejected:      !!(item.isRejection || item.isRejected),
          }, context, rtbCache).catch(e => ({ success:false, error:e.message }));`;

if (!s.includes(old2)) throw new Error('RTB call anchor not found');
s = s.replace(old2, new2);

fs.writeFileSync(path, s);
console.log('Patched rejected-sample RTB report dates');
