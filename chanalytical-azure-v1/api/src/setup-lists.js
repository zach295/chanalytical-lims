/**
 * setup-lists — Run ONCE to create all needed SharePoint Lists.
 * GET /api/setup-lists → creates all lists with correct columns.
 * Safe to run multiple times — skips lists that already exist.
 */
const { app } = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH  = 'https://graph.microsoft.com/v1.0';
const SITE_ID = process.env.SP_SITE_ID;

async function graphPost(path, body, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  if (!res.ok && res.status !== 409) throw new Error(`${path} → ${res.status}: ${text.slice(0,200)}`);
  try { return JSON.parse(text); } catch { return {}; }
}

async function ensureColumn(listId, col, token) {
  await fetch(`${GRAPH}/sites/${SITE_ID}/lists/${listId}/columns`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(col),
  }).catch(() => {}); // Ignore if column already exists
}

const TEXT   = n => ({ name: n, text: {} });
const NUM    = n => ({ name: n, number: {} });
const BOOL   = n => ({ name: n, boolean: {} });
const CHOICE = (n, choices) => ({ name: n, choice: { choices } });

const LISTS_SCHEMA = [
  {
    name: 'Archived Intake',
    columns: [
      TEXT('Timestamp'), TEXT('FullId'), TEXT('CoaTest'), TEXT('Customer'),
      TEXT('DateDrawn'), TEXT('TimeDrawn'), TEXT('ReceivedDate'), TEXT('ReceivedTime'),
      TEXT('Location'), TEXT('City'), TEXT('State'), TEXT('Zip'),
      TEXT('ReviewedBy'), { name:'Notes', text:{ allowMultipleLines:true } },
      CHOICE('ReportStatus', ['Pending','Sent']),
    ],
  },
  {
    name: 'Review Queue',
    columns: [
      TEXT('FileId'), TEXT('BarcodeId'), TEXT('ClientName'), TEXT('Email'),
      TEXT('Address'), TEXT('City'), TEXT('State'), TEXT('Zip'),
      TEXT('SampleDate'), TEXT('SampleTime'), TEXT('ReceivedDate'), TEXT('ReceivedTime'),
      { name:'TestSelections', text:{ allowMultipleLines:true } },
      TEXT('ValidationErrors'), NUM('OcrConfidence'),
      TEXT('ProcessedDate'), TEXT('ScannedBy'), TEXT('ApprovedBy'),
      CHOICE('ReviewStatus', ['Pending','Approved','Discarded','Processed']),
      CHOICE('OcrStatus',    ['OK','Error']),
    ],
  },
  {
    name: 'Clients',
    columns: [
      TEXT('ClientName'), TEXT('ClientCode'), TEXT('Abbrev'), TEXT('Email'),
      { name:'Aliases', text:{ allowMultipleLines:true } },
      TEXT('Phone'), CHOICE('Active',['TRUE','FALSE']), TEXT('Notes'),
    ],
  },
  {
    name: 'Users',
    columns: [
      TEXT('Email'), TEXT('Name'),
      CHOICE('Role', ['admin','office','lab','wq','public','deactivated']),
      TEXT('ClientKey'), TEXT('RegCode'), TEXT('CreatedBy'), TEXT('CreatedAt'),
      CHOICE('MustReset',['true','false']), CHOICE('Active',['TRUE','FALSE']),
    ],
  },
  {
    name: 'Rejected',
    columns: [
      TEXT('LabId'),
      CHOICE('RejectionType', ['Rejected - Timeout','Rejected - Chlorine','Rejected - Other']),
      { name:'Reason', text:{ allowMultipleLines:true } },
      TEXT('RejectedBy'), TEXT('Timestamp'),
    ],
  },
  {
    name: 'Accession Log',
    columns: [ TEXT('BaseId'), TEXT('FullId'), TEXT('CoaTest'), TEXT('Suffix'), TEXT('Timestamp') ],
  },
  {
    name: 'Test Types',
    columns: [
      TEXT('Name'), CHOICE('Category',['Package','AIO','Special','Individual']),
      TEXT('Price'), TEXT('Suffix'), { name:'Includes', text:{ allowMultipleLines:true } },
      CHOICE('Active',['TRUE','FALSE']),
    ],
  },
  {
    name: 'Elements',
    columns: [ TEXT('Name'), TEXT('Abbrev'), TEXT('Price'), CHOICE('Active',['TRUE','FALSE']) ],
  },
  {
    name: 'Activity Log',
    columns: [
      TEXT('Date'), TEXT('Time'), TEXT('Client'),
      CHOICE('Type',['check_in','sent','received','adjust','initial','assemble']),
      NUM('Qty'), { name:'Notes', text:{ allowMultipleLines:true } }, TEXT('By'),
    ],
  },
  {
    name: 'Client Inventory',
    columns: [
      TEXT('ClientKey'), NUM('InStock'), NUM('Sampled'),
      NUM('TotalSent'), NUM('TotalReceived'), TEXT('LastActivity'),
    ],
  },
  {
    name: 'Results Cache',
    columns: [
      TEXT('LabId'), TEXT('Timestamp'),
      { name:'Data', text:{ allowMultipleLines:true } },
      CHOICE('Sent',['','Sent']),
    ],
  },
];

app.http('setup-lists', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const log = [];
    try {
      const token = await getToken();

      for (const listDef of LISTS_SCHEMA) {
        try {
          // Create list (409 = already exists, that's fine)
          const created = await graphPost(
            `/sites/${SITE_ID}/lists`,
            { displayName: listDef.name, list: { template: 'genericList' } },
            token
          );
          const listId = created.id;
          if (!listId) { log.push(`⚠️ ${listDef.name}: could not get list ID (may already exist)`); continue; }
          log.push(`✅ List created: ${listDef.name} (${listId})`);

          // Add columns
          for (const col of listDef.columns) {
            await ensureColumn(listId, col, token);
          }
          log.push(`   └ ${listDef.columns.length} columns added`);
        } catch(e) {
          log.push(`⚠️ ${listDef.name}: ${e.message}`);
        }
      }

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, log }),
      };
    } catch(e) {
      return { status: 500, body: JSON.stringify({ error: e.message, log }) };
    }
  }
});
