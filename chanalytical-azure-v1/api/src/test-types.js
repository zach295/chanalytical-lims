const { app } = require('@azure/functions');
const { listItems, createItem, updateItem, LISTS } = require('../shared/graph');

// ── READ ──────────────────────────────────────────────────────────────────────
app.http('test-types-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const [ttItems, elItems] = await Promise.all([
        listItems(LISTS.TEST_TYPES),
        listItems(LISTS.ELEMENTS),
      ]);

      const testTypes = ttItems.map(r => ({
        _row:     r._id,
        name:     r.Name     || r.Title || '',
        category: r.Category || 'Package',
        price:    r.Price    || '',
        suffix:   r.Suffix   || '',
        includes: r.Includes || '',
        active:   r.Active !== 'FALSE',
      }));

      const elements = elItems.map(r => ({
        _row:   r._id,
        name:   r.Name   || r.Title || '',
        abbrev: r.Abbrev || '',
        price:  r.Price  || '',
        active: r.Active !== 'FALSE',
      }));

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ testTypes, elements }),
      };
    } catch(e) {
      context.log('[test-types-read] Error:', e.message);
      return { status:500, body: JSON.stringify({ error: e.message }) };
    }
  }
});

// ── WRITE ─────────────────────────────────────────────────────────────────────
app.http('test-types-write', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const { action, rowNum } = body;

      if (action === 'saveTestType') {
        const { name, category, price, suffix, includes } = body;
        if (!name) return { status:400, body: JSON.stringify({ error:'Name required' }) };
        const fields = { Title:name, Name:name, Category:category||'Package', Price:price||'', Suffix:suffix||'', Includes:includes||'' };
        if (rowNum) await updateItem(LISTS.TEST_TYPES, rowNum, fields);
        else        await createItem(LISTS.TEST_TYPES, { ...fields, Active:'TRUE' });
        return { status:200, headers:{'content-type':'application/json'}, body: JSON.stringify({ success:true }) };
      }

      if (action === 'deleteTestType') {
        const { active } = body;
        await updateItem(LISTS.TEST_TYPES, rowNum, { Active: active || 'FALSE' });
        return { status:200, headers:{'content-type':'application/json'}, body: JSON.stringify({ success:true }) };
      }

      if (action === 'saveElement') {
        const { name, abbrev, price } = body;
        if (!name) return { status:400, body: JSON.stringify({ error:'Name required' }) };
        const fields = { Title:name, Name:name, Abbrev:abbrev||'', Price:price||'' };
        if (rowNum) await updateItem(LISTS.ELEMENTS, rowNum, fields);
        else        await createItem(LISTS.ELEMENTS, { ...fields, Active:'TRUE' });
        return { status:200, headers:{'content-type':'application/json'}, body: JSON.stringify({ success:true }) };
      }

      if (action === 'deleteElement') {
        const { active } = body;
        await updateItem(LISTS.ELEMENTS, rowNum, { Active: active || 'FALSE' });
        return { status:200, headers:{'content-type':'application/json'}, body: JSON.stringify({ success:true }) };
      }

      return { status:400, body: JSON.stringify({ error:'Unknown action' }) };
    } catch(e) {
      context.log('[test-types-write] Error:', e.message);
      return { status:500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
