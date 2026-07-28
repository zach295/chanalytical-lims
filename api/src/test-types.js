const { app } = require('@azure/functions');
const { listItems, createItem, updateItem, LISTS } = require('../shared/graph');

// SharePoint internal field mapping (Test Types):
// Title=name, field_1=category, field_2=price, field_3=suffix, field_4=includes, field_5=active

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
        name:     r.Title    || '',
        category: r.field_1  || 'Package',
        price:    r.field_2  !== undefined && r.field_2 !== null ? r.field_2 : '',
        suffix:   r.field_3  || '',
        includes: r.field_4  || '',
        active:   r.field_5  !== false,
      }));

      const elements = elItems.map(r => ({
        _row:   r._id,
        name:   r.Title   || '',
        abbrev: r.field_1 || '',
        price:  r.field_2 !== undefined && r.field_2 !== null ? r.field_2 : '',
        active: r.field_3 !== false,
      }));

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ testTypes, elements }),
      };
    } catch(e) {
      context.log('[test-types-read] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
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
        if (!name) return { status: 400, body: JSON.stringify({ error: 'Name required' }) };
        const fields = {
          Title:   name,
          field_1: category || 'Package',
          field_2: parseFloat(price) || 0,
          field_3: suffix   || '',
          field_4: includes || '',
        };
        if (rowNum) await updateItem(LISTS.TEST_TYPES, rowNum, fields);
        else        await createItem(LISTS.TEST_TYPES, { ...fields, field_5: true });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'deleteTestType') {
        const { active } = body;
        const isActive = active === 'TRUE' || active === true;
        await updateItem(LISTS.TEST_TYPES, rowNum, { field_5: isActive });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'saveElement') {
        const { name, abbrev, price } = body;
        if (!name) return { status: 400, body: JSON.stringify({ error: 'Name required' }) };
        const fields = {
          Title:   name,
          field_1: abbrev || '',
          field_2: parseFloat(price) || 0,
        };
        if (rowNum) await updateItem(LISTS.ELEMENTS, rowNum, fields);
        else        await createItem(LISTS.ELEMENTS, { ...fields, field_3: true });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'deleteElement') {
        const { active } = body;
        const isActive = active === 'TRUE' || active === true;
        await updateItem(LISTS.ELEMENTS, rowNum, { field_3: isActive });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action' }) };
    } catch(e) {
      context.log('[test-types-write] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
