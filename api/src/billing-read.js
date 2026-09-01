/**
 * billing-read.js — reads Reports to be Billed SharePoint list
 * Returns all items for billing tab and revenue graphs
 */
const { app }       = require('@azure/functions');
const { getToken, listItems } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

app.http('billing-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const token  = await getToken();
      const siteId = process.env.SP_SITE_ID;

      // Fetch all items from Reports to be Billed list
      const res = await fetch(
        `${GRAPH}/sites/${siteId}/lists?$select=id,displayName`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const allLists = (await res.json()).value || [];
      const listId   = allLists.find(l => l.displayName === 'Reports to be Billed')?.id;
      if (!listId) return { status: 404, jsonBody: { error: 'Reports to be Billed list not found' } };

      // Fetch all items with fields expanded
      let items = [], nextLink = `${GRAPH}/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=999`;
      while (nextLink) {
        const r    = await fetch(nextLink, { headers: { Authorization: `Bearer ${token}` } });
        const data = await r.json();
        items.push(...(data.value || []).map(i => i.fields || {}));
        nextLink = data['@odata.nextLink'] || null;
      }

      // Map fields to clean objects
      const rows = items.map(f => ({
        id:            f.id,
        labId:         f.Title        || '',
        dateRec:       f.Date_x0020_Rec_x0027_d || f.field_1 || '',
        timeRec:       f.Time_x0020_Rec_x0027_d || f.field_2 || '',
        dateDrawn:     f.Date_x0020_Drawn  || '',
        timeDrawn:     f.Time_x0020_Drawn  || '',
        customer:      f.Customer     || '',
        clientCode:    f.Client_x0020_Code || '',
        reportDate:    f.Report_x0020_Date || '',
        location:      f.Location     || '',
        city:          f.City_x002F_Town   || '',
        state:         f.State        || '',
        zip:           f.Zip          || '',
        service:       f.Item_x002F_Service || '',
        testTypeSKU:   f.Test_x0020_Type_x0020_SKU || '',
        rwResults:     f.RWResults  || '',
        qty:           parseFloat(f.Qty || 1),
        rate:          parseFloat(f.Rate || 0),
        amt:           parseFloat(f.Amt  || 0),
        qb:            !!f.QB,
        disc:          parseFloat(f.Disc || 0),
        stmtDate:      f.Statement_x002F_Inv_x0020_Date || '',
        paid:          !!f.Pd,
        amtPaid:       parseFloat(f.Amt_x0020_Pd || 0),
        datePaid:      f.Date_x0020_Pd || '',
      }));

      context.log(`[billing-read] ${rows.length} items`);
      return { status: 200, jsonBody: { success: true, rows } };

    } catch (err) {
      context.log('[billing-read] error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
