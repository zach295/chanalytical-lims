/**
 * fix-intake-field.js — one-shot fix utility
 * POST { itemId, field, value } → patches that field on an Archived Intake item
 * DELETE this file after use
 */
const { app } = require('@azure/functions');
const { getToken, updateItem, LISTS } = require('../shared/graph');

app.http('fix-intake-field', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const { itemId, field, value } = await request.json();
    if (!itemId || !field || value === undefined)
      return { status: 400, jsonBody: { error: 'itemId, field, value required' } };

    const token = await getToken();
    const patch = { [field]: value };
    await updateItem(LISTS.ARCHIVED_INTAKE, itemId, patch);
    context.log(`[fix-intake] Patched item ${itemId}: ${field} = "${value}"`);
    return { status: 200, jsonBody: { success: true, itemId, field, value } };
  }
});
