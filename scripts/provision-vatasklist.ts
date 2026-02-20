/**
 * Provision VATaskList using @pnp/sp and @pnp/nodejs
 *
 * Requirements:
 *  - Node.js
 *  - Install dependencies (locally):
 *      npm install @pnp/sp @pnp/nodejs cross-fetch
 *  - Provide these environment variables:
 *      SITE_URL       - SharePoint site URL (e.g. https://contoso.sharepoint.com/sites/mysite)
 *      CLIENT_ID      - Azure AD App Client ID (app must have Sites.FullControl.All or similar)
 *      CLIENT_SECRET  - Azure AD App Client Secret
 *
 * Run with ts-node or compile to JS first. Example using ts-node:
 *  npx ts-node scripts/provision-vatasklist.ts
 */

import { spfi } from '@pnp/sp';
import { SPFI } from '@pnp/sp';
import { SPFetchClient } from '@pnp/nodejs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';

const siteUrl = process.env.SITE_URL || process.env.TENANT_URL;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

if (!siteUrl || !clientId || !clientSecret) {
  console.error('Missing environment variables. Set SITE_URL, CLIENT_ID, CLIENT_SECRET');
  process.exit(1);
}

function getSP(): SPFI {
  // Initialize spfi with node fetch client
  return spfi(siteUrl).using(new SPFetchClient(clientId, clientSecret));
}

async function ensureVATaskList() {
  const sp = getSP();
  const title = 'VATaskList';

  try {
    // Check if list exists
    await sp.web.lists.getByTitle(title)();
    console.log(`${title} already exists.`);
    return;
  } catch (err) {
    console.log(`${title} not found — creating...`);
  }

  // Create a custom list (Template 100)
  const createRes = await sp.web.lists.add(title, `Provisioned by script: ${new Date().toISOString()}`, 100);
  console.log('List created:', createRes.data?.Id || createRes);

  const list = sp.web.lists.getByTitle(title);

  // Add common fields used by the boilerplate
  try {
    await list.fields.addText('TaskTitle');
    await list.fields.addText('AssignedTo');
    await list.fields.addBoolean('Completed');
    await list.fields.addDateTime('DueDate');
    await list.fields.addNote('Details');
    console.log('Fields added to', title);
  } catch (err) {
    console.warn('Error adding fields (some may already exist):', err.message || err);
  }
}

ensureVATaskList()
  .then(() => console.log('Provisioning complete.'))
  .catch((err) => {
    console.error('Provisioning failed:', err);
    process.exit(1);
  });
