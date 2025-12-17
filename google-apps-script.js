/**
 * AI Pipelines Dashboard - Google Apps Script
 *
 * SHEET STRUCTURE:
 * 1. Pipelines: id, name
 * 2. Clients: id, pipelineId, name
 * 3. ClientData: id, clientId, item
 * 4. Sections: id, clientId, name, sortOrder
 * 5. Requirements: id, sectionId, name, subname, priority, status, sortOrder
 * 6. Bullets: id, requirementId, text, sortOrder
 * 7. Technologies: id, requirementId, name, type, stage, progress, links
 */

function doGet(e) {
  var output;

  try {
    const action = e.parameter.action;

    if (action === 'getData') {
      output = getData();
    } else {
      output = ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' }));
    }
  } catch (error) {
    output = ContentService.createTextOutput(JSON.stringify({ error: error.toString() }));
  }

  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function doPost(e) {
  var output;

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'updateRequirementStatus') {
      output = updateRequirementStatus(data.id, data.status);
    } else if (action === 'updateTechnologyProgress') {
      output = updateTechnologyProgress(data.id, data.progress);
    } else {
      output = ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' }));
    }
  } catch (error) {
    output = ContentService.createTextOutput(JSON.stringify({ error: error.toString() }));
  }

  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pipelines = sheetToObjects(ss.getSheetByName('Pipelines'));
  const clients = sheetToObjects(ss.getSheetByName('Clients'));
  const clientData = sheetToObjects(ss.getSheetByName('ClientData'));
  const sections = sheetToObjects(ss.getSheetByName('Sections'));
  const requirements = sheetToObjects(ss.getSheetByName('Requirements'));
  const bullets = sheetToObjects(ss.getSheetByName('Bullets'));
  const technologies = sheetToObjects(ss.getSheetByName('Technologies'));

  // Sort helper
  const bySort = (a, b) => (a.sortOrder || 0) - (b.sortOrder || 0);

  // Nest bullets under requirements
  const requirementsWithBullets = requirements.map(req => ({
    ...req,
    bullets: bullets.filter(b => b.requirementId === req.id).sort(bySort),
    technologies: technologies.filter(t => t.requirementId === req.id)
  })).sort(bySort);

  // Nest requirements under sections
  const sectionsWithReqs = sections.map(sec => ({
    ...sec,
    requirements: requirementsWithBullets.filter(r => r.sectionId === sec.id)
  })).sort(bySort);

  // Nest sections and clientData under clients
  const clientsWithData = clients.map(client => ({
    ...client,
    data: clientData.filter(d => d.clientId === client.id),
    sections: sectionsWithReqs.filter(s => s.clientId === client.id)
  }));

  // Nest clients under pipelines
  const result = pipelines.map(pipeline => ({
    ...pipeline,
    clients: clientsWithData.filter(c => c.pipelineId === pipeline.id)
  }));

  return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }));
}

function updateRequirementStatus(id, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Requirements');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('status') + 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, statusCol).setValue(status);
      break;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true }));
}

function updateTechnologyProgress(id, progress) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Technologies');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const progressCol = headers.indexOf('progress') + 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, progressCol).setValue(progress);
      break;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true }));
}

function sheetToObjects(sheet) {
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const objects = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    let hasData = false;

    for (let j = 0; j < headers.length; j++) {
      let value = data[i][j];

      // Skip empty rows
      if (j === 0 && !value) continue;

      // Parse links column as JSON
      if (headers[j] === 'links' && typeof value === 'string' && value) {
        try {
          value = JSON.parse(value);
        } catch (e) {
          value = [];
        }
      }

      // Parse numeric columns
      if (['progress', 'sortOrder'].includes(headers[j])) {
        value = Number(value) || 0;
      }

      obj[headers[j]] = value;
      if (value !== '') hasData = true;
    }

    if (hasData && obj.id) {
      objects.push(obj);
    }
  }

  return objects;
}
