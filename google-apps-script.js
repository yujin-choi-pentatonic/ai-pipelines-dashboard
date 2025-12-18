/**
 * AI Pipelines Dashboard - Google Apps Script
 *
 * SHEET STRUCTURE:
 * 1. Categories: id, name, sortOrder
 * 2. Pipelines: id, categoryId, name, sortOrder
 * 3. Clients: id, pipelineId, name
 * 4. ClientData: id, clientId, item
 * 5. Sections: id, clientId, name, sortOrder
 * 6. Requirements: id, sectionId, name, subname, priority, status, sortOrder
 * 7. Bullets: id, requirementId, text, sortOrder
 * 8. Technologies: id, requirementId, name, type, stage, progress, links
 * 9. Signoffs: id, requirementId, personName, signedAt
 * 10. Diagrams: id, clientId, data (JSON)
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
    } else if (action === 'addSignoff') {
      output = addSignoff(data.requirementId, data.personName);
    } else if (action === 'removeSignoff') {
      output = removeSignoff(data.id);
    } else if (action === 'saveDiagram') {
      output = saveDiagram(data.clientId, data.diagramData);
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

  const categories = sheetToObjects(ss.getSheetByName('Categories'));
  const pipelines = sheetToObjects(ss.getSheetByName('Pipelines'));
  const clients = sheetToObjects(ss.getSheetByName('Clients'));
  const clientData = sheetToObjects(ss.getSheetByName('ClientData'));
  const sections = sheetToObjects(ss.getSheetByName('Sections'));
  const requirements = sheetToObjects(ss.getSheetByName('Requirements'));
  const bullets = sheetToObjects(ss.getSheetByName('Bullets'));
  const technologies = sheetToObjects(ss.getSheetByName('Technologies'));
  const signoffs = sheetToObjects(ss.getSheetByName('Signoffs'));
  const diagrams = sheetToObjects(ss.getSheetByName('Diagrams'));

  // Sort helper
  const bySort = (a, b) => (a.sortOrder || 0) - (b.sortOrder || 0);

  // Nest signoffs under requirements
  const requirementsWithData = requirements.map(req => ({
    ...req,
    bullets: bullets.filter(b => b.requirementId === req.id).sort(bySort),
    technologies: technologies.filter(t => t.requirementId === req.id),
    signoffs: signoffs.filter(s => s.requirementId === req.id)
  })).sort(bySort);

  // Nest requirements under sections
  const sectionsWithReqs = sections.map(sec => ({
    ...sec,
    requirements: requirementsWithData.filter(r => r.sectionId === sec.id)
  })).sort(bySort);

  // Nest sections, clientData, and diagrams under clients
  const clientsWithData = clients.map(client => ({
    ...client,
    data: clientData.filter(d => d.clientId === client.id),
    sections: sectionsWithReqs.filter(s => s.clientId === client.id),
    diagram: diagrams.find(d => d.clientId === client.id) || null
  }));

  // Nest clients under pipelines
  const pipelinesWithClients = pipelines.map(pipeline => ({
    ...pipeline,
    clients: clientsWithData.filter(c => c.pipelineId === pipeline.id)
  })).sort(bySort);

  // Nest pipelines under categories
  const result = categories.map(category => ({
    ...category,
    pipelines: pipelinesWithClients.filter(p => p.categoryId === category.id)
  })).sort(bySort);

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

function addSignoff(requirementId, personName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Signoffs');
  const id = 'signoff-' + new Date().getTime();
  const signedAt = new Date().toISOString();

  sheet.appendRow([id, requirementId, personName, signedAt]);

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    signoff: { id, requirementId, personName, signedAt }
  }));
}

function removeSignoff(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Signoffs');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true }));
}

function saveDiagram(clientId, diagramData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Diagrams');
  const data = sheet.getDataRange().getValues();
  const diagramJson = JSON.stringify(diagramData);

  // Check if diagram exists for this client
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === clientId) {
      sheet.getRange(i + 1, 3).setValue(diagramJson);
      return ContentService.createTextOutput(JSON.stringify({ success: true }));
    }
  }

  // Create new diagram entry
  const id = 'diagram-' + new Date().getTime();
  sheet.appendRow([id, clientId, diagramJson]);

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

      // Parse JSON columns
      if (['links', 'data'].includes(headers[j]) && typeof value === 'string' && value) {
        try {
          value = JSON.parse(value);
        } catch (e) {
          if (headers[j] === 'links') value = [];
          else value = null;
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
