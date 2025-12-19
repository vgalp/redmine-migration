import axios from "axios";
import YAML from "yaml";
import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";


// ============================================================
// CONFIGURATION
// ============================================================

const cfg = YAML.parse(fs.readFileSync("config.yaml", "utf-8"));

// Azure DevOps Configuration
const ADO_ORG = cfg.azure_devops?.organization || process.env.ADO_ORGANIZATION;
const ADO_PROJECT = cfg.azure_devops?.project || process.env.ADO_PROJECT;
const ADO_PAT = cfg.azure_devops?.pat || process.env.ADO_PAT;

if (!ADO_ORG || !ADO_PROJECT || !ADO_PAT) {
  console.error("Error: Missing Azure DevOps configuration!");
  console.error("Please configure in config.yaml under 'azure_devops' section:");
  console.error("  azure_devops:");
  console.error("    organization: 'your-org-name'");
  console.error("    project: 'your-project-name'");
  console.error("    pat: 'your-personal-access-token'");
  console.error("\nOr set environment variables: ADO_ORGANIZATION, ADO_PROJECT, ADO_PAT");
  process.exit(1);
}

const BASE_URL = `https://dev.azure.com/${ADO_ORG}`;
const AUTH = Buffer.from(`:${ADO_PAT}`).toString('base64');

const http = axios.create({
  baseURL: BASE_URL,
  headers: {
    'Authorization': `Basic ${AUTH}`,
    'Content-Type': 'application/json'
  },
  timeout: 60_000
});

const delay = ms => new Promise(res => setTimeout(res, ms || 50));


// ============================================================
// AZURE DEVOPS API FUNCTIONS
// ============================================================

async function getWorkItemTypes() {
  try {
    const { data } = await http.get(`/${ADO_PROJECT}/_apis/wit/workitemtypes`, {
      params: { 'api-version': '7.1-preview.2' }
    });
    return data.value || [];
  } catch (err) {
    console.error("Error fetching work item types:", err.message);
    if (err.response) {
      console.error("Status:", err.response.status);
      console.error("Data:", err.response.data);
    }
    throw err;
  }
}

async function getFieldsForWorkItemType(workItemType) {
  try {
    const { data } = await http.get(`/${ADO_PROJECT}/_apis/wit/workitemtypes/${workItemType}`, {
      params: { 
        'api-version': '7.1-preview.2',
        '$expand': 'fields'
      }
    });
    return data.fields || [];
  } catch (err) {
    console.error(`Error fetching fields for ${workItemType}:`, err.message);
    return [];
  }
}

async function getAllFields() {
  try {
    const { data } = await http.get(`/_apis/wit/fields`, {
      params: { 'api-version': '7.1-preview.2' }
    });
    return data.value || [];
  } catch (err) {
    console.error("Error fetching all fields:", err.message);
    if (err.response) {
      console.error("Status:", err.response.status);
      console.error("Data:", err.response.data);
    }
    throw err;
  }
}

async function getProcesses() {
  try {
    const { data } = await http.get(`/_apis/work/processes`, {
      params: { 'api-version': '7.1-preview.2' }
    });
    return data.value || [];
  } catch (err) {
    console.error("Error fetching processes:", err.message);
    return [];
  }
}

async function getProjectProperties() {
  try {
    const { data } = await http.get(`/_apis/projects/${ADO_PROJECT}`, {
      params: { 
        'api-version': '7.1-preview.4',
        'includeCapabilities': true
      }
    });
    return data;
  } catch (err) {
    console.error("Error fetching project properties:", err.message);
    return null;
  }
}


// ============================================================
// EXCEL GENERATION
// ============================================================

function createAllFieldsWorksheet(wb, fields) {
  const ws = wb.addWorksheet("All Fields");
  ws.columns = [
    { header: "Field Name", key: "name", width: 40 },
    { header: "Reference Name", key: "referenceName", width: 50 },
    { header: "Type", key: "type", width: 20 },
    { header: "Usage", key: "usage", width: 15 },
    { header: "Read Only", key: "readOnly", width: 12 },
    { header: "Can Sort By", key: "canSortBy", width: 12 },
    { header: "Is Queryable", key: "isQueryable", width: 12 },
    { header: "Is Identity", key: "isIdentity", width: 12 },
    { header: "Is Picklist", key: "isPicklist", width: 12 },
    { header: "Is Deleted", key: "isDeleted", width: 12 },
    { header: "Supported Operations", key: "supportedOperations", width: 40 },
    { header: "Description", key: "description", width: 60 }
  ];

  for (const field of fields) {
    ws.addRow({
      name: field.name,
      referenceName: field.referenceName,
      type: field.type,
      usage: field.usage || "",
      readOnly: field.readOnly ? "Yes" : "No",
      canSortBy: field.canSortBy ? "Yes" : "No",
      isQueryable: field.isQueryable ? "Yes" : "No",
      isIdentity: field.isIdentity ? "Yes" : "No",
      isPicklist: field.isPicklist ? "Yes" : "No",
      isDeleted: field.isDeleted ? "Yes" : "No",
      supportedOperations: field.supportedOperations?.map(op => op.name || op.referenceName).join(", ") || "",
      description: field.description || ""
    });
  }

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  return ws;
}

function createCustomFieldsWorksheet(wb, fields) {
  const customFields = fields.filter(f => 
    f.usage === "workItemTypeExtension" || 
    f.referenceName?.startsWith("Custom.") ||
    !f.referenceName?.startsWith("System.") && 
    !f.referenceName?.startsWith("Microsoft.")
  );

  const ws = wb.addWorksheet("Custom Fields");
  ws.columns = [
    { header: "Field Name", key: "name", width: 40 },
    { header: "Reference Name", key: "referenceName", width: 50 },
    { header: "Type", key: "type", width: 20 },
    { header: "Usage", key: "usage", width: 15 },
    { header: "Read Only", key: "readOnly", width: 12 },
    { header: "Is Queryable", key: "isQueryable", width: 12 },
    { header: "Is Picklist", key: "isPicklist", width: 12 },
    { header: "Supported Operations", key: "supportedOperations", width: 40 },
    { header: "Description", key: "description", width: 60 }
  ];

  for (const field of customFields) {
    ws.addRow({
      name: field.name,
      referenceName: field.referenceName,
      type: field.type,
      usage: field.usage || "",
      readOnly: field.readOnly ? "Yes" : "No",
      isQueryable: field.isQueryable ? "Yes" : "No",
      isPicklist: field.isPicklist ? "Yes" : "No",
      supportedOperations: field.supportedOperations?.map(op => op.name || op.referenceName).join(", ") || "",
      description: field.description || ""
    });
  }

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  return ws;
}

function createFieldsByTypeWorksheet(wb, fields) {
  const ws = wb.addWorksheet("Fields by Type");
  ws.columns = [
    { header: "Field Type", key: "type", width: 30 },
    { header: "Field Name", key: "name", width: 40 },
    { header: "Reference Name", key: "referenceName", width: 50 },
    { header: "Usage", key: "usage", width: 15 },
    { header: "Is Custom", key: "isCustom", width: 12 },
    { header: "Read Only", key: "readOnly", width: 12 },
    { header: "Is Queryable", key: "isQueryable", width: 12 }
  ];

  // Group fields by type and sort
  const fieldsByType = {};
  for (const field of fields) {
    const type = field.type || "Unknown";
    if (!fieldsByType[type]) {
      fieldsByType[type] = [];
    }
    fieldsByType[type].push(field);
  }

  const sortedTypes = Object.keys(fieldsByType).sort();
  
  for (const type of sortedTypes) {
    const typeFields = fieldsByType[type].sort((a, b) => a.name.localeCompare(b.name));
    
    for (const field of typeFields) {
      const isCustom = field.usage === "workItemTypeExtension" || 
                      field.referenceName?.startsWith("Custom.") ||
                      (!field.referenceName?.startsWith("System.") && 
                       !field.referenceName?.startsWith("Microsoft."));
      
      ws.addRow({
        type: type,
        name: field.name,
        referenceName: field.referenceName,
        usage: field.usage || "",
        isCustom: isCustom ? "Yes" : "No",
        readOnly: field.readOnly ? "Yes" : "No",
        isQueryable: field.isQueryable ? "Yes" : "No"
      });
    }
  }

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  return ws;
}

function createFieldsByWorkItemTypeWorksheet(wb, workItemTypes, workItemTypeFields) {
  const ws = wb.addWorksheet("Fields by Work Item Type");
  ws.columns = [
    { header: "Work Item Type", key: "workItemType", width: 30 },
    { header: "Field Name", key: "name", width: 40 },
    { header: "Reference Name", key: "referenceName", width: 50 },
    { header: "Type", key: "type", width: 20 },
    { header: "Required", key: "alwaysRequired", width: 12 },
    { header: "Read Only", key: "readOnly", width: 12 },
    { header: "Default Value", key: "defaultValue", width: 30 },
    { header: "Help Text", key: "helpText", width: 60 }
  ];

  for (const wit of workItemTypes) {
    const fields = workItemTypeFields.get(wit.name) || [];
    
    for (const field of fields) {
      ws.addRow({
        workItemType: wit.name,
        name: field.name,
        referenceName: field.referenceName,
        type: field.type,
        alwaysRequired: field.alwaysRequired ? "Yes" : "No",
        readOnly: field.readOnly ? "Yes" : "No",
        defaultValue: field.defaultValue || "",
        helpText: field.helpText || ""
      });
    }
  }

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  return ws;
}

function createSummaryWorksheet(wb, fields, workItemTypes, projectInfo) {
  const ws = wb.addWorksheet("Summary");
  ws.columns = [
    { header: "Metric", key: "metric", width: 50 },
    { header: "Value", key: "value", width: 30 }
  ];

  ws.addRow({ metric: "=== PROJECT INFORMATION ===", value: "" });
  ws.addRow({ metric: "Organization", value: ADO_ORG });
  ws.addRow({ metric: "Project", value: ADO_PROJECT });
  ws.addRow({ metric: "Project ID", value: projectInfo?.id || "" });
  ws.addRow({ metric: "Project Description", value: projectInfo?.description || "" });
  ws.addRow({ metric: "Process Template", value: projectInfo?.capabilities?.processTemplate?.templateName || "" });
  ws.addRow({ metric: "", value: "" });

  ws.addRow({ metric: "=== FIELDS STATISTICS ===", value: "" });
  ws.addRow({ metric: "Total Fields", value: fields.length });

  const customFields = fields.filter(f => 
    f.usage === "workItemTypeExtension" || 
    f.referenceName?.startsWith("Custom.") ||
    (!f.referenceName?.startsWith("System.") && 
     !f.referenceName?.startsWith("Microsoft."))
  );
  ws.addRow({ metric: "Custom Fields", value: customFields.length });
  ws.addRow({ metric: "System Fields", value: fields.length - customFields.length });
  ws.addRow({ metric: "", value: "" });

  ws.addRow({ metric: "=== FIELDS BY TYPE ===", value: "" });
  const typeCount = {};
  for (const field of fields) {
    const type = field.type || "Unknown";
    typeCount[type] = (typeCount[type] || 0) + 1;
  }
  
  const sortedTypes = Object.keys(typeCount).sort();
  for (const type of sortedTypes) {
    ws.addRow({ metric: `  - ${type}`, value: typeCount[type] });
  }

  ws.addRow({ metric: "", value: "" });
  ws.addRow({ metric: "=== FIELD CHARACTERISTICS ===", value: "" });
  ws.addRow({ metric: "Read-Only Fields", value: fields.filter(f => f.readOnly).length });
  ws.addRow({ metric: "Queryable Fields", value: fields.filter(f => f.isQueryable).length });
  ws.addRow({ metric: "Picklist Fields", value: fields.filter(f => f.isPicklist).length });
  ws.addRow({ metric: "Identity Fields", value: fields.filter(f => f.isIdentity).length });
  ws.addRow({ metric: "Sortable Fields", value: fields.filter(f => f.canSortBy).length });

  ws.addRow({ metric: "", value: "" });
  ws.addRow({ metric: "=== WORK ITEM TYPES ===", value: "" });
  ws.addRow({ metric: "Total Work Item Types", value: workItemTypes.length });
  for (const wit of workItemTypes.sort((a, b) => a.name.localeCompare(b.name))) {
    ws.addRow({ metric: `  - ${wit.name}`, value: wit.description || "" });
  }

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Make summary metrics bold
  for (let i = 1; i <= ws.rowCount; i++) {
    const row = ws.getRow(i);
    const metricCell = row.getCell(1);
    if (metricCell.value && String(metricCell.value).startsWith("===")) {
      row.font = { bold: true };
    }
  }

  return ws;
}


// ============================================================
// MAIN
// ============================================================

async function main() {
  console.log("=== AZURE DEVOPS FIELDS EXPORT ===\n");
  console.log(`Organization: ${ADO_ORG}`);
  console.log(`Project: ${ADO_PROJECT}\n`);

  try {
    console.log("Fetching project information...");
    const projectInfo = await getProjectProperties();
    if (projectInfo) {
      console.log(`  - Process Template: ${projectInfo.capabilities?.processTemplate?.templateName || 'Unknown'}\n`);
    }

    console.log("Fetching all fields...");
    const fields = await getAllFields();
    console.log(`Found ${fields.length} total fields\n`);

    const customFields = fields.filter(f => 
      f.usage === "workItemTypeExtension" || 
      f.referenceName?.startsWith("Custom.") ||
      (!f.referenceName?.startsWith("System.") && 
       !f.referenceName?.startsWith("Microsoft."))
    );
    console.log(`  - Custom fields: ${customFields.length}`);
    console.log(`  - System fields: ${fields.length - customFields.length}\n`);

    console.log("Fetching work item types...");
    const workItemTypes = await getWorkItemTypes();
    console.log(`Found ${workItemTypes.length} work item types\n`);

    console.log("Fetching fields for each work item type...");
    const workItemTypeFields = new Map();
    for (const wit of workItemTypes) {
      console.log(`  - Fetching fields for: ${wit.name}`);
      const witFields = await getFieldsForWorkItemType(wit.name);
      workItemTypeFields.set(wit.name, witFields);
      await delay(100);
    }
    console.log("");

    // Create Excel workbook
    console.log("Generating Excel workbook...\n");
    const wb = new ExcelJS.Workbook();

    createSummaryWorksheet(wb, fields, workItemTypes, projectInfo);
    createAllFieldsWorksheet(wb, fields);
    createCustomFieldsWorksheet(wb, fields);
    createFieldsByTypeWorksheet(wb, fields);
    createFieldsByWorkItemTypeWorksheet(wb, workItemTypes, workItemTypeFields);

    // Save the workbook
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outputPath = cfg.output?.ado_fields_excel_path || `ado-fields_${ADO_PROJECT}_${timestamp}.xlsx`;
    await wb.xlsx.writeFile(outputPath);

    console.log(`✓ Excel file saved: ${path.resolve(outputPath)}`);
    console.log(`\nWorksheets created:`);
    console.log(`  1. Summary - Project and field statistics`);
    console.log(`  2. All Fields - ${fields.length} total fields`);
    console.log(`  3. Custom Fields - ${customFields.length} custom fields`);
    console.log(`  4. Fields by Type - Grouped by field type`);
    console.log(`  5. Fields by Work Item Type - ${workItemTypes.length} work item types`);

  } catch (err) {
    console.error("\nFatal error:", err.message);
    if (err.response) {
      console.error("Status:", err.response.status);
      console.error("Data:", err.response.data);
    }
    process.exit(1);
  }
}

main();
