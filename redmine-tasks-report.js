import https from "https";
import axios from "axios";
import YAML from "yaml";
import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";


const cfg = YAML.parse(fs.readFileSync("config.yaml", "utf-8"));
const BASE = cfg.redmine.base_url.replace(/\/+$/, "");
const HEADERS = { "X-Redmine-API-Key": cfg.redmine.api_key };

const httpsAgent = new https.Agent({ rejectUnauthorized: false });

const http = axios.create({ 
  baseURL: BASE, 
  headers: HEADERS, 
  timeout: 60_000,
  httpsAgent
});

const delay = ms => new Promise(res => setTimeout(res, ms || 50));


// ============================================================
// REDMINE API FUNCTIONS
// ============================================================

async function getAllIssues(projectIdentifier = null) {
  const params = {
    limit: 100,
    status_id: '*'  // Get all issues regardless of status
  };

  if (projectIdentifier) {
    params.project_id = projectIdentifier;
  }

  let allIssues = [];
  let offset = 0;
  let hasMore = true;
  
  while (hasMore) {
    params.offset = offset;
    const { data } = await http.get("/issues.json", { params });
    const issues = data.issues || [];
    allIssues = allIssues.concat(issues);
    
    if (issues.length < params.limit || allIssues.length >= data.total_count) {
      hasMore = false;
    } else {
      offset += params.limit;
      await delay(cfg.runtime?.delay_ms || 100);
    }
  }
  
  console.log(`Fetched ${allIssues.length} issues`);
  return allIssues;
}

async function getIssueDetails(issueId) {
  try {
    const { data } = await http.get(`/issues/${issueId}.json`, {
      params: {
        include: 'attachments,relations,children'
      }
    });
    return data.issue;
  } catch (err) {
    return null;
  }
}

async function processIssueBatch(issues, batchSize = 10) {
  const results = [];
  
  for (let i = 0; i < issues.length; i += batchSize) {
    const batch = issues.slice(i, i + batchSize);
    const batchResults = await Promise.all(
      batch.map(issue => getIssueDetails(issue.id))
    );
    results.push(...batchResults);
    
    // Progress update every 100 issues
    if ((i + batchSize) % 100 === 0 || i + batchSize >= issues.length) {
      console.log(`Processed ${Math.min(i + batchSize, issues.length)}/${issues.length} issues`);
    }
    
    await delay(cfg.runtime?.delay_ms || 50);
  }
  
  return results;
}


// ============================================================
// EXCEL GENERATION
// ============================================================

function createTasksReportWorksheet(wb, issuesData) {
  const ws = wb.addWorksheet("Tasks Report");
  ws.columns = [
    { header: "Issue ID", key: "id", width: 12 },
    { header: "Project", key: "project", width: 30 },
    { header: "Tracker", key: "tracker", width: 15 },
    { header: "Status", key: "status", width: 15 },
    { header: "Priority", key: "priority", width: 15 },
    { header: "Subject", key: "subject", width: 60 },
    { header: "Assigned To", key: "assigned_to", width: 25 },
    { header: "Author", key: "author", width: 25 },
    { header: "Created", key: "created_on", width: 20 },
    { header: "Updated", key: "updated_on", width: 20 },
    { header: "Has Attachments", key: "has_attachments", width: 18 },
    { header: "Attachment Count", key: "attachment_count", width: 18 },
    { header: "Has Relations", key: "has_relations", width: 18 },
    { header: "Relation Count", key: "relation_count", width: 18 },
    { header: "Relation Types", key: "relation_types", width: 40 },
    { header: "Has Subtasks", key: "has_subtasks", width: 18 },
    { header: "Subtask Count", key: "subtask_count", width: 18 },
    { header: "Subtask IDs", key: "subtask_ids", width: 30 },
    { header: "Issue URL", key: "url", width: 60 }
  ];

  for (const issue of issuesData) {
    ws.addRow({
      id: issue.id,
      project: issue.project,
      tracker: issue.tracker,
      status: issue.status,
      priority: issue.priority,
      subject: issue.subject,
      assigned_to: issue.assigned_to,
      author: issue.author,
      created_on: issue.created_on,
      updated_on: issue.updated_on,
      has_attachments: issue.has_attachments,
      attachment_count: issue.attachment_count,
      has_relations: issue.has_relations,
      relation_count: issue.relation_count,
      relation_types: issue.relation_types,
      has_subtasks: issue.has_subtasks,
      subtask_count: issue.subtask_count,
      subtask_ids: issue.subtask_ids,
      url: issue.url
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

function createAttachmentsWorksheet(wb, issuesData) {
  const ws = wb.addWorksheet("Issues with Attachments");
  ws.columns = [
    { header: "Issue ID", key: "id", width: 12 },
    { header: "Project", key: "project", width: 30 },
    { header: "Subject", key: "subject", width: 60 },
    { header: "Attachment Count", key: "attachment_count", width: 18 },
    { header: "Attachment Names", key: "attachment_names", width: 80 },
    { header: "Total Size (KB)", key: "total_size", width: 18 }
  ];

  const issuesWithAttachments = issuesData.filter(i => i.attachment_count > 0);

  for (const issue of issuesWithAttachments) {
    ws.addRow({
      id: issue.id,
      project: issue.project,
      subject: issue.subject,
      attachment_count: issue.attachment_count,
      attachment_names: issue.attachment_names,
      total_size: issue.attachment_total_size
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

function createRelationsWorksheet(wb, issuesData) {
  const ws = wb.addWorksheet("Issues with Relations");
  ws.columns = [
    { header: "Issue ID", key: "id", width: 12 },
    { header: "Project", key: "project", width: 30 },
    { header: "Subject", key: "subject", width: 60 },
    { header: "Relation Count", key: "relation_count", width: 18 },
    { header: "Relation Types", key: "relation_types", width: 40 },
    { header: "Related Issue IDs", key: "related_ids", width: 40 }
  ];

  const issuesWithRelations = issuesData.filter(i => i.relation_count > 0);

  for (const issue of issuesWithRelations) {
    ws.addRow({
      id: issue.id,
      project: issue.project,
      subject: issue.subject,
      relation_count: issue.relation_count,
      relation_types: issue.relation_types,
      related_ids: issue.related_ids
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

function createSubtasksWorksheet(wb, issuesData) {
  const ws = wb.addWorksheet("Issues with Subtasks");
  ws.columns = [
    { header: "Parent Issue ID", key: "id", width: 16 },
    { header: "Project", key: "project", width: 30 },
    { header: "Subject", key: "subject", width: 60 },
    { header: "Subtask Count", key: "subtask_count", width: 18 },
    { header: "Subtask IDs", key: "subtask_ids", width: 30 },
    { header: "Status", key: "status", width: 15 }
  ];

  const issuesWithSubtasks = issuesData.filter(i => i.subtask_count > 0);

  for (const issue of issuesWithSubtasks) {
    ws.addRow({
      id: issue.id,
      project: issue.project,
      subject: issue.subject,
      subtask_count: issue.subtask_count,
      subtask_ids: issue.subtask_ids,
      status: issue.status
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

function createSummaryWorksheet(wb, issuesData) {
  const ws = wb.addWorksheet("Summary");
  ws.columns = [
    { header: "Metric", key: "metric", width: 50 },
    { header: "Value", key: "value", width: 20 }
  ];

  const issuesWithAttachments = issuesData.filter(i => i.attachment_count > 0).length;
  const issuesWithRelations = issuesData.filter(i => i.relation_count > 0).length;
  const issuesWithSubtasks = issuesData.filter(i => i.subtask_count > 0).length;
  const totalAttachments = issuesData.reduce((sum, i) => sum + i.attachment_count, 0);
  const totalRelations = issuesData.reduce((sum, i) => sum + i.relation_count, 0);
  const totalSubtasks = issuesData.reduce((sum, i) => sum + i.subtask_count, 0);

  ws.addRow({ metric: "=== ISSUES OVERVIEW ===", value: "" });
  ws.addRow({ metric: "Total Issues", value: issuesData.length });
  ws.addRow({ metric: "", value: "" });

  ws.addRow({ metric: "=== ATTACHMENTS ===", value: "" });
  ws.addRow({ metric: "Issues with Attachments", value: issuesWithAttachments });
  ws.addRow({ metric: "Issues without Attachments", value: issuesData.length - issuesWithAttachments });
  ws.addRow({ metric: "Total Attachments", value: totalAttachments });
  ws.addRow({ metric: "Average Attachments per Issue (all)", value: (totalAttachments / issuesData.length).toFixed(2) });
  ws.addRow({ metric: "", value: "" });

  ws.addRow({ metric: "=== RELATIONS ===", value: "" });
  ws.addRow({ metric: "Issues with Relations", value: issuesWithRelations });
  ws.addRow({ metric: "Issues without Relations", value: issuesData.length - issuesWithRelations });
  ws.addRow({ metric: "Total Relations", value: totalRelations });
  ws.addRow({ metric: "", value: "" });

  ws.addRow({ metric: "=== SUBTASKS ===", value: "" });
  ws.addRow({ metric: "Parent Issues with Subtasks", value: issuesWithSubtasks });
  ws.addRow({ metric: "Issues without Subtasks", value: issuesData.length - issuesWithSubtasks });
  ws.addRow({ metric: "Total Subtasks", value: totalSubtasks });
  ws.addRow({ metric: "", value: "" });

  ws.addRow({ metric: "=== BY STATUS ===", value: "" });
  const statusCount = {};
  for (const issue of issuesData) {
    statusCount[issue.status] = (statusCount[issue.status] || 0) + 1;
  }
  for (const [status, count] of Object.entries(statusCount).sort()) {
    ws.addRow({ metric: `  - ${status}`, value: count });
  }

  ws.addRow({ metric: "", value: "" });
  ws.addRow({ metric: "=== BY TRACKER ===", value: "" });
  const trackerCount = {};
  for (const issue of issuesData) {
    trackerCount[issue.tracker] = (trackerCount[issue.tracker] || 0) + 1;
  }
  for (const [tracker, count] of Object.entries(trackerCount).sort()) {
    ws.addRow({ metric: `  - ${tracker}`, value: count });
  }

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Make section headers bold
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
  console.log("=== REDMINE TASKS REPORT ===\n");

  try {
    const projectIdentifier = cfg.redmine?.project_identifier || null;
    if (projectIdentifier) {
      console.log(`Fetching issues for project: ${projectIdentifier}`);
    } else {
      console.log("Fetching all issues from all projects...");
    }
    const issues = await getAllIssues(projectIdentifier);

    console.log(`Fetching detailed information (batch processing)...`);
    const batchSize = cfg.runtime?.batch_size || 10;
    const detailsList = await processIssueBatch(issues, batchSize);

    console.log("Processing issue data...");
    const issuesData = [];

    for (let i = 0; i < detailsList.length; i++) {
      const details = detailsList[i];
      if (!details) continue;

      const attachments = details.attachments || [];
      const relations = details.relations || [];
      const children = details.children || [];

      // Get unique relation types
      const relationTypes = [...new Set(relations.map(r => r.relation_type))].join(", ");
      
      // Get related issue IDs
      const relatedIds = relations.map(r => {
        if (r.issue_id === details.id) {
          return r.issue_to_id;
        }
        return r.issue_id;
      }).join(", ");

      issuesData.push({
        id: details.id,
        project: details.project?.name || "",
        tracker: details.tracker?.name || "",
        status: details.status?.name || "",
        priority: details.priority?.name || "",
        subject: details.subject,
        assigned_to: details.assigned_to?.name || "",
        author: details.author?.name || "",
        created_on: details.created_on,
        updated_on: details.updated_on,
        has_attachments: attachments.length > 0 ? "Yes" : "No",
        attachment_count: attachments.length,
        attachment_names: attachments.map(a => a.filename).join("; "),
        attachment_total_size: (attachments.reduce((sum, a) => sum + a.filesize, 0) / 1024).toFixed(2),
        has_relations: relations.length > 0 ? "Yes" : "No",
        relation_count: relations.length,
        relation_types: relationTypes,
        related_ids: relatedIds,
        has_subtasks: children.length > 0 ? "Yes" : "No",
        subtask_count: children.length,
        subtask_ids: children.map(c => c.id).join(", "),
        url: `${BASE}/issues/${details.id}`
      });
    }

    console.log("Generating Excel report...");
    const wb = new ExcelJS.Workbook();

    createSummaryWorksheet(wb, issuesData);
    createTasksReportWorksheet(wb, issuesData);
    createAttachmentsWorksheet(wb, issuesData);
    createRelationsWorksheet(wb, issuesData);
    createSubtasksWorksheet(wb, issuesData);

    // Save the workbook
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outputPath = cfg.output?.tasks_report_excel_path || `redmine-tasks-report_${timestamp}.xlsx`;
    await wb.xlsx.writeFile(outputPath);

    console.log(`\n✓ Report complete: ${path.resolve(outputPath)}`);
    console.log(`  Total Issues: ${issuesData.length}`);
    console.log(`  With Attachments: ${issuesData.filter(i => i.attachment_count > 0).length}`);
    console.log(`  With Relations: ${issuesData.filter(i => i.relation_count > 0).length}`);
    console.log(`  With Subtasks: ${issuesData.filter(i => i.subtask_count > 0).length}`);

  } catch (err) {
    console.error("\nError:", err.message);
    if (err.response) {
      console.error("Status:", err.response.status);
    }
    process.exit(1);
  }
}

main();
