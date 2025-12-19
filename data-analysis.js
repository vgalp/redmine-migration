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
// SHARED UTILITIES
// ============================================================

async function listAllProjects() {
  const params = {
    limit: 100,
    include: "issue_custom_fields"
  };

  let allProjects = [];
  let offset = 0;
  let hasMore = true;
  
  while (hasMore) {
    params.offset = offset;
    const { data } = await http.get("/projects.json", { params });
    const projects = data.projects || [];
    allProjects = allProjects.concat(projects);
    
    console.log(`Fetched ${projects.length} projects (offset ${offset}). Total so far: ${allProjects.length}`);
    
    if (projects.length < params.limit || allProjects.length >= data.total_count) {
      hasMore = false;
    } else {
      offset += params.limit;
      await delay(cfg.runtime?.delay_ms || 100);
    }
  }
  
  return allProjects;
}

function buildProjectHierarchy(projects) {
  const projectMap = new Map();
  const hierarchy = [];

  for (const project of projects) {
    projectMap.set(project.id, {
      ...project,
      children: [],
      parent_name: null,
      level: 0,
      full_path: project.name
    });
  }

  for (const project of projects) {
    const projectInfo = projectMap.get(project.id);
    
    if (project.parent) {
      const parent = projectMap.get(project.parent.id);
      if (parent) {
        parent.children.push(projectInfo);
        projectInfo.parent_name = parent.name;
        projectInfo.level = parent.level + 1;
        projectInfo.full_path = `${parent.full_path} > ${project.name}`;
      }
    } else {
      hierarchy.push(projectInfo);
    }
  }
  
  return { projectMap, hierarchy };
}


// ============================================================
// CUSTOM FIELDS ANALYSIS
// ============================================================

async function listCustomFields() {
  try {
    const { data } = await http.get("/custom_fields.json");
    return data.custom_fields || [];
  } catch (err) {
    if (err.response?.status === 403) {
      console.error("Error: Admin privileges required to fetch custom field definitions");
      return [];
    }
    throw err;
  }
}

async function analyzeCustomFields(projects, projectMap) {
  console.log("\n=== CUSTOM FIELDS ANALYSIS ===\n");
  
  console.log("Fetching all custom field definitions...");
  const customFields = await listCustomFields();
  console.log(`Found ${customFields.length} custom fields\n`);

  const cfMap = new Map();
  for (const cf of customFields) {
    cfMap.set(cf.id, {
      id: cf.id,
      name: cf.name,
      field_format: cf.field_format,
      customized_type: cf.customized_type,
      is_required: cf.is_required,
      is_filter: cf.is_filter,
      searchable: cf.searchable,
      multiple: cf.multiple,
      visible: cf.visible,
      possible_values: cf.possible_values?.map(v => v.value) || []
    });
  }

  const customFieldUsage = new Map();
  
  for (const project of projects) {
    const projectCustomFields = project.issue_custom_fields || [];
    
    for (const cf of projectCustomFields) {
      if (!customFieldUsage.has(cf.id)) {
        customFieldUsage.set(cf.id, []);
      }
      const projectInfo = projectMap.get(project.id);
      customFieldUsage.get(cf.id).push({
        project_id: project.id,
        project_identifier: project.identifier,
        project_name: project.name,
        parent_id: project.parent?.id || null,
        parent_name: projectInfo.parent_name,
        level: projectInfo.level,
        full_path: projectInfo.full_path,
        is_subproject: project.parent ? "Yes" : "No"
      });
    }
  }

  const projectCustomFieldSets = new Map();
  for (const project of projects) {
    const cfSet = new Set();
    for (const cf of (project.issue_custom_fields || [])) {
      cfSet.add(cf.id);
    }
    projectCustomFieldSets.set(project.id, cfSet);
  }

  return { customFields, cfMap, customFieldUsage, projectCustomFieldSets };
}


// ============================================================
// CUSTOM QUERIES ANALYSIS
// ============================================================

async function getAllCustomQueries() {
  try {
    let allQueries = [];
    let offset = 0;
    const limit = 100;
    let hasMore = true;
    
    while (hasMore) {
      const { data } = await http.get('/queries.json', { 
        params: { limit, offset } 
      });
      
      const queries = data.queries || [];
      allQueries = allQueries.concat(queries);
      
      console.log(`Fetched ${queries.length} queries (offset ${offset}). Total so far: ${allQueries.length}`);
      
      if (queries.length < limit || allQueries.length >= data.total_count) {
        hasMore = false;
      } else {
        offset += limit;
        await delay(cfg.runtime?.delay_ms || 100);
      }
    }
    
    return allQueries;
  } catch (err) {
    console.error('Error fetching global queries:', err.message);
    if (err.response) {
      console.error('Response status:', err.response.status);
      console.error('Response data:', err.response.data);
    }
    return [];
  }
}

async function getQueryDetails(queryId) {
  try {
    const { data } = await http.get(`/queries/${queryId}.json`);
    return data.query || null;
  } catch (err) {
    try {
      const { data: issuesData } = await http.get('/issues.json', {
        params: { query_id: queryId, limit: 1 }
      });
      
      return {
        accessible: true,
        total_issues: issuesData.total_count,
        note: "Query definition not accessible via API - only execution available"
      };
    } catch (issueErr) {
      return {
        accessible: false,
        note: "Query not accessible"
      };
    }
  }
}

async function analyzeCustomQueries(projects, projectMap) {
  console.log("\n=== CUSTOM QUERIES ANALYSIS ===\n");
  
  const allQueriesData = [];
  
  console.log("Fetching all queries from Redmine...");
  const queries = await getAllCustomQueries();
  
  console.log(`\nTotal queries found: ${queries.length}\n`);
  
  if (queries.length === 0) {
    console.log("No queries found. This might be due to permissions or no queries exist.\n");
  }
  
  console.log("Fetching detailed information for each query...\n");
  
  for (let i = 0; i < queries.length; i++) {
    const query = queries[i];
    console.log(`[${i + 1}/${queries.length}] Fetching details for query ${query.id}: ${query.name}`);
    
    const details = await getQueryDetails(query.id);
    await delay(cfg.runtime?.delay_ms || 100);
    
    const projectId = query.project_id;
    const projectInfo = projectId ? projectMap.get(projectId) : null;
    
    allQueriesData.push({
      query_id: query.id,
      query_name: query.name,
      visibility: query.is_public ? "Public" : "Private",
      project_id: projectId || "(Global)",
      project_identifier: projectInfo?.identifier || "",
      project_name: projectInfo?.name || "(Global Query)",
      parent_name: projectInfo?.parent_name || "",
      level: projectInfo?.level || 0,
      full_path: projectInfo?.full_path || "(Global Query)",
      is_subproject: projectInfo && projectInfo.level > 0 ? "Yes" : "No",
      view_url: `${BASE}/projects/${projectInfo?.identifier || 'redmine'}/issues?query_id=${query.id}`,
      edit_url: `${BASE}/queries/${query.id}/edit`,
      total_issues: details?.total_issues !== undefined ? details.total_issues : ""
    });
  }
  
  console.log(`\nCompleted fetching details for all queries.\n`);

  return { allQueriesData };
}


// ============================================================
// EXCEL GENERATION
// ============================================================

function createCustomFieldsWorksheets(wb, customFields, cfMap, customFieldUsage, projectCustomFieldSets, projects, projectMap, hierarchy) {
  // Worksheet: Custom Fields Overview
  const wsOverview = wb.addWorksheet("Custom Fields Overview");
  wsOverview.columns = [
    { header: "Custom Field ID", key: "id", width: 16 },
    { header: "Name", key: "name", width: 40 },
    { header: "Field Format", key: "field_format", width: 20 },
    { header: "Customized Type", key: "customized_type", width: 20 },
    { header: "Required", key: "is_required", width: 12 },
    { header: "Multiple", key: "multiple", width: 12 },
    { header: "Visible", key: "visible", width: 12 },
    { header: "Searchable", key: "searchable", width: 12 },
    { header: "Filter", key: "is_filter", width: 12 },
    { header: "Total Projects", key: "projects_count", width: 16 },
    { header: "Top-Level Projects", key: "top_level_count", width: 20 },
    { header: "Subprojects", key: "subproject_count", width: 16 },
    { header: "Possible Values", key: "possible_values", width: 60 }
  ];

  for (const cf of customFields) {
    const projectsUsing = customFieldUsage.get(cf.id) || [];
    const topLevel = projectsUsing.filter(p => p.is_subproject === "No").length;
    const subprojects = projectsUsing.filter(p => p.is_subproject === "Yes").length;
    
    wsOverview.addRow({
      id: cf.id,
      name: cf.name,
      field_format: cf.field_format,
      customized_type: cf.customized_type,
      is_required: cf.is_required ? "Yes" : "No",
      multiple: cf.multiple ? "Yes" : "No",
      visible: cf.visible ? "Yes" : "No",
      searchable: cf.searchable ? "Yes" : "No",
      is_filter: cf.is_filter ? "Yes" : "No",
      projects_count: projectsUsing.length,
      top_level_count: topLevel,
      subproject_count: subprojects,
      possible_values: cf.possible_values?.map(v => v.value).join("; ") || ""
    });
  }

  wsOverview.getRow(1).font = { bold: true };
  wsOverview.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Worksheet: Custom Fields by Project
  const wsByProject = wb.addWorksheet("Custom Fields by Project");
  wsByProject.columns = [
    { header: "Project ID", key: "project_id", width: 12 },
    { header: "Project Identifier", key: "project_identifier", width: 40 },
    { header: "Project Name", key: "project_name", width: 40 },
    { header: "Is Subproject", key: "is_subproject", width: 15 },
    { header: "Parent Project", key: "parent_name", width: 40 },
    { header: "Level", key: "level", width: 10 },
    { header: "Full Path", key: "full_path", width: 80 },
    { header: "Custom Field ID", key: "custom_field_id", width: 16 },
    { header: "Custom Field Name", key: "custom_field_name", width: 40 },
    { header: "Field Format", key: "field_format", width: 20 },
    { header: "Required", key: "is_required", width: 12 },
    { header: "Multiple", key: "multiple", width: 12 }
  ];

  for (const project of projects) {
    const projectCustomFields = project.issue_custom_fields || [];
    const projectInfo = projectMap.get(project.id);
    
    for (const cf of projectCustomFields) {
      const cfDetails = cfMap.get(cf.id);
      wsByProject.addRow({
        project_id: project.id,
        project_identifier: project.identifier,
        project_name: project.name,
        is_subproject: project.parent ? "Yes" : "No",
        parent_name: projectInfo.parent_name || "",
        level: projectInfo.level,
        full_path: projectInfo.full_path,
        custom_field_id: cf.id,
        custom_field_name: cf.name,
        field_format: cfDetails?.field_format || "",
        is_required: cfDetails?.is_required ? "Yes" : "No",
        multiple: cfDetails?.multiple ? "Yes" : "No"
      });
    }
  }

  wsByProject.getRow(1).font = { bold: true };
  wsByProject.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Worksheet: Projects by Custom Field
  const wsByCustomField = wb.addWorksheet("Projects by Custom Field");
  wsByCustomField.columns = [
    { header: "Custom Field ID", key: "custom_field_id", width: 16 },
    { header: "Custom Field Name", key: "custom_field_name", width: 40 },
    { header: "Field Format", key: "field_format", width: 20 },
    { header: "Project ID", key: "project_id", width: 12 },
    { header: "Project Identifier", key: "project_identifier", width: 40 },
    { header: "Project Name", key: "project_name", width: 40 },
    { header: "Is Subproject", key: "is_subproject", width: 15 },
    { header: "Parent Project", key: "parent_name", width: 40 },
    { header: "Level", key: "level", width: 10 },
    { header: "Full Path", key: "full_path", width: 80 }
  ];

  for (const [cfId, projectsList] of customFieldUsage.entries()) {
    const cfDetails = cfMap.get(cfId);
    const cfName = cfDetails?.name || `Unknown (ID: ${cfId})`;
    const fieldFormat = cfDetails?.field_format || "";
    
    for (const proj of projectsList) {
      wsByCustomField.addRow({
        custom_field_id: cfId,
        custom_field_name: cfName,
        field_format: fieldFormat,
        project_id: proj.project_id,
        project_identifier: proj.project_identifier,
        project_name: proj.project_name,
        is_subproject: proj.is_subproject,
        parent_name: proj.parent_name || "",
        level: proj.level,
        full_path: proj.full_path
      });
    }
  }

  wsByCustomField.getRow(1).font = { bold: true };
  wsByCustomField.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Worksheet: Parent-Child Custom Field Comparison
  const wsParentChild = wb.addWorksheet("Parent-Child Field Comparison");
  wsParentChild.columns = [
    { header: "Parent Project ID", key: "parent_id", width: 16 },
    { header: "Parent Project Name", key: "parent_name", width: 40 },
    { header: "Child Project ID", key: "child_id", width: 16 },
    { header: "Child Project Name", key: "child_name", width: 40 },
    { header: "Child Full Path", key: "child_full_path", width: 80 },
    { header: "Custom Field ID", key: "custom_field_id", width: 16 },
    { header: "Custom Field Name", key: "custom_field_name", width: 40 },
    { header: "In Parent", key: "in_parent", width: 12 },
    { header: "In Child", key: "in_child", width: 12 },
    { header: "Inheritance Status", key: "status", width: 30 }
  ];

  for (const project of projects) {
    if (project.parent) {
      const parentInfo = projectMap.get(project.parent.id);
      const childInfo = projectMap.get(project.id);
      
      const parentFields = projectCustomFieldSets.get(project.parent.id) || new Set();
      const childFields = projectCustomFieldSets.get(project.id) || new Set();
      
      const allFields = new Set([...parentFields, ...childFields]);
      
      for (const cfId of allFields) {
        const inParent = parentFields.has(cfId);
        const inChild = childFields.has(cfId);
        const cfDetails = cfMap.get(cfId);
        
        let status;
        if (inParent && inChild) {
          status = "Inherited (in both)";
        } else if (!inParent && inChild) {
          status = "Child only";
        } else if (inParent && !inChild) {
          status = "Parent only";
        }
        
        wsParentChild.addRow({
          parent_id: project.parent.id,
          parent_name: parentInfo?.name || project.parent.name,
          child_id: project.id,
          child_name: project.name,
          child_full_path: childInfo.full_path,
          custom_field_id: cfId,
          custom_field_name: cfDetails?.name || `Unknown (ID: ${cfId})`,
          in_parent: inParent ? "Yes" : "No",
          in_child: inChild ? "Yes" : "No",
          status: status
        });
      }
    }
  }

  wsParentChild.getRow(1).font = { bold: true };
  wsParentChild.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Worksheet: Field Inheritance Summary
  const wsInheritance = wb.addWorksheet("Field Inheritance Summary");
  wsInheritance.columns = [
    { header: "Custom Field ID", key: "custom_field_id", width: 16 },
    { header: "Custom Field Name", key: "custom_field_name", width: 40 },
    { header: "Total Projects", key: "total_projects", width: 16 },
    { header: "Top-Level Only", key: "top_level_only", width: 18 },
    { header: "Subprojects Only", key: "subprojects_only", width: 18 },
    { header: "Both Levels", key: "both_levels", width: 16 },
    { header: "Parent-Child Pairs (Inherited)", key: "inherited_pairs", width: 30 },
    { header: "Parent-Child Pairs (Parent Only)", key: "parent_only_pairs", width: 30 },
    { header: "Parent-Child Pairs (Child Only)", key: "child_only_pairs", width: 30 }
  ];

  for (const cf of customFields) {
    const projectsUsing = customFieldUsage.get(cf.id) || [];
    const topLevelProjects = projectsUsing.filter(p => p.is_subproject === "No");
    const subprojects = projectsUsing.filter(p => p.is_subproject === "Yes");
    
    let inheritedPairs = 0;
    let parentOnlyPairs = 0;
    let childOnlyPairs = 0;
    
    for (const project of projects) {
      if (project.parent) {
        const parentFields = projectCustomFieldSets.get(project.parent.id) || new Set();
        const childFields = projectCustomFieldSets.get(project.id) || new Set();
        
        if (parentFields.has(cf.id) && childFields.has(cf.id)) {
          inheritedPairs++;
        } else if (parentFields.has(cf.id) && !childFields.has(cf.id)) {
          parentOnlyPairs++;
        } else if (!parentFields.has(cf.id) && childFields.has(cf.id)) {
          childOnlyPairs++;
        }
      }
    }
    
    wsInheritance.addRow({
      custom_field_id: cf.id,
      custom_field_name: cf.name,
      total_projects: projectsUsing.length,
      top_level_only: topLevelProjects.length - inheritedPairs,
      subprojects_only: subprojects.length - inheritedPairs,
      both_levels: topLevelProjects.length > 0 && subprojects.length > 0 ? "Yes" : "No",
      inherited_pairs: inheritedPairs,
      parent_only_pairs: parentOnlyPairs,
      child_only_pairs: childOnlyPairs
    });
  }

  wsInheritance.getRow(1).font = { bold: true };
  wsInheritance.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };
}

function createCustomQueriesWorksheets(wb, allQueriesData, projects, projectMap) {
  // Worksheet: All Queries
  const wsAllQueries = wb.addWorksheet("All Queries");
  wsAllQueries.columns = [
    { header: "Query ID", key: "query_id", width: 12 },
    { header: "Query Name", key: "query_name", width: 40 },
    { header: "Visibility", key: "visibility", width: 12 },
    { header: "Project ID", key: "project_id", width: 12 },
    { header: "Project Identifier", key: "project_identifier", width: 20 },
    { header: "Project Name", key: "project_name", width: 30 },
    { header: "Full Path", key: "full_path", width: 60 },
    { header: "Total Issues", key: "total_issues", width: 15 },
    { header: "View Query URL", key: "view_url", width: 80 },
    { header: "Edit Query URL", key: "edit_url", width: 80 }
  ];

  for (const queryData of allQueriesData) {
    wsAllQueries.addRow(queryData);
  }

  wsAllQueries.getRow(1).font = { bold: true };
  wsAllQueries.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Worksheet: Queries Summary by Project
  const wsSummary = wb.addWorksheet("Queries Summary by Project");
  wsSummary.columns = [
    { header: "Project ID", key: "project_id", width: 12 },
    { header: "Project Identifier", key: "project_identifier", width: 20 },
    { header: "Project Name", key: "project_name", width: 30 },
    { header: "Parent Project", key: "parent_name", width: 30 },
    { header: "Full Path", key: "full_path", width: 60 },
    { header: "Is Subproject", key: "is_subproject", width: 15 },
    { header: "Total Queries", key: "total_queries", width: 15 },
    { header: "Public Queries", key: "public_queries", width: 15 },
    { header: "Private Queries", key: "private_queries", width: 15 }
  ];

  const projectQueryMap = new Map();
  for (const project of projects) {
    const projectInfo = projectMap.get(project.id);
    projectQueryMap.set(project.id, {
      project_id: project.id,
      project_identifier: project.identifier,
      project_name: project.name,
      parent_name: projectInfo.parent_name || "",
      full_path: projectInfo.full_path,
      is_subproject: project.parent ? "Yes" : "No",
      total: 0,
      public: 0,
      private: 0
    });
  }
  
  for (const queryData of allQueriesData) {
    const key = queryData.project_id;
    if (key && key !== "(Global)" && projectQueryMap.has(key)) {
      const stats = projectQueryMap.get(key);
      stats.total++;
      if (queryData.visibility === "Public") {
        stats.public++;
      } else {
        stats.private++;
      }
    }
  }

  const sortedSummary = Array.from(projectQueryMap.values()).sort((a, b) => 
    a.full_path.localeCompare(b.full_path)
  );

  for (const stats of sortedSummary) {
    wsSummary.addRow({
      project_id: stats.project_id,
      project_identifier: stats.project_identifier,
      project_name: stats.project_name,
      parent_name: stats.parent_name,
      full_path: stats.full_path,
      is_subproject: stats.is_subproject,
      total_queries: stats.total,
      public_queries: stats.public,
      private_queries: stats.private
    });
  }

  wsSummary.getRow(1).font = { bold: true };
  wsSummary.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };
}

function createProjectWorksheets(wb, projects, projectMap, hierarchy) {
  // Worksheet: All Projects
  const wsAllProjects = wb.addWorksheet("All Projects");
  wsAllProjects.columns = [
    { header: "Project ID", key: "project_id", width: 12 },
    { header: "Project Identifier", key: "project_identifier", width: 20 },
    { header: "Project Name", key: "project_name", width: 30 },
    { header: "Parent Project", key: "parent_name", width: 30 },
    { header: "Level", key: "level", width: 10 },
    { header: "Full Path", key: "full_path", width: 60 },
    { header: "Is Subproject", key: "is_subproject", width: 15 },
    { header: "Custom Fields Count", key: "custom_fields_count", width: 20 }
  ];

  const sortedProjects = [...projects].sort((a, b) => {
    const pathA = projectMap.get(a.id).full_path;
    const pathB = projectMap.get(b.id).full_path;
    return pathA.localeCompare(pathB);
  });

  for (const project of sortedProjects) {
    const projectInfo = projectMap.get(project.id);
    
    wsAllProjects.addRow({
      project_id: project.id,
      project_identifier: project.identifier,
      project_name: project.name,
      parent_name: projectInfo.parent_name || "",
      level: projectInfo.level,
      full_path: projectInfo.full_path,
      is_subproject: project.parent ? "Yes" : "No",
      custom_fields_count: project.issue_custom_fields?.length || 0
    });
  }

  wsAllProjects.getRow(1).font = { bold: true };
  wsAllProjects.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };

  // Worksheet: Project Hierarchy
  const wsHierarchy = wb.addWorksheet("Project Hierarchy");
  wsHierarchy.columns = [
    { header: "Project ID", key: "project_id", width: 12 },
    { header: "Project Identifier", key: "project_identifier", width: 40 },
    { header: "Project Name", key: "project_name", width: 40 },
    { header: "Level", key: "level", width: 10 },
    { header: "Is Subproject", key: "is_subproject", width: 15 },
    { header: "Parent ID", key: "parent_id", width: 12 },
    { header: "Parent Name", key: "parent_name", width: 40 },
    { header: "Full Path", key: "full_path", width: 80 },
    { header: "Custom Fields Count", key: "custom_fields_count", width: 20 }
  ];

  function addProjectsToSheet(projectsList, sheet) {
    for (const project of projectsList) {
      sheet.addRow({
        project_id: project.id,
        project_identifier: project.identifier,
        project_name: project.name,
        level: project.level,
        is_subproject: project.parent ? "Yes" : "No",
        parent_id: project.parent?.id || "",
        parent_name: project.parent_name || "",
        full_path: project.full_path,
        custom_fields_count: project.issue_custom_fields?.length || 0
      });
      
      if (project.children.length > 0) {
        addProjectsToSheet(project.children, sheet);
      }
    }
  }

  addProjectsToSheet(hierarchy, wsHierarchy);

  wsHierarchy.getRow(1).font = { bold: true };
  wsHierarchy.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };
}

function createSummaryWorksheet(wb, customFields, customFieldUsage, allQueriesData, projects, hierarchy) {
  const wsSummary = wb.addWorksheet("Summary");
  wsSummary.columns = [
    { header: "Metric", key: "metric", width: 50 },
    { header: "Value", key: "value", width: 20 }
  ];

  const topLevelCount = hierarchy.length;
  const subprojectCount = projects.length - topLevelCount;

  wsSummary.addRow({ metric: "=== PROJECTS ===", value: "" });
  wsSummary.addRow({ metric: "Total Projects", value: projects.length });
  wsSummary.addRow({ metric: "Top-Level Projects", value: topLevelCount });
  wsSummary.addRow({ metric: "Subprojects", value: subprojectCount });
  wsSummary.addRow({ metric: "", value: "" });
  
  wsSummary.addRow({ metric: "=== CUSTOM FIELDS ===", value: "" });
  wsSummary.addRow({ metric: "Total Custom Fields", value: customFields.length });
  wsSummary.addRow({ metric: "Custom Fields Used in Projects", value: customFieldUsage.size });
  wsSummary.addRow({ metric: "Custom Fields Not Used", value: customFields.length - customFieldUsage.size });
  wsSummary.addRow({ metric: "", value: "" });
  
  wsSummary.addRow({ metric: "Custom Fields by Type", value: "" });
  const typeCount = new Map();
  for (const cf of customFields) {
    const type = cf.field_format || "unknown";
    typeCount.set(type, (typeCount.get(type) || 0) + 1);
  }
  for (const [type, count] of typeCount.entries()) {
    wsSummary.addRow({ metric: `  - ${type}`, value: count });
  }
  
  wsSummary.addRow({ metric: "", value: "" });
  wsSummary.addRow({ metric: "=== CUSTOM QUERIES ===", value: "" });
  wsSummary.addRow({ metric: "Total Queries", value: allQueriesData.length });
  
  const publicQueries = allQueriesData.filter(q => q.visibility === "Public").length;
  const privateQueries = allQueriesData.filter(q => q.visibility === "Private").length;
  const globalQueries = allQueriesData.filter(q => q.project_id === "(Global)").length;
  const projectQueries = allQueriesData.filter(q => q.project_id !== "(Global)").length;
  
  wsSummary.addRow({ metric: "Public Queries", value: publicQueries });
  wsSummary.addRow({ metric: "Private Queries", value: privateQueries });
  wsSummary.addRow({ metric: "Global Queries", value: globalQueries });
  wsSummary.addRow({ metric: "Project-specific Queries", value: projectQueries });

  wsSummary.getRow(1).font = { bold: true };
  wsSummary.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAD3" }
  };
}


// ============================================================
// MAIN
// ============================================================

async function main() {
  console.log("=== REDMINE DATA ANALYSIS ===\n");

  try {
    console.log("Fetching all projects...");
    const projects = await listAllProjects();
    console.log(`Found ${projects.length} projects\n`);

    console.log("Building project hierarchy...");
    const { projectMap, hierarchy } = buildProjectHierarchy(projects);
    const topLevelCount = hierarchy.length;
    const subprojectCount = projects.length - topLevelCount;
    console.log(`${topLevelCount} top-level projects, ${subprojectCount} subprojects\n`);

    // Analyze Custom Fields
    const { customFields, cfMap, customFieldUsage, projectCustomFieldSets } = 
      await analyzeCustomFields(projects, projectMap);

    // Analyze Custom Queries
    const { allQueriesData } = await analyzeCustomQueries(projects, projectMap);

    // Create Excel workbook
    console.log("\n=== GENERATING EXCEL REPORT ===\n");
    const wb = new ExcelJS.Workbook();

    // Add Summary first
    createSummaryWorksheet(wb, customFields, customFieldUsage, allQueriesData, projects, hierarchy);

    // Add Project worksheets
    createProjectWorksheets(wb, projects, projectMap, hierarchy);

    // Add Custom Fields worksheets
    createCustomFieldsWorksheets(wb, customFields, cfMap, customFieldUsage, projectCustomFieldSets, projects, projectMap, hierarchy);

    // Add Custom Queries worksheets
    createCustomQueriesWorksheets(wb, allQueriesData, projects, projectMap);

    // Save the workbook
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outputPath = cfg.output?.data_analysis_excel_path || `redmine-data-analysis_${timestamp}.xlsx`;
    await wb.xlsx.writeFile(outputPath);
    
    console.log(`\n✓ Excel file saved: ${path.resolve(outputPath)}`);
    console.log(`\nWorksheets created:`);
    console.log(`  1. Summary - Overall statistics`);
    console.log(`  2. All Projects - ${projects.length} projects`);
    console.log(`  3. Project Hierarchy - Hierarchical view`);
    console.log(`  4. Custom Fields Overview - ${customFields.length} custom fields`);
    console.log(`  5. Custom Fields by Project - Field usage details`);
    console.log(`  6. Projects by Custom Field - Project lists per field`);
    console.log(`  7. Parent-Child Field Comparison - Inheritance analysis`);
    console.log(`  8. Field Inheritance Summary - Inheritance statistics`);
    console.log(`  9. All Queries - ${allQueriesData.length} queries`);
    console.log(` 10. Queries Summary by Project - Query counts per project`);
    
    console.log(`\nKey Statistics:`);
    console.log(`  - Total Projects: ${projects.length} (${topLevelCount} top-level, ${subprojectCount} subprojects)`);
    console.log(`  - Total Custom Fields: ${customFields.length}`);
    console.log(`  - Custom Fields in Use: ${customFieldUsage.size}`);
    console.log(`  - Total Queries: ${allQueriesData.length}`);
    console.log(`  - Public Queries: ${allQueriesData.filter(q => q.visibility === "Public").length}`);
    console.log(`  - Private Queries: ${allQueriesData.filter(q => q.visibility === "Private").length}`);

  } catch (err) {
    console.error("Fatal error:", err.message);
    if (err.response) {
      console.error("Response status:", err.response.status);
      console.error("Response data:", err.response.data);
    }
    process.exit(1);
  }
}

main();
