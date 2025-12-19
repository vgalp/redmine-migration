import YAML from "yaml";
import fs from "fs";
import { RedmineUtil } from "./redmine-util.js";
import { AdoUtil } from "./ado-util.js";
import { FieldMapper } from "./field-mapper.js";

const config = YAML.parse(fs.readFileSync("config.yaml", "utf-8"));

const redmineUtil = new RedmineUtil(
  config.redmine.base_url,
  config.redmine.api_key
);

const adoUtil = new AdoUtil(
  config.azure_devops.organization,
  config.azure_devops.project,
  config.azure_devops.pat
);

const fieldMapper = new FieldMapper("field-mapping.yaml");

const issueIdMap = new Map();

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function migrateIssue(redmineIssue) {
  console.log(`\n--- Migrating Issue #${redmineIssue.id}: ${redmineIssue.subject} ---`);
  
  const { workItemType, fields } = fieldMapper.mapRedmineIssueToAdoFields(redmineIssue);
  
  const workItem = await adoUtil.createWorkItem(workItemType, fields, true);
  
  if (!workItem) {
    console.error(`Failed to create work item for Redmine issue #${redmineIssue.id}`);
    return null;
  }
  
  issueIdMap.set(redmineIssue.id, workItem.id);
  
  if (fieldMapper.shouldMigrateComments() && redmineIssue.journals && redmineIssue.journals.length > 0) {
    await migrateComments(redmineIssue.journals, workItem.id);
  }
  
  if (fieldMapper.shouldMigrateAttachments() && redmineIssue.attachments && redmineIssue.attachments.length > 0) {
    await migrateAttachments(redmineIssue.attachments, workItem.id);
  }
  
  return workItem;
}

async function migrateComments(journals, adoWorkItemId) {
  console.log(`Migrating ${journals.length} journal entries as comments...`);
  
  for (const journal of journals) {
    const comment = redmineUtil.formatJournalAsComment(journal);
    await adoUtil.addComment(adoWorkItemId, comment);
    await delay(fieldMapper.getDelayMs());
  }
}

async function migrateAttachments(attachments, adoWorkItemId) {
  console.log(`Migrating ${attachments.length} attachments...`);
  
  for (const attachment of attachments) {
    try {
      const fileContent = await redmineUtil.getAttachmentContent(attachment.content_url);
      
      if (fileContent) {
        const attachmentUrl = await adoUtil.uploadAttachment(fileContent, attachment.filename);
        
        if (attachmentUrl) {
          await adoUtil.attachFileToWorkItem(adoWorkItemId, attachmentUrl, attachment.filename);
        }
      }
      
      await delay(fieldMapper.getDelayMs());
    } catch (error) {
      console.error(`Error migrating attachment ${attachment.filename}: ${error.message}`);
    }
  }
}

async function migrateRelations(redmineIssues) {
  console.log(`\n--- Migrating Relations ---`);
  
  for (const issue of redmineIssues) {
    const sourceAdoId = issueIdMap.get(issue.id);
    if (!sourceAdoId) continue;
    
    if (fieldMapper.shouldMigrateSubtasks()) {
      const parentId = redmineUtil.getParentId(issue);
      if (parentId) {
        const parentAdoId = issueIdMap.get(parentId);
        if (parentAdoId) {
          await adoUtil.createLink(
            sourceAdoId,
            parentAdoId,
            "System.LinkTypes.Hierarchy-Reverse",
            "Parent-Child relationship from Redmine"
          );
          await delay(fieldMapper.getDelayMs());
        }
      }
    }
    
    if (fieldMapper.shouldMigrateRelations() && issue.relations && issue.relations.length > 0) {
      for (const relation of issue.relations) {
        const targetRedmineId = relation.issue_id || relation.issue_to_id;
        const targetAdoId = issueIdMap.get(targetRedmineId);
        
        if (targetAdoId) {
          const linkType = fieldMapper.mapRelationType(relation.relation_type);
          await adoUtil.createLink(
            sourceAdoId,
            targetAdoId,
            linkType,
            `${relation.relation_type} relationship from Redmine`
          );
          await delay(fieldMapper.getDelayMs());
        }
      }
    }
  }
}

async function migrateSingleIssue(issueId) {
  console.log(`\n========================================`);
  console.log(`Migrating Single Issue: #${issueId}`);
  console.log(`========================================\n`);
  
  const issue = await redmineUtil.getIssueDetails(issueId);
  
  if (!issue) {
    console.error(`Issue #${issueId} not found`);
    return;
  }
  
  const workItem = await migrateIssue(issue);
  
  if (workItem) {
    console.log(`\nMigration completed successfully!`);
    console.log(`Redmine Issue #${issueId} -> ADO Work Item #${workItem.id}`);
    console.log(`URL: https://dev.azure.com/${config.azure_devops.organization}/${config.azure_devops.project}/_workitems/edit/${workItem.id}`);
  }
}

async function migrateAllIssues() {
  console.log(`\n========================================`);
  console.log(`Migrating All Issues from Redmine`);
  console.log(`========================================\n`);
  
  const projectIdentifier = config.redmine.project_identifier || null;
  const issues = await redmineUtil.getAllIssues(projectIdentifier);
  
  if (issues.length === 0) {
    console.log("No issues found to migrate");
    return;
  }
  
  console.log(`\nFetching detailed information for ${issues.length} issues...`);
  const detailedIssues = [];
  
  for (const issue of issues) {
    const detailed = await redmineUtil.getIssueDetails(issue.id);
    if (detailed) {
      detailedIssues.push(detailed);
    }
    await delay(fieldMapper.getDelayMs());
  }
  
  console.log(`\nMigrating ${detailedIssues.length} issues to Azure DevOps...`);
  
  const batchSize = fieldMapper.getBatchSize();
  for (let i = 0; i < detailedIssues.length; i += batchSize) {
    const batch = detailedIssues.slice(i, i + batchSize);
    
    for (const issue of batch) {
      await migrateIssue(issue);
      await delay(fieldMapper.getDelayMs());
    }
    
    console.log(`\nProgress: ${Math.min(i + batchSize, detailedIssues.length)}/${detailedIssues.length} issues migrated`);
  }
  
  console.log(`\n--- Migrating Relationships ---`);
  await migrateRelations(detailedIssues);
  
  console.log(`\n========================================`);
  console.log(`Migration Completed!`);
  console.log(`========================================`);
  console.log(`Total issues migrated: ${issueIdMap.size}`);
  console.log(`\nIssue ID Mapping:`);
  
  for (const [redmineId, adoId] of issueIdMap.entries()) {
    console.log(`  Redmine #${redmineId} -> ADO #${adoId}`);
  }
  
  fs.writeFileSync(
    "migration-mapping.json",
    JSON.stringify(Object.fromEntries(issueIdMap), null, 2)
  );
  console.log(`\nMapping saved to migration-mapping.json`);
}

async function testConnection() {
  console.log(`\n========================================`);
  console.log(`Testing Connections`);
  console.log(`========================================\n`);
  
  console.log("Testing Redmine connection...");
  try {
    const issues = await redmineUtil.getAllIssues(config.redmine.project_identifier);
    console.log(`Success: Found ${issues.length} issues in Redmine`);
  } catch (error) {
    console.error(`Failed: ${error.message}`);
  }
  
  console.log("\nTesting Azure DevOps connection...");
  try {
    const testFields = {
      "System.Title": "Test Connection - DELETE ME",
      "System.Description": "This is a test work item to verify connectivity. Please delete."
    };
    
    const workItemTypes = ["Issue", "User Story", "Feature", "Task"];
    let workItem = null;
    
    for (const type of workItemTypes) {
      console.log(`Trying to create ${type}...`);
      workItem = await adoUtil.createWorkItem(type, testFields, true);
      if (workItem) {
        console.log(`Success: Created test ${type} #${workItem.id}`);
        console.log(`Please delete it manually from ADO`);
        break;
      }
    }
    
    if (!workItem) {
      console.error(`Failed: Could not create any work item type. Check your ADO process template.`);
    }
  } catch (error) {
    console.error(`Failed: ${error.message}`);
  }
}

async function main() {
  const args = process.argv.slice(2);
  const command = args[0];
  
  if (command === "test") {
    await testConnection();
  } else if (command === "single" && args[1]) {
    const issueId = parseInt(args[1]);
    await migrateSingleIssue(issueId);
  } else if (command === "all") {
    await migrateAllIssues();
  } else {
    console.log(`
Redmine to Azure DevOps Migration Tool
======================================

Usage:
  node migrate.js test              - Test connections to Redmine and ADO
  node migrate.js single <id>       - Migrate a single issue by ID
  node migrate.js all               - Migrate all issues from Redmine

Examples:
  node migrate.js test
  node migrate.js single 1234
  node migrate.js all

Configuration:
  - Redmine and ADO settings: config.yaml
  - Field mappings: field-mapping.yaml
    `);
  }
}

main().catch(error => {
  console.error(`\nFatal Error: ${error.message}`);
  console.error(error.stack);
  process.exit(1);
});
