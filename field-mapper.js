import YAML from "yaml";
import fs from "fs";

export class FieldMapper {
  constructor(mappingFilePath = "field-mapping.yaml") {
    const content = fs.readFileSync(mappingFilePath, "utf-8");
    this.config = YAML.parse(content);
  }

  mapWorkItemType(redmineTracker) {
    const mapping = this.config.work_item_types?.[redmineTracker];
    return mapping || this.config.work_item_types?.default || "Issue";
  }

  mapStatus(redmineStatus, workItemType) {
    const workItemMappings = this.config.status_mappings?.[workItemType];
    if (workItemMappings && workItemMappings[redmineStatus]) {
      return workItemMappings[redmineStatus];
    }
    
    const defaultMappings = this.config.status_mappings?.default;
    if (defaultMappings && defaultMappings[redmineStatus]) {
      return defaultMappings[redmineStatus];
    }
    
    return "New";
  }

  mapPriority(redminePriority) {
    const mapping = this.config.priority_mappings?.[redminePriority];
    if (mapping) {
      return mapping;
    }
    return this.config.priority_mappings?.default || 3;
  }

  mapRelationType(redmineRelationType) {
    const mapping = this.config.relation_mappings?.[redmineRelationType];
    return mapping || "System.LinkTypes.Related";
  }

  mapRedmineIssueToAdoFields(issue) {
    const workItemType = this.mapWorkItemType(issue.tracker?.name || "Task");
    const fields = {};

    fields["System.Title"] = issue.subject || "Untitled";
    
    let description = "";
    if (this.config.migration_options?.preserve_redmine_id) {
      description += `<b>Migrated from Redmine Issue #${issue.id}</b><br/>`;
    }
    
    if (this.config.migration_options?.add_redmine_link && issue.id) {
      const redmineUrl = `${process.env.REDMINE_URL || ""}/issues/${issue.id}`;
      description += `<b>Original URL:</b> <a href="${redmineUrl}">${redmineUrl}</a><br/><br/>`;
      description += `<hr/><br/>`;
    }
    
    if (issue.description) {
      description += issue.description;
    }
    
    fields["System.Description"] = description;

    if (issue.status?.name) {
      fields["System.State"] = this.mapStatus(issue.status.name, workItemType);
    }

    if (issue.priority?.name) {
      fields["Microsoft.VSTS.Common.Priority"] = this.mapPriority(issue.priority.name);
    }

    if (issue.assigned_to?.name) {
      fields["System.AssignedTo"] = issue.assigned_to.name;
    }

    if (issue.start_date) {
      fields["Microsoft.VSTS.Scheduling.StartDate"] = issue.start_date;
    }

    if (issue.due_date) {
      fields["Microsoft.VSTS.Scheduling.DueDate"] = issue.due_date;
    }

    if (issue.created_on) {
      fields["System.CreatedDate"] = issue.created_on;
    }

    if (issue.updated_on) {
      fields["System.ChangedDate"] = issue.updated_on;
    }

    if (issue.closed_on) {
      fields["Microsoft.VSTS.Common.ClosedDate"] = issue.closed_on;
    }

    if (issue.estimated_hours) {
      fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] = issue.estimated_hours;
    }

    if (issue.done_ratio !== undefined && issue.done_ratio !== null) {
      fields["Microsoft.VSTS.Scheduling.CompletedWork"] = issue.done_ratio;
    }

    if (issue.custom_fields && Array.isArray(issue.custom_fields)) {
      for (const customField of issue.custom_fields) {
        const fieldName = customField.name;
        const adoFieldName = this.config.custom_field_mappings?.[fieldName];
        
        if (adoFieldName && customField.value) {
          fields[adoFieldName] = customField.value;
        }
      }
    }

    return { workItemType, fields };
  }

  shouldMigrateAttachments() {
    return this.config.migration_options?.migrate_attachments !== false;
  }

  shouldMigrateComments() {
    return this.config.migration_options?.migrate_comments !== false;
  }

  shouldMigrateRelations() {
    return this.config.migration_options?.migrate_relations !== false;
  }

  shouldMigrateSubtasks() {
    return this.config.migration_options?.migrate_subtasks !== false;
  }

  getDelayMs() {
    return this.config.migration_options?.delay_ms || 100;
  }

  getBatchSize() {
    return this.config.migration_options?.batch_size || 50;
  }
}
