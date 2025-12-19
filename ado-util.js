import axios from "axios";

export class AdoUtil {
  constructor(organization, project, pat) {
    this.organization = organization;
    this.project = project;
    this.pat = pat;
    this.baseUrl = `https://dev.azure.com/${organization}`;
    
    this.client = axios.create({
      baseURL: this.baseUrl,
      headers: {
        "Content-Type": "application/json-patch+json",
        "Authorization": `Basic ${Buffer.from(`:${pat}`).toString("base64")}`
      },
      timeout: 60000
    });
  }

  async createWorkItem(workItemType, fields, bypassRules = true) {
    const patchDocument = [];
    
    for (const [fieldPath, value] of Object.entries(fields)) {
      if (value !== null && value !== undefined) {
        patchDocument.push({
          op: "add",
          path: `/fields/${fieldPath}`,
          value: value
        });
      }
    }
    
    if (!patchDocument.some(field => field.path === "/fields/System.AreaPath")) {
      patchDocument.push({
        op: "add",
        path: "/fields/System.AreaPath",
        value: this.project
      });
    }
    
    if (!patchDocument.some(field => field.path === "/fields/System.IterationPath")) {
      patchDocument.push({
        op: "add",
        path: "/fields/System.IterationPath",
        value: this.project
      });
    }

    try {
      const url = `/${this.project}/_apis/wit/workitems/$${workItemType}?bypassRules=${bypassRules}&api-version=7.1`;
      const response = await this.client.post(url, patchDocument);
      
      console.log(`Created work item: ${response.data.id} - ${fields["System.Title"] || "No Title"}`);
      return response.data;
    } catch (error) {
      console.error(`Error creating work item: ${error.response?.status} - ${error.response?.data?.message || error.message}`);
      return null;
    }
  }

  async updateWorkItem(workItemId, fields, bypassRules = true) {
    const patchDocument = [];
    
    for (const [fieldPath, value] of Object.entries(fields)) {
      if (value !== null && value !== undefined) {
        patchDocument.push({
          op: "add",
          path: `/fields/${fieldPath}`,
          value: value
        });
      }
    }

    try {
      const url = `/${this.project}/_apis/wit/workitems/${workItemId}?bypassRules=${bypassRules}&api-version=7.1`;
      const response = await this.client.patch(url, patchDocument);
      
      console.log(`Updated work item: ${workItemId}`);
      return response.data;
    } catch (error) {
      console.error(`Error updating work item ${workItemId}: ${error.response?.status} - ${error.response?.data?.message || error.message}`);
      return null;
    }
  }

  async addComment(workItemId, commentText) {
    const patchDocument = [{
      op: "add",
      path: "/fields/System.History",
      value: commentText
    }];

    try {
      const url = `/${this.project}/_apis/wit/workitems/${workItemId}?api-version=7.1`;
      await this.client.patch(url, patchDocument);
      
      console.log(`Added comment to work item ${workItemId}`);
      return true;
    } catch (error) {
      console.error(`Error adding comment to ${workItemId}: ${error.response?.status} - ${error.message}`);
      return false;
    }
  }

  async uploadAttachment(fileContent, fileName) {
    try {
      const url = `/${this.project}/_apis/wit/attachments?fileName=${encodeURIComponent(fileName)}&api-version=7.1`;
      const response = await this.client.post(url, fileContent, {
        headers: {
          "Content-Type": "application/octet-stream"
        }
      });
      
      console.log(`Uploaded attachment: ${fileName}`);
      return response.data.url;
    } catch (error) {
      console.error(`Error uploading attachment ${fileName}: ${error.response?.status} - ${error.message}`);
      return null;
    }
  }

  async attachFileToWorkItem(workItemId, attachmentUrl, fileName) {
    const patchDocument = [{
      op: "add",
      path: "/relations/-",
      value: {
        rel: "AttachedFile",
        url: attachmentUrl,
        attributes: {
          comment: `Attachment: ${fileName}`
        }
      }
    }];

    try {
      const url = `/${this.project}/_apis/wit/workitems/${workItemId}?api-version=7.1`;
      await this.client.patch(url, patchDocument);
      
      console.log(`Attached file ${fileName} to work item ${workItemId}`);
      return true;
    } catch (error) {
      console.error(`Error attaching file to ${workItemId}: ${error.response?.status} - ${error.message}`);
      return false;
    }
  }

  async createLink(sourceWorkItemId, targetWorkItemId, linkType, comment = "") {
    if (sourceWorkItemId === targetWorkItemId) {
      console.log(`Skipping link: source and target are the same (${sourceWorkItemId})`);
      return false;
    }

    const patchDocument = [{
      op: "add",
      path: "/relations/-",
      value: {
        rel: linkType,
        url: `${this.baseUrl}/${this.project}/_apis/wit/workitems/${targetWorkItemId}`,
        attributes: {
          comment: comment || `Link: ${linkType}`
        }
      }
    }];

    try {
      const url = `/${this.project}/_apis/wit/workitems/${sourceWorkItemId}?api-version=7.1`;
      await this.client.patch(url, patchDocument);
      
      console.log(`Created link: ${sourceWorkItemId} -> ${targetWorkItemId} (${linkType})`);
      return true;
    } catch (error) {
      console.error(`Error creating link: ${error.response?.status} - ${error.message}`);
      return false;
    }
  }

  async getWorkItem(workItemId) {
    try {
      const url = `/${this.project}/_apis/wit/workitems/${workItemId}?api-version=7.1`;
      const response = await this.client.get(url);
      return response.data;
    } catch (error) {
      console.error(`Error getting work item ${workItemId}: ${error.response?.status} - ${error.message}`);
      return null;
    }
  }
}
