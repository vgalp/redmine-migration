import https from "https";
import axios from "axios";

export class RedmineUtil {
  constructor(baseUrl, apiKey) {
    this.baseUrl = baseUrl.replace(/\/+$/, "");
    this.apiKey = apiKey;
    
    const httpsAgent = new https.Agent({ rejectUnauthorized: false });
    
    this.client = axios.create({
      baseURL: this.baseUrl,
      headers: {
        "X-Redmine-API-Key": this.apiKey
      },
      timeout: 60000,
      httpsAgent
    });
  }

  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms || 50));
  }

  async getAllIssues(projectIdentifier = null) {
    const params = {
      limit: 100,
      status_id: "*",
      include: "attachments,relations,children,journals"
    };

    if (projectIdentifier) {
      params.project_id = projectIdentifier;
    }

    let allIssues = [];
    let offset = 0;
    let hasMore = true;

    while (hasMore) {
      params.offset = offset;
      try {
        const { data } = await this.client.get("/issues.json", { params });
        const issues = data.issues || [];
        allIssues = allIssues.concat(issues);

        if (issues.length < params.limit || allIssues.length >= data.total_count) {
          hasMore = false;
        } else {
          offset += params.limit;
          await this.delay(100);
        }
      } catch (error) {
        console.error(`Error fetching issues: ${error.message}`);
        break;
      }
    }

    console.log(`Fetched ${allIssues.length} issues from Redmine`);
    return allIssues;
  }

  async getIssueDetails(issueId) {
    try {
      const { data } = await this.client.get(`/issues/${issueId}.json`, {
        params: {
          include: "attachments,relations,children,journals"
        }
      });
      return data.issue;
    } catch (error) {
      console.error(`Error fetching issue ${issueId}: ${error.message}`);
      return null;
    }
  }

  async getAttachmentContent(attachmentUrl) {
    try {
      const response = await this.client.get(attachmentUrl, {
        responseType: "arraybuffer",
        headers: {
          "X-Redmine-API-Key": this.apiKey
        }
      });
      return response.data;
    } catch (error) {
      console.error(`Error downloading attachment: ${error.message}`);
      return null;
    }
  }

  async getIssueHistory(issueId) {
    try {
      const issue = await this.getIssueDetails(issueId);
      return issue?.journals || [];
    } catch (error) {
      console.error(`Error fetching history for issue ${issueId}: ${error.message}`);
      return [];
    }
  }

  async processIssueBatch(issues, batchSize = 10, processor) {
    const results = [];

    for (let i = 0; i < issues.length; i += batchSize) {
      const batch = issues.slice(i, i + batchSize);
      const batchResults = await Promise.all(
        batch.map(issue => processor(issue))
      );
      results.push(...batchResults);

      if ((i + batchSize) % 100 === 0 || i + batchSize >= issues.length) {
        console.log(`Processed ${Math.min(i + batchSize, issues.length)}/${issues.length} issues`);
      }

      await this.delay(50);
    }

    return results;
  }

  formatJournalAsComment(journal) {
    const date = journal.created_on ? new Date(journal.created_on).toISOString() : "Unknown Date";
    const user = journal.user?.name || "Unknown User";
    const notes = journal.notes || "No notes provided.";
    
    let comment = `<b>Update by ${user} on ${date}</b><br/>`;
    
    if (journal.details && journal.details.length > 0) {
      comment += `<br/><b>Changes:</b><ul>`;
      for (const detail of journal.details) {
        const fieldName = detail.name || "Unknown field";
        const oldValue = detail.old_value || "(empty)";
        const newValue = detail.new_value || "(empty)";
        comment += `<li>${fieldName}: ${oldValue} -> ${newValue}</li>`;
      }
      comment += `</ul>`;
    }
    
    if (notes) {
      comment += `<br/><b>Notes:</b><br/>${notes}`;
    }
    
    return comment;
  }

  getRelations(issue) {
    return issue.relations || [];
  }

  getChildren(issue) {
    return issue.children || [];
  }

  getParentId(issue) {
    return issue.parent?.id || null;
  }
}
