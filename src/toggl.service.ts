import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

export interface DetailReportItem {
  id: number;
  pid: number;
  tid: number;
  uid: number;
  description: string;
  start: string;
  end: string;
  updated: string;
  dur: number;
  user: string;
  use_stop: boolean;
  client: string;
  project: string;
  task: string | null;
  billable: number;
  is_billable: boolean;
  cur: string;
  tags: Array<string>;
}

export class TogglService {
  private readonly baseUrl = "https://toggl.com/reports/api";
  private readonly apiVersion = "v2";
  private readonly defaultReportType = "details";
  private apiKey: string;
  private userAgent: string;
  private workspaceId: string;

  constructor(apiKey: string, userAgent: string, workspaceId: string) {
    this.apiKey = apiKey;
    this.userAgent = userAgent;
    this.workspaceId = workspaceId;
  }

  public getDetailReport(
    since: string = "",
    until: string = "",
    reportType: string = ""
  ): Array<DetailReportItem> {
    // set api url
    let apiUrl = `${this.baseUrl}/${this.apiVersion}/${
      reportType || this.defaultReportType
    }?user_agent=${this.userAgent}&workspace_id=${this.workspaceId}`;
    if (since) apiUrl += `&since=${since}`;
    if (until) apiUrl += `&until=${until}`;

    // request
    const headers = {
      Authorization: `Basic ${Utilities.base64Encode(
        `${this.apiKey}:api_token`
      )}`,
    };
    const options: URLFetchRequestOptions = {
      headers: headers,
    };

    return JSON.parse(
      UrlFetchApp.fetch(apiUrl, options).getContentText("UTF-8")
    ).data;
  }
}
