import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/items';
import '@pnp/sp/lists';
import "@pnp/sp/folders";
import '@pnp/sp/attachments';
import '@pnp/sp/site-users/web';
import "@pnp/sp/files";

export interface ISiteInfo {title: string;  logoUrl: string;}

export class SpService {
  private sp: SPFI;
  private context: any;

  constructor(context: any) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  /* ================= READ ================= */
    public async getSiteInfo(): Promise<ISiteInfo> {
    const web = await this.sp.web.select(
      "Title",
      "SiteLogoUrl"
    )();

    return {
      title: web.Title,
      logoUrl: web.SiteLogoUrl??''
    };
  }
  
 
  public getList = (listName: string) => {
    return this.sp.web.lists.getByTitle(listName);
  };

  public createBatch = () => {
    return this.sp.batched();
  };
  
public async getPagedItems(
  listName: string,
  options: {
    pageSize: number;
    nextLink?: string;
    orderBy?: string;
    isDesc?: boolean;
    filter?: string;
    select?: string[];
    expand?: string[];
  },
  siteurl: string
): Promise<{
  items: any[];
  nextLink?: string;
}> {
  const {
    pageSize,
    nextLink,
    orderBy,
    isDesc,
    filter,
    select,
    expand
  } = options;

  let url = nextLink;

  if (!url) {
    let query = `?$top=${pageSize}`;

    if (orderBy) {
      query += `&$orderby=${orderBy} ${isDesc ? 'desc' : 'asc'}`;
    }

    if (filter) {
      query += `&$filter=${filter}`;
    }

    if (select?.length) {
      query += `&$select=${select.join(',')}`;
    }

    if (expand?.length) {
      query += `&$expand=${expand.join(',')}`;
    }

    url = `${siteurl}/_api/web/lists/getbytitle('${listName}')/items${query}`;
  }

  const res = await fetch(url!, {
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });

  const data = await res.json();
  console.log("🔵 RAW SHAREPOINT RESPONSE:", data);
  return {
    items: data.value,
    nextLink: data['odata.nextLink'] 
  };
}



  public async getItems(
    listTitle: string,
    selectFields: string[] = [],
    filter?: string,
    orderBy?: { field: string; ascending?: boolean },
    top: number = 2000,
    expandFields: string[] = []
  ): Promise<any[]> {
    try {
      let query = this.sp.web.lists.getByTitle(listTitle).items;

      if (selectFields.length) {
        query = query.select(...selectFields);
      }

      if (filter) {
        query = query.filter(filter);
      }
      
      if (expandFields.length) {
        query = query.expand(...expandFields);
      }

      if (orderBy) {
        query = query.orderBy(orderBy.field, orderBy.ascending ?? true);
      }

      return await query.top(top)();
    } catch (error) {
      console.error(`Error getting items from ${listTitle}`, error);
      throw error;
    }
  }

  public async getItemById(listTitle: string, id: any): Promise<any> {
    return this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(id)
      .select('*', 'Author/Id','Author/Title','Author/EMail')
      .expand('Author')();
  }
    public async getItemByIdd(listTitle: string, id: any): Promise<any> {
    return this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(id)
      .select('*', 'SiteLead/Id','SiteLead/Title','SiteLead/EMail',
        'SiteSupport/Id','SiteSupport/Title','SiteSupport/EMail',
        'SiteUpdateOwner/Id','SiteUpdateOwner/Title','SiteUpdateOwner/EMail','Author/Title')
      //.select('*', 'SiteLead/Id','SiteLead/Title','SiteSupport/Id','SiteSupport/Title','SiteUpdateOwner/Id','SiteUpdateOwner/Title')
      .expand('SiteLead','SiteSupport','SiteUpdateOwner','Author')();
  }
  private readonly itemSelectFields = [
    "*",
    "SiteLead/Id",
    "SiteLead/Title",
    // "SiteLead/EMail",
    "SiteSupport/Id",
    "SiteSupport/Title",
    // "SiteSupport/EMail",
    "SiteUpdateOwner/Id",
    "SiteUpdateOwner/Title",
    // "SiteUpdateOwner/EMail",
    "Author/Title",
  ];

  private readonly itemExpandFields = [
    "SiteLead",
    "SiteSupport",
    "SiteUpdateOwner",
    "Author",
  ];

  public async getListItemById(
    listTitle: string,
    id: number
  ): Promise<any> {
    return this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(id)
      .select(...this.itemSelectFields)
      .expand(...this.itemExpandFields)();
  }


  /* ================= NAVIGATION ================= */

  public async getNavigationItems(
    listTitle: string
  ): Promise<any[]> {
    try {
      const items = await this.getItems(
        listTitle,
        [
          'Id',
          'Title',
          'Url',
          'Icon',
          'Order',
          'IsEnabled',
          'Value',
          'ParentId' ,
          'IsAdmin' 
        ],
        'IsEnabled eq 1',
        { field: 'Order', ascending: true }
      );  
      
      const map = new Map<number, any>();
      const roots: any[] = [];
  
      items.forEach(item => {
        map.set(item.Id, { ...item, children: [] });
      });
  
      items.forEach(item => {
        const mappedItem = map.get(item.Id);
  
        if (item.ParentId && map.has(item.ParentId)) {      
          map.get(item.ParentId).children.push(mappedItem);
        } else {
          
          roots.push(mappedItem);
        }
      });
      
      const sortRecursive = (items: any[]) => {
        items.sort((a, b) => (a.Order ?? 0) - (b.Order ?? 0));
        items.forEach(i => i.children && sortRecursive(i.children));
      };
      const rootsSorted=sortRecursive(roots);
      return roots; 
    } catch (error) {
      console.error('Error loading navigation', {
        error,
        listTitle
      });
      throw error;
    }
  }
  /* ================= WRITE ================= */

  public async createItem(listTitle: string, payload: any): Promise<any> {
    const res = await this.sp.web.lists
      .getByTitle(listTitle)
      .items.add(payload);
    return res;
  }

  public async updateItem(
    listTitle: string,
    id: number,
    payload: any
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(id)
      .update(payload);
  }
  public async deleteItem(listTitle: string, itemId: number){
    try {
      const recycleId = await this.sp.web.lists
        .getByTitle(listTitle)
        .items.getById(itemId)
        .recycle();
    } catch (error) {
      console.error(`Error recycling item ${itemId} from ${listTitle}`, error);
      throw error;
    }
  }

  /* ================= ATTACHMENTS ================= */
 
  public async createFolder(relativeWebUrl:string,libraryName: string, folderName: string): Promise<void> {   

    const safeWebUrl = relativeWebUrl.endsWith('/') ? relativeWebUrl : `${relativeWebUrl}/`;
    const folderPath = `${safeWebUrl}${libraryName}/${folderName}`;

    try {
      await this.sp.web.folders.addUsingPath(folderPath)
    } catch (error: any) {
      if (error?.message?.includes("already exists")) {
        return;
      }
      throw error;
  
    }
  }
  public async uploadFileToFolderWithMetadata(
    libraryName: string,
    folderName: string,
    fileName: string,
    fileBuffer: ArrayBuffer,
    metadata: any
  ) {
    const folderPath = `${libraryName}/${folderName}`;
  
    const uploadResult = await this.sp.web
      .getFolderByServerRelativePath(folderPath)
      .files.addUsingPath(fileName, fileBuffer, { Overwrite: true });
    const fileQueryable = this.sp.web.getFileByServerRelativePath(uploadResult.ServerRelativeUrl);

    const item = await fileQueryable.getItem(); 

    await item.update(metadata);
  }
  public async getFilesFromFolder(libraryName: string, folderName: string) {
    const folderPath = `${libraryName}/${folderName}`;
  
    try {
      const files = await this.sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();
  
      return files;
    } catch {
      return [];
    }
  }
  
  
  public async deleteFile(serverRelativeUrl: string): Promise<void> {
    
    if (!serverRelativeUrl) {
      throw new Error("Invalid file path");
    }
    try {
      await this.sp.web
        .getFileByServerRelativePath(serverRelativeUrl)
        .recycle();
    } catch (error) {
      console.error("Error deleting file:", error);
      throw error;
  
    }
  }

  /* ================= USER ================= */

  public async getCurrentUserGroups(): Promise<string[]> {
    const groups = await this.sp.web.currentUser.groups();
    return groups.map((group: { Title: any; }) => group.Title);
  }
  
  public async isUserInGroup(groupName: string): Promise<boolean> {
    try {
      const groups = await this.getCurrentUserGroups();
      return groups.includes(groupName);
    } catch (error) {
      console.error('Failed to check user group membership', error);
      return false;
    }
  }

  public async getCurrentUser(): Promise<{
    id: number;
    loginName: string;
    displayName: string;
    email: string;
  }> {
    const user = await this.sp.web.currentUser();
    return {
      id: user.Id,
      loginName: user.LoginName,
      displayName: user.Title,
      email: user.Email,
    };
  }
  
  public async ensureUser(upn: string): Promise<any> {
    if (!upn) {
      throw new Error('UPN or email is required to ensure user');
    }

    try {
      const user = await this.sp.web.ensureUser(upn);
      return user; 
    } catch (error) {
      console.error(`Failed to ensure user: ${upn}`, error);
      throw error;
    }
  }
private userCache: Map<number, any> = new Map();

public async getUsersByIds(ids: number[]): Promise<any[]> {
  if (!ids || ids.length === 0) return [];


  const uniqueIds = Array.from(new Set(ids));

 
  const cachedResults: any[] = [];
  const missingIds: number[] = [];

  for (const id of uniqueIds) {
    if (this.userCache.has(id)) {
      cachedResults.push(this.userCache.get(id));
    } else {
      missingIds.push(id);
    }
  }

  let fetchedResults: any[] = [];


  if (missingIds.length > 0) {
    try {
      const filterString = missingIds.map(id => `Id eq ${id}`).join(" or ");

      const users = await this.sp.web.siteUsers
        .filter(filterString)
        .select("Id", "Title", "Email")();      

      fetchedResults = users.map((u: { Id: any; Title: any; Email: any; }) => ({
        id: u.Id,
        displayName: u.Title,
        email: u.Email,
        userPrincipalName: u.Email,
      }));
      fetchedResults.forEach(user => {
        this.userCache.set(user.id, user);
      });

    } catch (batchError) {
      console.warn("Batch user fetch failed, falling back", batchError);

     
      const fallbackResults = await Promise.all(
        missingIds.map(async (id) => {
          try {
            const u = await this.sp.web.siteUsers
              .getById(id)
              .select("Id", "Title", "Email")();

            const mapped = {
              id: u.Id,
              displayName: u.Title,
              email: u.Email,
              userPrincipalName: u.Email,
            };

            this.userCache.set(id, mapped);
            return mapped;

          } catch (err) {
            console.error(`User fetch failed for ID ${id}`, err);
            return null;
          }
        })
      );

      fetchedResults = fallbackResults.filter(Boolean);
    }
  }

  
  const finalResults = [...cachedResults, ...fetchedResults];

  const ordered = uniqueIds
    .map(id => finalResults.find(u => u.id === id))
    .filter(Boolean);

  return ordered;
} 

public async getLegalExport(context: any): Promise<any[]> {

  /* ==========================================
     CURRENT QUARTER
  =========================================== */

  const configItems = await this.getItems(
  "ConfigurationSetting"
);

const currentQuarterConfig = configItems.find(
  (item: any) =>
    item.Title === "CurrentQuarter"
);

const currentQuarter =
  currentQuarterConfig?.Value0 ??
  currentQuarterConfig?.Value ??
  "";

  /* ==========================================
     ADDITIONAL SUPPORT
     Build Outside Counsel dictionary
  =========================================== */

  const additionalSupportItems = await this.sp.web.lists
    .getByTitle("SiteAdditionalSupport")
    .items
    .select(
      "SiteID/Title",
      "AdditionalSupportRole",
      "AdditionalContact/LastName"
    )
    .expand(
      "SiteID",
      "AdditionalContact"
    )
    .top(5000)();

  const outsideCounselMap: Record<string, string> = {};

  additionalSupportItems.forEach((item: any) => {

    const siteId = item.SiteID?.Title;

    if (!siteId) {
      return;
    }

    if (!outsideCounselMap[siteId]) {
      outsideCounselMap[siteId] = "";
    }

    if (
      item.AdditionalSupportRole ===
      "Outside Counsel"
    ) {

      const lastName =
        item.AdditionalContact?.LastName || "";

      if (lastName) {

        outsideCounselMap[siteId] +=
          outsideCounselMap[siteId]
            ? `; ${lastName}`
            : lastName;
      }
    }
  });

  /* ==========================================
     SITE UPDATE
     Legacy code uses RecommendedChange
  =========================================== */

  const siteUpdateItems = await this.sp.web.lists
    .getByTitle("SiteUpdate")
    .items
    .select(
      "RecommendedChange",
      "SiteName/Title",
      "Quarter"
    )
    .expand("SiteName")
    .filter(`Quarter eq '${currentQuarter}'`)
    .top(5000)();

  const updateMap: Record<string, number> = {};

  siteUpdateItems.forEach((item: any) => {

    const siteId =
      item.SiteName?.Title;

    if (!siteId) {
      return;
    }

    if (updateMap[siteId] == null) {

      updateMap[siteId] =
        Number(item.RecommendedChange || 0);
    }
  });

  /* ==========================================
     CONTACTS
     Legacy fallback support
  =========================================== */

  const contacts = await this.sp.web.lists
    .getByTitle("Contacts")
    .items
    .select(
      "PfizerNetworkUserName",
      "LastName"
    )
    .top(5000)();

  const contactsMap: Record<string, string> = {};

  contacts.forEach((contact: any) => {

    const userName =
      contact.PfizerNetworkUserName;

    if (userName) {

      contactsMap[userName.toLowerCase()] =
        contact.LastName || "";
    }
  });

  /* ==========================================
     SITE LIST
     PRIMARY DATA SOURCE
  =========================================== */

  const sites = await this.sp.web.lists
    .getByTitle("Site")
    .items
    .select(
      "Title",
      "RSRSSiteName",
      "Claim_Type",
      "City",
      "State",
      "Country",
      "Site_Activity",
      "Site_Type",
      "Legacy_Company",

      "TotalProjectCost",
      "BETEstimatedCost",
      "TotalSpendingToDate",

      "AccountProjectCode",

      "RSRS_Update_Owner3",
      "Site_Lead",
      "Site_Support",

      "SiteLead/Title",
      "SiteSupport/Title",
      "SiteUpdateOwner/Title"
    )
    .expand(
      "SiteLead",
      "SiteSupport",
      "SiteUpdateOwner"
    )
    .top(5000)();

  const results = sites.map((site: any) => {

    /* ======================================
       SITE LEAD
    ======================================= */

    let siteLead = "";

    if (
      site.SiteLead &&
      Array.isArray(site.SiteLead.results) &&
      site.SiteLead.results.length > 0
    ) {

      siteLead =
        site.SiteLead.results[0].Title || "";
    }
    else if (site.SiteLead?.Title) {

      siteLead =
        site.SiteLead.Title;
    }
    else if (
      site.Site_Lead
    ) {

      siteLead =
        contactsMap[
          site.Site_Lead.toLowerCase()
        ] || site.Site_Lead;
    }

    if (siteLead.indexOf(",") > 0) {
      siteLead =
        siteLead.split(",")[0];
    }

    /* ======================================
       SITE SUPPORT
    ======================================= */

    let siteSupport = "";

    if (
      site.SiteSupport &&
      Array.isArray(site.SiteSupport.results) &&
      site.SiteSupport.results.length > 0
    ) {

      siteSupport =
        site.SiteSupport.results[0].Title || "";
    }
    else if (site.SiteSupport?.Title) {

      siteSupport =
        site.SiteSupport.Title;
    }
    else if (
      site.Site_Support
    ) {

      siteSupport =
        contactsMap[
          site.Site_Support.toLowerCase()
        ] || site.Site_Support;
    }

    if (siteSupport.indexOf(",") > 0) {
      siteSupport =
        siteSupport.split(",")[0];
    }

    /* ======================================
       UPDATE OWNER
    ======================================= */

    let siteUpdateOwner = "";

    if (
      site.SiteUpdateOwner &&
      Array.isArray(site.SiteUpdateOwner.results) &&
      site.SiteUpdateOwner.results.length > 0
    ) {

      siteUpdateOwner =
        site.SiteUpdateOwner.results[0].Title || "";
    }
    else if (
      site.SiteUpdateOwner?.Title
    ) {

      siteUpdateOwner =
        site.SiteUpdateOwner.Title;
    }
    else if (
      site.RSRS_Update_Owner3
    ) {

      siteUpdateOwner =
        contactsMap[
          site.RSRS_Update_Owner3.toLowerCase()
        ] ||
        site.RSRS_Update_Owner3;
    }

    if (
      siteUpdateOwner.indexOf(",") > 0
    ) {

      siteUpdateOwner =
        siteUpdateOwner.split(",")[0];
    }

    /* ======================================
       LEGACY COST LOGIC
    ======================================= */

    const siteUpdateChange =
      updateMap[site.Title] || 0;

    const totalProjectCost =
      Number(site.TotalProjectCost || 0) +
      Number(siteUpdateChange);

    const totalSpending =
      Number(site.TotalSpendingToDate || 0);

    const reserveBalance =
      totalProjectCost - totalSpending;

    return {

      Title:
        site.Title || "",

      RSRSSiteName:
        site.RSRSSiteName || "",

      Claim_Type:
        site.Claim_Type || "",

      Site_Type:
        site.Site_Type || "",

      Site_Activity:
        site.Site_Activity || "",

      Legacy_Company:
        site.Legacy_Company || "",

      City:
        site.City || "",

      State:
        site.State || "",

      Country:
        site.Country || "",

      AccountProjectCode:
        site.AccountProjectCode || "",

      SiteLead:
        siteLead,

      SiteSupport:
        siteSupport,

      SiteUpdateOwner:
        siteUpdateOwner,

      OutsideCounsel:
        outsideCounselMap[
          site.Title
        ] || "",

      TotalProjectCost:
        totalProjectCost.toFixed(2),

      BETEstimatedCost:
        Number(
          site.BETEstimatedCost || 0
        ).toFixed(2),

      TotalSpendingToDate:
        totalSpending.toFixed(2),

      ReserveBalance:
        reserveBalance.toFixed(2)
    };
  });

  results.sort(
    (a: any, b: any) =>
      (a.RSRSSiteName || "")
        .localeCompare(
          b.RSRSSiteName || ""
        )
  );

  return results;
}
public async getActiveContacts(context: any) {

  const data = await this.sp.web.lists
    .getByTitle("Contacts")
    .items
    .filter("User_Security_Level ne 'No Access'")();

  return data.map((r: any) => ({
    Title: r.Title,
    LastName: r.LastName,
    Contact_Type: r.Contact_Type,
    Email: r.ContactEmail,
    Phone: r.Telephone,
    Username: r.PfizerNetworkUserName,
    Security: r.User_Security_Level,
    Company: r.ContactCompany,
    Date: r.DateContactAdded
      ? r.DateContactAdded.split("T")[0]
      : ""
  }));
}

public async getNegativeReserve(
  context: any
): Promise<any[]> {

  const data = await this.sp.web.lists
    .getByTitle("Site")
    .items
    .select(
      "Title",
      "RSRSSiteName",
      "TotalProjectCost",
      "TotalSpendingToDate",
      "Site_Activity"
    )
    .top(5000)();

  const result = data
    .map((r: any) => {

      // Legacy logic
      const totalCost =
        Math.round(
          Number(r.TotalProjectCost || 0) * 1000
        );

      const spendingToDate =
        Math.round(
          Number(r.TotalSpendingToDate || 0) * 1000
        );

      const reserve =
        (totalCost - spendingToDate) / 1000;

      return {
        Title: r.Title || "",
        RSRSSiteName:
          r.RSRSSiteName || "",
        Reserve: reserve,
        Site_Activity:
          r.Site_Activity || ""
      };
    })
    .filter((item: any) => {

      // Legacy condition #1
      if (
        item.Site_Activity !== "Closed" &&
        item.Reserve < 0
      ) {
        return true;
      }

      // Legacy condition #2
      if (
        item.Site_Activity === "Closed" &&
        item.Reserve !== 0
      ) {
        return true;
      }

      return false;
    });

  result.sort((a: any, b: any) =>
    (a.Title || "").localeCompare(
      b.Title || ""
    )
  );

  return result;
}

/* public async getQuarterReview(context: any) {

  const data = await this.sp.web.lists
    .getByTitle("SiteAdditionalSupport")
    .items
    .select(
      "SiteID/Title",
      "SiteID/RSRSSiteName",
      "SiteID/Claim_Type",
      "SiteID/Site_Activity",
      "SiteID/Site_Type",
      "SiteID/Legacy_Company",
      "SiteID/Cost_Estimate_Method",
      "SiteID/TotalProjectCost",
      "SiteID/BETEstimatedCost",
      "SiteID/Project_Stage",
      "AdditionalSupportRole",
      "AdditionalContact/LastName",
      "IsActive"
    )
    .expand(
      "SiteID",
      "AdditionalContact"
    )
    .filter(
      "(IsActive eq 1) and ((SiteID/Site_Activity eq 'Open') or (SiteID/Site_Activity eq 'Open - Monitor Only'))"
    )
    .top(2000)
    ();

  // client-side sort
  data.sort((a: any, b: any) => {
    const claimCompare = (a.SiteID?.Claim_Type || "").localeCompare(
      b.SiteID?.Claim_Type || ""
    );

    if (claimCompare !== 0) {
      return claimCompare;
    }

    return (a.SiteID?.Title || "").localeCompare(
      b.SiteID?.Title || ""
    );
  });

  const map: any = {};

  data.forEach((r: any) => {
    const id = r.SiteID?.Title;

    if (!id) return;

    if (!map[id]) {
      map[id] = {
        Claim_Type: r.SiteID?.Claim_Type || "",        
        Title: r.SiteID?.Title || "",        
        RSRSSiteName: r.SiteID?.RSRSSiteName || "",        
        Site_Activity: r.SiteID?.Site_Activity || "",        
        Site_Type: r.SiteID?.Site_Type || "",        
        Legacy_Company: r.SiteID?.Legacy_Company || "",        
        Cost_Estimate_Method: r.SiteID?.Cost_Estimate_Method || "",
        Project_Stage:r.SiteID?.Project_Stage || "",

        TotalProjectCost: r.SiteID?.TotalProjectCost || 0,
        BETEstimatedCost:r.SiteID?.BETEstimatedCost || 0,
        
        AdditionalSupportRole:r.AdditionalSupportRole || "",
        AdditionalContact:r.AdditionalContact?.LastName || ""
      };
    }
  });

  return Object.values(map);
} */

/* public async getQuarterReview(context: any) {

const additionalSupport = await this.sp.web.lists
  .getByTitle("SiteAdditionalSupport")
  .items
  .select(
    "SiteID/Title",
    "AdditionalSupportRole",
    "AdditionalContact/LastName",
    "IsActive"
  )
  .expand(
    "SiteID",
    "AdditionalContact"
  )
  .filter("IsActive eq 1")
  .top(2000)();
  const additionalSupportMap: any = {};

additionalSupport.forEach((item: any) => {

  const rsrsSiteId = item.SiteID?.Title;

  if (!rsrsSiteId) {
    return;
  }

  if (!additionalSupportMap[rsrsSiteId]) {
    additionalSupportMap[rsrsSiteId] = {
      outsideC: "",
      assistantPM: "",
      technicalC: "",
      otherAdditionalS: ""
    };
  }

  const contact =
    item.AdditionalContact?.LastName || "";

  switch (item.AdditionalSupportRole) {

    case "Outside Counsel":

      additionalSupportMap[rsrsSiteId].outsideC +=
        additionalSupportMap[rsrsSiteId].outsideC
          ? `; ${contact}`
          : contact;

      break;

    case "Assistant Project Manager":

      additionalSupportMap[rsrsSiteId].assistantPM +=
        additionalSupportMap[rsrsSiteId].assistantPM
          ? `; ${contact}`
          : contact;

      break;

    case "Technical Consultant":

      additionalSupportMap[rsrsSiteId].technicalC +=
        additionalSupportMap[rsrsSiteId].technicalC
          ? `; ${contact}`
          : contact;

      break;

    default:

      additionalSupportMap[rsrsSiteId].otherAdditionalS +=
        additionalSupportMap[rsrsSiteId].otherAdditionalS
          ? `; ${contact}`
          : contact;

      break;
  }
});
const sites = await this.sp.web.lists
  .getByTitle("Site")
  .items
  .select(
    "Title",
    "RSRSSiteName",
    "Claim_Type",
    "Site_Activity",
    "Site_Type",
    "Legacy_Company",
    "Cost_Estimate_Method",
    "TotalProjectCost",
    "BETEstimatedCost",
    "Project_Stage",
    "RSRS_Update_Owner3",
    "Site_Lead",
    "Site_Support",

    "SiteLead/Title",
    "SiteSupport/Title",
    "SiteUpdateOwner/Title"
  )
  .expand(
    "SiteLead",
    "SiteSupport",
    "SiteUpdateOwner"
  )
  .filter(
    "(Site_Activity eq 'Open') or (Site_Activity eq 'Open - Monitor Only')"
  )
  .top(3000)();
  const result = sites.map((s: any) => {

  const support =
    additionalSupportMap[s.Title] || {};

  return {

    Claim_Type: s.Claim_Type || "",

    Title: s.Title || "",

    RSRSSiteName: s.RSRSSiteName || "",

    Site_Activity: s.Site_Activity || "",

    Site_Type: s.Site_Type || "",

    Legacy_Company: s.Legacy_Company || "",

    Cost_Estimate_Method:
      s.Cost_Estimate_Method || "",

    Project_Stage:
      s.Project_Stage || "",

    SiteLead:
      s.SiteLead?.[0]?.Title ||
      s.SiteLead?.Title ||
      "",

    SiteSupport:
      s.SiteSupport?.[0]?.Title ||
      s.SiteSupport?.Title ||
      "",

    SiteUpdateOwner:
      s.SiteUpdateOwner?.[0]?.Title ||
      s.SiteUpdateOwner?.Title ||
      "",

    outsideC:
      support.outsideC || "",

    assistantPM:
      support.assistantPM || "",

    technicalC:
      support.technicalC || "",

    otherAdditionalS:
      support.otherAdditionalS || "",

    TotalProjectCost:
  Number(s.TotalProjectCost || 0).toFixed(2),

BETEstimatedCost:
  Number(s.BETEstimatedCost || 0).toFixed(2),

  };
});
result.sort((a, b) =>
  (a.RSRSSiteName || "").localeCompare(
    b.RSRSSiteName || ""
  )
);

return result;
} */
public async getQuarterReview(context: any) {

  /* ==========================
     Additional Support
  ========================== */

  const additionalSupport = await this.sp.web.lists
    .getByTitle("SiteAdditionalSupport")
    .items
    .select(
      "SiteID/Title",
      "AdditionalSupportRole",
      "AdditionalContact/LastName",
      "IsActive"
    )
    .expand(
      "SiteID",
      "AdditionalContact"
    )
    .filter("IsActive eq true")
    .top(5000)();

  const additionalSupportMap: any = {};

  additionalSupport.forEach((item: any) => {

    const siteId = item.SiteID?.Title;

    if (!siteId) {
      return;
    }

    if (!additionalSupportMap[siteId]) {

      additionalSupportMap[siteId] = {
        outsideC: "",
        assistantPM: "",
        technicalC: "",
        otherAdditionalS: ""
      };
    }

    const contact =
      item.AdditionalContact?.LastName || "";

    switch (item.AdditionalSupportRole) {

      case "Outside Counsel":

        additionalSupportMap[siteId].outsideC +=
          additionalSupportMap[siteId].outsideC
            ? `; ${contact}`
            : contact;

        break;

      case "Assistant Project Manager":

        additionalSupportMap[siteId].assistantPM +=
          additionalSupportMap[siteId].assistantPM
            ? `; ${contact}`
            : contact;

        break;

      case "Technical Consultant":

        additionalSupportMap[siteId].technicalC +=
          additionalSupportMap[siteId].technicalC
            ? `; ${contact}`
            : contact;

        break;

      default:

        additionalSupportMap[siteId].otherAdditionalS +=
          additionalSupportMap[siteId].otherAdditionalS
            ? `; ${contact}`
            : contact;

        break;
    }
  });

  /* ==========================
     Get ALL Sites (Paged)
  ========================== */

  const siteUrl =
    this.context.pageContext.web.absoluteUrl;

  let allSiteItems: any[] = [];
  let nextLink: string | undefined;

  do {

    const response =
      await this.getPagedItems(
        "Site",
        {
          pageSize: 2000,
          nextLink,
          filter:
            "(Site_Activity eq 'Open') or (Site_Activity eq 'Open - Monitor Only')",
          orderBy: "RSRSSiteName",
          select: [
            "Title",
            "RSRSSiteName",
            "Claim_Type",
            "Site_Activity",
            "Site_Type",
            "Legacy_Company",
            "Cost_Estimate_Method",
            "TotalProjectCost",
            "BETEstimatedCost",
            "Project_Stage",
            "RSRS_Update_Owner3",
            "Site_Lead",
            "Site_Support",
            "SiteLead/Title",
            "SiteSupport/Title",
            "SiteUpdateOwner/Title"
          ],
          expand: [
            "SiteLead",
            "SiteSupport",
            "SiteUpdateOwner"
          ]
        },
        siteUrl
      );

    allSiteItems = [
      ...allSiteItems,
      ...response.items
    ];

    nextLink = response.nextLink;

  } while (nextLink);

  console.log(
    "Quarter Review Site Count:",
    allSiteItems.length
  );

  /* ==========================
     Build Result
  ========================== */

  const result = allSiteItems.map(
    (site: any) => {

      const support =
        additionalSupportMap[
          site.Title
        ] || {};

      let siteLead =
        site.SiteLead?.Title ||
        site.SiteLead?.[0]?.Title ||
        "";

      let siteSupport =
        site.SiteSupport?.Title ||
        site.SiteSupport?.[0]?.Title ||
        "";

      let siteUpdateOwner =
        site.SiteUpdateOwner?.Title ||
        site.SiteUpdateOwner?.[0]?.Title ||
        "";

      if (
        siteLead &&
        siteLead.indexOf(",") > 0
      ) {

        siteLead =
          siteLead.split(",")[0];
      }

      if (
        siteSupport &&
        siteSupport.indexOf(",") > 0
      ) {

        siteSupport =
          siteSupport.split(",")[0];
      }

      if (
        siteUpdateOwner &&
        siteUpdateOwner.indexOf(",") > 0
      ) {

        siteUpdateOwner =
          siteUpdateOwner.split(",")[0];
      }

      return {

        Claim_Type:
          site.Claim_Type || "",

        Title:
          site.Title || "",

        RSRSSiteName:
          site.RSRSSiteName || "",

        Site_Activity:
          site.Site_Activity || "",

        Site_Type:
          site.Site_Type || "",

        Legacy_Company:
          site.Legacy_Company || "",

        Cost_Estimate_Method:
          site.Cost_Estimate_Method || "",

        Project_Stage:
          site.Project_Stage || "",

        SiteLead:
          siteLead,

        SiteSupport:
          siteSupport,

        SiteUpdateOwner:
          siteUpdateOwner,

        outsideC:
          support.outsideC || "",

        assistantPM:
          support.assistantPM || "",

        technicalC:
          support.technicalC || "",

        otherAdditionalS:
          support.otherAdditionalS || "",

        TotalProjectCost:
          Number(
            site.TotalProjectCost || 0
          ).toFixed(2),

        BETEstimatedCost:
          Number(
            site.BETEstimatedCost || 0
          ).toFixed(2)
      };
    }
  );

  console.log(
    "Quarter Review Result Count:",
    result.length
  );

  return result;
}
public async getRFTS(context: any) {

  // Get Current Quarter
  const configItems = await this.getItems(
    "ConfigurationSetting"
  );

  const currentQuarterConfig = configItems.find(
    (item: any) =>
      item.Title === "CurrentQuarter"
  );

  const quarter =
    currentQuarterConfig?.Value0 ??
    currentQuarterConfig?.Value ??
    "";

  // Site Update Data
  const updates = await this.sp.web.lists
    .getByTitle("SiteUpdate")
    .items
    .select(
      "RecommendedChange",
      "Quarter",
      "SiteName/Title"
    )
    .expand("SiteName")
    .filter(`Quarter eq '${quarter}'`)
    .top(2000)();

  const updateMap: any = {};

  updates.forEach((u: any) => {

    const siteId = u.SiteName?.Title;

    if (!siteId) {
      return;
    }

    updateMap[siteId] =
      Number(u.RecommendedChange || 0);
  });

  // Technical Project Manager Data
  const supportItems = await this.sp.web.lists
    .getByTitle("SiteAdditionalSupport")
    .items
    .select(
      "SiteID/Title",
      "AdditionalContact/LastName",
      "AdditionalContact/Contact_Type",
      "IsActive"
    )
    .expand(
      "SiteID",
      "AdditionalContact"
    )
    .filter("IsActive eq 1")
    .top(2000)();

  const technicalManagerMap: any = {};

  supportItems.forEach((item: any) => {

    const siteId = item.SiteID?.Title;

    if (!siteId) {
      return;
    }

    if (
      item.AdditionalContact?.Contact_Type ===
      "Technical"
    ) {
      technicalManagerMap[siteId] =
        item.AdditionalContact?.LastName || "";
    }
  });

  // Site Data
  const sites = await this.sp.web.lists
    .getByTitle("Site")
    .items
    .select(
      "Title",
      "RSRSSiteName",
      "Claim_Type",
      "Site_Type",
      "Site_Activity",
      "Legacy_Company",
      "TotalProjectCost",
      "AccountProjectCode",
      "SpecialReporting",
      "SiteLead/Title",
      "SiteSupport/Title"
    )
    .expand(
      "SiteLead",
      "SiteSupport"
    )
    .top(2000)();

  const result = sites.map((site: any) => {

    const recommendedChange =
      updateMap[site.Title] || 0;

    let pgeRole = "Not Required";

    if (site.Claim_Type === "Indemnification") {

      pgeRole = "Management";
    }
    else if (
      site.SiteLead &&
      site.SiteLead.length > 0
    ) {

      pgeRole = "Project Management";
    }
    else if (
      site.SiteSupport &&
      site.SiteSupport.length > 0
    ) {

      pgeRole = "Project Support";
    }

    return {

      Title: site.Title || "",

      RSRSSiteName:
        site.RSRSSiteName || "",

      Site_Type:
        site.Site_Type || "",

      Site_Activity:
        site.Site_Activity || "",

      Legacy_Company:
        site.Legacy_Company || "",

      TechnicalProjectManager:
        technicalManagerMap[
          site.Title
        ] || "",

      PGERole: pgeRole,

      AccountProjectCode:
        site.AccountProjectCode || "",

      TotalProjectCost:
        (
          Number(
            site.TotalProjectCost || 0
          ) +
          Number(recommendedChange)
        ).toFixed(2)
    };
  });

  result.sort((a: any, b: any) =>
    (a.RSRSSiteName || "").localeCompare(
      b.RSRSSiteName || ""
    )
  );

  return result;
}

public async uploadExcelFile(file: File): Promise<number> {

  const result = await this.sp.web.lists
    .getByTitle("Outside Counsel")
    .rootFolder.files.addUsingPath(
      file.name,
      file,
      { Overwrite: true }
    );

  const item: any = await this.sp.web
    .getFileByServerRelativePath(result.ServerRelativeUrl)
    .getItem();

  return item.Id;
}
public async importOutsideCounsel(
  documentId: number,
  siteId: number,
  siteName: string,
  claimType: string
): Promise<any> {

  const endpoint =
    claimType === "Indemnification"
      ? "ImportOutsideCounselQuarterUpdatesID"
      : "ImportOutsideCounselQuarterUpdates";

  const url ="";
  // `${this.context.pageContext.web.absoluteUrl}` +  `/_layouts/15/Pfizer.SP.RSRSApplication` +    `/_vti_bin/Pfizer.SP.RSRSApplication` +    `/SPServices.svc/${endpoint}`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      docId: documentId,
      siteName: siteName,
      siteId: siteId
    })
  });

  if (!response.ok) {
    throw new Error(
      `Import service failed: ${response.status}`
    );
  }

  return await response.json();
}

public async uploadBETFile(
  file: File,
  siteId: number
): Promise<number> {

  const site = await this.getItemById(
    "Site",
    siteId
  );

  const quarter = localStorage.getItem("quarter") ?? "";

  const parts = file.name.split(".");
  const ext = parts.pop();
  const baseName = parts.join(".");

  const newFileName =
    `${site.Title}_${baseName}_${quarter}.${ext}`;

  const result = await this.sp.web.lists
    .getByTitle("BETDocs")
    .rootFolder.files.addUsingPath(
      newFileName,
      file,
      {
        Overwrite: true
      }
    );

  const item: any =
    await this.sp.web
      .getFileByServerRelativePath(
        result.ServerRelativeUrl
      )
      .getItem();

  return item.Id;
}

public async performBETCalculation(
  docId: number,
  itemId: number
): Promise<string> {

  const url ="";
   // `${this.context.pageContext.web.absoluteUrl}` +   `/_vti_bin/Pfizer.SP.RSRSApplication` +    `/SPServices.svc/PerformBETCalculation`;
  const response = await fetch(
    url,
    {
      method: "POST",
      headers: {
        "Content-Type":
          "application/json"
      },
      body: JSON.stringify({
        docId,
        itemId
      })
    }
  );

  return await response.json();
}

public async updateBETData(
  siteId: number,
  totalBETCost: number,
  fileUrl: string
): Promise<void> {

  const items =
    await this.sp.web.lists
      .getByTitle("SiteUpdate")
      .items
      .select(
        "*",
        "SiteName/Id",
        "SiteName/TotalProjectCost"
      )
      .expand("SiteName")
      .filter(
        `SiteName/Id eq ${siteId}`
      )
      .orderBy("Modified", false)
      .top(1)();

  if (!items.length) {
    return;
  }

  const latest: any = items[0];

  const totalProjectCost =
    Number(
      latest.SiteName
        ?.TotalProjectCost || 0
    );

  const difference =
    totalProjectCost -
    totalBETCost;

  const differencePercentage =
    totalProjectCost === 0
      ? 0
      : (
          difference * 100
        ) / totalProjectCost;

  await this.sp.web.lists
    .getByTitle("SiteUpdate")
    .items
    .getById(latest.Id)
    .update({
      TotalBETCost:
        totalBETCost,
      BETFileURL:
        fileUrl,
      BETDiffrence:
        difference,
      BETDiffrencePercentage:
        differencePercentage.toString(),
      BETFlag:
        differencePercentage > 5,
        BETFileuploadTime: new Date().toLocaleString()
    });
}

public async exportOutsideCounsel(
  siteUpdateId: number,
  siteName: string,
  claimType: string,
  template: string
): Promise<ArrayBuffer> {

  const url ="";
   // `${this.context.pageContext.web.absoluteUrl}` +  `/_layouts/15/Pfizer.SP.RSRSApplication` +    `/_vti_bin/Pfizer.SP.RSRSApplication` +    `/SPServices.svc/ExportExcelforQuarterUpdates`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      siteUpdateId,
      siteName,
      claimType,
      template
    })
  });

  return await response.arrayBuffer();
}



  
}
