import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/items';
import '@pnp/sp/lists';
import "@pnp/sp/folders";
import '@pnp/sp/attachments';
import '@pnp/sp/site-users/web';
import "@pnp/sp/files";
import { useExportQuarterUpdate } from './ExportQuarterUpdate';
import { getSplitButtonClassNames } from '@fluentui/react';
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
  
  public  getSP=()=>{
    return this.sp

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
        'SiteUpdateOwner/Id','SiteUpdateOwner/Title','SiteUpdateOwner/EMail')
      //.select('*', 'SiteLead/Id','SiteLead/Title','SiteSupport/Id','SiteSupport/Title','SiteUpdateOwner/Id','SiteUpdateOwner/Title')
      .expand('SiteLead','SiteSupport','SiteUpdateOwner')();
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

public async getLegalExport(context: any) {

  const support = await this.sp.web.lists
    .getByTitle("SiteAdditionalSupport")
    .items.select(
      "SiteID/Title",
      "SiteID/RSRSSiteName",
      "SiteID/City",
      "SiteID/State",
      "SiteID/Country",
      "SiteID/Site_Type",
      "SiteID/Site_Activity",
      "SiteID/Legacy_Company",
      "SiteID/Claim_Type",
      "SiteID/TotalProjectCost",
      "SiteID/TotalSpendingToDate",
      "SiteID/AccountProjectCode",
      "AdditionalSupportRole",
      "AdditionalContact/LastName"
    )
    .expand("SiteID", "AdditionalContact")();

  const sites = await this.sp.web.lists
    .getByTitle("Site")
    .items.select(
      "RSRSSiteName",
      "SiteLead/Title",
      "SiteSupport/Title",
      "SiteUpdateOwner/Title"
    )
    .expand(
      "SiteLead",
      "SiteSupport",
      "SiteUpdateOwner"
    )();

  const map: any = {};

  support.forEach((r: any) => {
    const id = r.SiteID?.Title;

    if (!id) return;

    if (!map[id]) {
      const s = sites.find(
        (x: any) =>
          x.RSRSSiteName === r.SiteID.RSRSSiteName
      );

      map[id] = {
        Title: id,
        RSRSSiteName: r.SiteID.RSRSSiteName,
        City: r.SiteID.City,
        State: r.SiteID.State,
        Country: r.SiteID.Country,
        Site_Type: r.SiteID.Site_Type,
        Site_Activity: r.SiteID.Site_Activity,
        Legacy_Company: r.SiteID.Legacy_Company,
        SiteLead: s?.SiteLead?.Title || "",
        SiteSupport: s?.SiteSupport?.Title || "",
        SiteUpdateOwner: s?.SiteUpdateOwner?.Title || "",
        OutsideCounsel: "",
        AccountProjectCode:
          r.SiteID.AccountProjectCode,
        Claim_Type: r.SiteID.Claim_Type,
        TotalProjectCost:
          r.SiteID.TotalProjectCost || 0,
        TotalSpendingToDate:
          r.SiteID.TotalSpendingToDate || 0,
        ReserveBalance: 0
      };
    }

    if (
      r.AdditionalSupportRole ===
      "Outside Counsel"
    ) {
      map[id].OutsideCounsel +=
        (map[id].OutsideCounsel ? "; " : "") +
        (r.AdditionalContact?.LastName || "");
    }
  });

  Object.values(map).forEach((m: any) => {
    m.ReserveBalance =
      m.TotalProjectCost -
      m.TotalSpendingToDate;
  });

  return Object.values(map);
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

public async getNegativeReserve(context: any) {

  const data = await this.sp.web.lists
    .getByTitle("Site")
    .items
    .filter("Site_Activity eq 'Closed'")();

  return data
    .map((r: any) => {
      const reserve =
        (r.TotalProjectCost || 0) -
        (r.TotalSpendingToDate || 0);

      return {
        Title: r.Title,
        RSRSSiteName: r.RSRSSiteName,
        Reserve: reserve,
        Site_Activity: r.Site_Activity
      };
    })
    .filter((x: any) => x.Reserve < 0);
}

public async getQuarterReview(context: any) {

  const data = await this.sp.web.lists
    .getByTitle("SiteAdditionalSupport")
    .items
    .select(
      "SiteID/Title",
      "SiteID/RSRSSiteName",
      "SiteID/Claim_Type",
      "SiteID/Site_Activity",
      "SiteID/Site_Type",
      "SiteID/Legacy_Company"
    )
    .expand("SiteID")();

  const map: any = {};

  data.forEach((r: any) => {
    const id = r.SiteID?.Title;

    if (!id) return;

    if (!map[id]) {
      map[id] = {
        Title: id,
        RSRSSiteName: r.SiteID.RSRSSiteName,
        Claim_Type: r.SiteID.Claim_Type,
        Site_Activity: r.SiteID.Site_Activity,
        Site_Type: r.SiteID.Site_Type,
        Legacy_Company: r.SiteID.Legacy_Company
      };
    }
  });

  return Object.values(map);
}

public async getRFTS(context: any) {

  const sites = await this.sp.web.lists
    .getByTitle("Site")
    .items
    .select(
      "Title",
      "RSRSSiteName",
      "Site_Activity",
      "Legacy_Company",
      "TotalProjectCost",
      "AccountProjectCode"
    )();

  return sites.map((r: any) => ({
    Title: r.Title,
    RSRSSiteName: r.RSRSSiteName,
    Site_Activity: r.Site_Activity,
    Company: r.Legacy_Company,
    AccountProjectCode: r.AccountProjectCode,
    Cost: r.TotalProjectCost
  }));
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
): Promise<void> {
  const {loading,exportWorkbook}=useExportQuarterUpdate(this.sp)
  await exportWorkbook({
    siteName: siteName,
    claimType: claimType,
    templateFileName:
     claimType ===
      "Indemnification"
        ? "Site Update Template Indemnification.xlsx"
        : "Site Update Template.xlsx",
  });
}



  
}
