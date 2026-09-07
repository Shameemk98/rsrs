import ExcelJS from "exceljs";
import { SPFI } from "@pnp/sp";

export default class OutsideCounselImportService {

  constructor(private readonly sp: SPFI) {}

  public async ImportOutsideCounselQuarterUpdates(
    file: File,
    siteId: number,
    siteName: string
  ): Promise<void> {

    const workbook = await this.loadWorkbook(file);

    if (!this.validateSiteName(workbook, siteName)) {
      throw new Error(
        "RSRS Site Name DO NOT match in the file."
      );
    }

    /* const worksheet =
      workbook.getWorksheet("SITE UPDATE"); */
      const worksheet =
          workbook.worksheets.find(s=>s.name.toUpperCase()==="SITE UPDATE");
      

    if (!worksheet) {
      throw new Error(
        "SITE UPDATE worksheet not found."
      );
    }

    const reserve =
      this.getNumber(worksheet, "B6");

    const result = {

      KeySiteMilestones:
        this.getText(worksheet, "B7"),

      CurrentSiteStatus:
        this.getText(worksheet, "B8"),

      PercentageContributionLiability:
        this.getNumber(worksheet, "E9"),

      RecommendedChange:
        this.getNumber(worksheet, "E10"),

      RecommendedClosure:
        this.yesNoToFlag(
          this.getText(worksheet, "B11")
        ),

      BasisForChange:
        this.getText(worksheet, "B12"),

      CompletedReview:
        this.yesNoToFlag(
          this.getText(worksheet, "B13")
        )
    };

    if (result.RecommendedClosure === "1") {
      result.RecommendedChange = reserve * -1;
    }

    await this.saveNormalClaim(
      siteId,
      result
    );
  }

  public async ImportOutsideCounselQuarterUpdatesID(
    file: File,
    siteId: number,
    siteName: string
  ): Promise<void> {

    const workbook = await this.loadWorkbook(file);

    if (!this.validateSiteName(workbook, siteName)) {
      throw new Error(
        "RSRS Site Name DO NOT match in the file."
      );
    }

   // const worksheet = workbook.getWorksheet("SITE UPDATE");
    const worksheet =workbook.worksheets.find(s=>s.name.toUpperCase()==="SITE UPDATE");

    if (!worksheet) {
      throw new Error(
        "SITE UPDATE worksheet not found."
      );
    }

    const reserve =
      this.getNumber(worksheet, "B6");

    const result = {

      MaxPotentialPayments:
        this.getNumber(worksheet, "E7"),

      RecourseProvisions:
        this.getNumber(worksheet, "E9"),

      RecourseProvisionsNA:
        this.yesNoToFlag(
          this.getText(worksheet, "B10")
        ),

      AssetsHeldCollateral:
        this.getNumber(worksheet, "E11"),

      AssetsHeldCollateralNA:
        this.yesNoToFlag(
          this.getText(worksheet, "B12")
        ),

      Term:
        this.getText(worksheet, "B13"),

      Expiration:
        this.getText(worksheet, "B14"),

      ExpirationNA:
        this.yesNoToFlag(
          this.getText(worksheet, "B15")
        ),

      HowIndemnificationArose:
        this.getText(worksheet, "B16"),

      EventsReqIndemnifierGuarantor:
        this.getText(worksheet, "B17"),

      AmountPaymentsIndemnification:
        this.getNumber(worksheet, "E18"),

      CumulativeAmountIndemnification:
        this.getNumber(worksheet, "E19"),

      LikelihoodofthoseEventsOccurring:
        this.getText(worksheet, "B20"),

      Remarks:
        this.getText(worksheet, "B21"),

      RecommendedClosure:
        this.yesNoToFlag(
          this.getText(worksheet, "B22")
        ),

      RecommendedChange:
        this.getNumber(worksheet, "E23"),

      BasisForChange:
        this.getText(worksheet, "B24"),

      CompletedReview:
        this.yesNoToFlag(
          this.getText(worksheet, "B25")
        )
    };

    if (result.RecommendedClosure === "1") {
      result.RecommendedChange = reserve * -1;
    }

    await this.saveIndemnification(
      siteId,
      result
    );
  }

    private async loadWorkbook(
    file: File
  ): Promise<ExcelJS.Workbook> {

    const workbook =
      new ExcelJS.Workbook();

    const buffer =
      await file.arrayBuffer();
      console.log("Workbook loaded", workbook);

console.log("Worksheets", workbook.worksheets);
    const int8 = new  Uint8Array(buffer)
    console.log(int8.length)
    await workbook.xlsx.load(buffer );

    return workbook;
  } 
 
  private validateSiteName(
    workbook: ExcelJS.Workbook,
    siteName: string
  ): boolean {

    /* const overview = workbook.getWorksheet("SITE OVERVIEW"); */
      const overview = workbook.worksheets.find(s=>s.name.toUpperCase()==="SITE OVERVIEW");

    if (!overview) {
      throw new Error(
        "SITE OVERVIEW worksheet not found."
      );
    }

    const excelSiteName =
      this.getText(overview, "B3");

    return (
      excelSiteName.trim().toUpperCase() ===
      siteName.trim().toUpperCase()
    );
  }

  private async getCurrentQuarter(): Promise<string> {

    const result = await this.sp.web.lists
      .getByTitle("ConfigurationSetting")
      .items
      .filter("Title eq 'CurrentQuarter'")
      .top(1)();

    return (
      result[0]?.Value0 ??
      result[0]?.Value ??
      ""
    );
  }

  private async saveNormalClaim(
    siteId: number,
    model: any
  ): Promise<void> {

    const quarter =
      await this.getCurrentQuarter();

    const existing =
      await this.sp.web.lists
        .getByTitle("SiteUpdate")
        .items
        .filter(
          `SiteNameId eq ${siteId} and Quarter eq '${quarter}'`).top(1)();
        console.log("existing:", existing);
    const previousRecord = await this.getPreviousQuarterRecord(
          siteId,
          quarter
        );
        console.log("existpreviousRecording:", previousRecord);
  
    const approvedFields =
        this.buildApprovedFieldPayload(
          previousRecord
        );
        console.log("approvedFields:", approvedFields);
    const payload = {
      KeySiteMilestones: model.KeySiteMilestones,

      CurrentSiteStatus:    model.CurrentSiteStatus,

      PercentageContributionLiability:  model.PercentageContributionLiability,

      RecommendedChange: model.RecommendedChange,

      RecommendedClosure:  model.RecommendedClosure === "1",

      BasisForChange:  model.BasisForChange,

      CompletedReview:    model.CompletedReview === "1",
        ...approvedFields
    };
console.log("payload:", payload);
    if (existing.length > 0) {

      await this.sp.web.lists
        .getByTitle("SiteUpdate")
        .items
        .getById(existing[0].Id)
        .update(payload);

    } else {

      await this.sp.web.lists
        .getByTitle("SiteUpdate")
        .items
        .add({
          ...payload,
          SiteNameId: siteId,
          Quarter: quarter
        });
    }

    await this.updateSite(
      siteId,
      model.CompletedReview
    );
  }

  private async saveIndemnification(
    siteId: number,
    model: any
  ): Promise<void> {

    const quarter =
      await this.getCurrentQuarter();
      

    const existing =
      await this.sp.web.lists
        .getByTitle("SiteUpdate")
        .items
        .filter(
          `SiteNameId eq ${siteId} and Quarter eq '${quarter}'`
          ).top(1)();
          console.log(existing,"existingexisting");
          
    const previousRecord =
        await this.getPreviousQuarterRecord(
          siteId,
          quarter
        );

        console.log(previousRecord,"previousRecord");
  
    const approvedFields =
        this.buildApprovedFieldPayload(
          previousRecord
        );
        console.log(approvedFields,"approvedFields");
        
    const payload = {
      MaxPotentialPayments:
        model.MaxPotentialPayments,

      RecourseProvisions:
        model.RecourseProvisions,

      RecourseProvisionsNA:
        model.RecourseProvisionsNA === "1",

      AssetsHeldCollateral:
        model.AssetsHeldCollateral,

      AssetsHeldCollateralNA:
        model.AssetsHeldCollateralNA === "1",

      Term:
        model.Term,

      Expiration:
        model.Expiration,

      ExpirationNA:
        model.ExpirationNA === "1",

      HowIndemnificationArose:
        model.HowIndemnificationArose,

      EventsReqIndemnifierGuarantor:
        model.EventsReqIndemnifierGuarantor,

      AmountPaymentsIndemnification:
        model.AmountPaymentsIndemnification,

      CumulativeAmountIndemnification:
        model.CumulativeAmountIndemnification,

      LikelihoodofthoseEventsOccurring:
        model.LikelihoodofthoseEventsOccurring,

      Remarks:
        model.Remarks,

      RecommendedClosure:
        model.RecommendedClosure === "1",

      RecommendedChange:
        model.RecommendedChange,

      BasisForChange:
        model.BasisForChange,

      CompletedReview:
        model.CompletedReview === "1"
        ,
        ...approvedFields
    };

    console.log(payload,"payload");

   
    if (existing.length > 0) {

      await this.sp.web.lists
        .getByTitle("SiteUpdate")
        .items
        .getById(existing[0].Id)
        .update(payload);

    } else {

      await this.sp.web.lists
      .getByTitle("SiteUpdate")
      .items
      .add({
        ...payload,
        SiteNameId: siteId,
        Quarter: quarter
      });
    }

    await this.updateSite(
      siteId,
      model.CompletedReview
    );
  }

  private async updateSite(
    siteId: number,
    completedReview: string
  ): Promise<void> {

    await this.sp.web.lists
      .getByTitle("Site")
      .items
      .getById(siteId)
      .update({
        IsQuarterReviewDone: completedReview === "1" ? "true" : "false",
        DateLastUpdated: new Date()
      });
  }
  private async getPreviousQuarterRecord(
    siteId: number,
    currentQuarter: string
  ): Promise<any | null> {
  
    const items = await this.sp.web.lists
      .getByTitle("SiteUpdate")
      .items
      .filter(
        `SiteNameId eq ${siteId} and Quarter ne '${currentQuarter}'`
      )
      .orderBy("Created", false)
      .top(1)();
  
    return items.length > 0
      ? items[0]
      : null;
  }
  private buildApprovedFieldPayload(
    previousRecord: any
  ): any {
  
    if (!previousRecord) {
      return {};
    }
  
    return {
  
      Key_x0020_Site_x0020_Milestones_:
        previousRecord.Key_x0020_Site_x0020_Milestones_,
  
      Current_x0020_Site_x0020_Status_:
        previousRecord.Current_x0020_Site_x0020_Status_,
  
      Percentage_x0020_Contribution_x0:
        previousRecord.Percentage_x0020_Contribution_x0,
  
      Recourse_x0020_Provisions_x0020_:
        previousRecord.Recourse_x0020_Provisions_x0020_,
      Recourse_x0020_Provisions_x0020_0:
        previousRecord.Recourse_x0020_Provisions_x0020_0,
  
      Assets_x0020_Held_x0020_as_x0020:
        previousRecord.Assets_x0020_Held_x0020_as_x0020,
  
      Assets_x0020_Held_x0020_as_x00200:
        previousRecord.Assets_x0020_Held_x0020_as_x00200,
  
      Term_x0020__x0028_Approved_x0029:
        previousRecord.Term_x0020__x0028_Approved_x0029,
  
      Expiration_x0020__x0028_Approved:
        previousRecord.Expiration_x0020__x0028_Approved,
  
      Expiration_x0020_Not_x0020_Appli:
        previousRecord.Expiration_x0020_Not_x0020_Appli,
  
      How_x0020_Indemnification_x0020_:
        previousRecord.How_x0020_Indemnification_x0020_,
  
      Events_x0020_that_x0020_Require_:
        previousRecord.Events_x0020_that_x0020_Require_,
  
      Amount_x0020_of_x0020_any_x0020_:
        previousRecord.Amount_x0020_of_x0020_any_x0020_,
  
      Cumulative_x0020_Amount_x0020_of:
        previousRecord.Cumulative_x0020_Amount_x0020_of,
  
      Likelihood_x0020_of_x0020_those_:
        previousRecord.Likelihood_x0020_of_x0020_those_,
  
      Remarks_x0020__x0028_Approved_x0:
        previousRecord.Remarks_x0020__x0028_Approved_x0,
  
      Date_of_Last_Data_Update:
        previousRecord.Date_of_Last_Data_Update,
  
      Maximum_x0020_Potential_x0020_Pa:
        previousRecord.Maximum_x0020_Potential_x0020_Pa
    };
  }

  private getText(
    worksheet: ExcelJS.Worksheet,
    address: string
  ): string {
  
    const value =
      worksheet.getCell(address).value;
  
    if (!value) {
      return "";
    }
  
    if (
      typeof value === "object" &&
      "richText" in value
    ) {
      return value.richText
        .map((r: any) => r.text)
        .join("")
        .trim();
    }
  
    return String(value).trim();
  }

  private yesNoToFlag(
    value: string
  ): string {

    return value?.trim().toUpperCase() === "YES"
      ? "1"
      : "0";
  }
  private getNumber(
    worksheet: ExcelJS.Worksheet,
    address: string
  ): number {
  
    const value =
      worksheet.getCell(address).value;
  
    if (
      value &&
      typeof value === "object" &&
      "result" in value
    ) {
      return Number(
        (value as any).result
      ) || 0;
    }
  
    return Number(value) || 0;
  }
}