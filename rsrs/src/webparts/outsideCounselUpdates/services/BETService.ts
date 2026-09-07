import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { BetCalculationService, BETValidationStatus } from "./BetCalculationService";

export interface IBETCalculationResult {
  totalBETCost: number;
  fileUrl: string;
}

export class BETValidationError extends Error {
  constructor(
    public readonly type: BETValidationStatus,
    message: string
  ) {
    super(message);
    this.name = "BETValidationError";
  }
}

export class BETService {
  private readonly sp: SPFI;

  constructor(context: any) {
    this.sp = spfi().using(SPFx(context));
  }
 
  public async uploadBETFile(
    file: File,
    fileName: string
  ): Promise<{
    itemId: number;
    serverRelativeUrl: string;
  }> {
    try {
      const webInfo = await this.sp.web();
      const folderPath= `${webInfo.ServerRelativeUrl}/BETDocs`
      const fileInfo = await this.sp.web
        .getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(
          fileName,
          file,
          {
            Overwrite: true
          }
        );

      const fileRef =
        this.sp.web.getFileByServerRelativePath(
          fileInfo.ServerRelativeUrl
        );

      const item = await fileRef.getItem<{ Id: number }>();

      if (!item?.Id) {
        throw new BETValidationError(
          BETValidationStatus.UploadFailed,
          `Unable to retrieve list item for uploaded file ${fileName}`
        );
      }

      return {
        itemId: item.Id,
        serverRelativeUrl: fileInfo.ServerRelativeUrl
      };
    } catch (error) {
      throw new BETValidationError(
        BETValidationStatus.UploadFailed,
        "Unable to upload BET file."
      );
    }
  }

  public async performBETCalculation(
    serverRelativeUrl: string,
    siteId: number
  ): Promise<any>{
    try{
      const service = new BetCalculationService(this.sp);
      console.log("hii from betservc");
      
      const responseData =await service.performBETCalculation(serverRelativeUrl,siteId);
       console.log(responseData,"responseData");
      if (!responseData.success) {

  const normalized =
    (responseData.message ?? "")
      .trim()
      .toLowerCase();

  console.log(
    "normalized:",
    normalized
  );

  if (normalized === "sitename not matched") {

    throw new BETValidationError(
      BETValidationStatus.SiteNameMismatch,
      "RSRS Site Name does not match the site selected in the application. Please upload a corrected BET file."
    );
  }

  if (normalized === "invalid file") {

    throw new BETValidationError(
      BETValidationStatus.InvalidFile,
      "File validation failed. Please ensure that you are using the approved BET template."
    );
  }

  if (normalized === "invalid file content") {

    throw new BETValidationError(
      BETValidationStatus.InvalidContent,
      "BET file validation failed. Verify RSRS Site Name, required worksheets and ensure there are no Excel calculation errors."
    );
  }

  throw new BETValidationError(
    BETValidationStatus.CalculationFailed,
    responseData.message ??
      "BET calculation failed."
  );
}

      else{
        const totalBETCost =responseData.total??0;
        
        if (Number.isNaN(totalBETCost)) {
          throw new BETValidationError(
            BETValidationStatus.CalculationFailed,
            "BET calculation result is invalid."
          );
        }
        return {
          totalBETCost,
          fileUrl: responseData.approvedFileUrl??""
        };
      }      
    }
    catch (error: any) {

  await this.logError(
    "BET Calculation",
    error?.message ?? "",
    error?.stack
  );

  if (
    error instanceof BETValidationError
  ) {
    throw error;
  }

  throw new BETValidationError(
    BETValidationStatus.CalculationFailed,
    error?.message ??
      "BET calculation failed."
  );
}
    
  }
  public buildBETFileName(
    originalFileName: string,
    rsrsSiteId: string,
    quarter: string
  ): string {

    const ext = originalFileName
      .split(".")
      .pop();

    const fileName =
      originalFileName.replace(
        /\.xlsm$/i,
        ""
      );

    const safeQuarter =
      quarter.replace(/\s+/g, "_");

    return `${rsrsSiteId}_${fileName}_${safeQuarter}.${ext}`;
  }
  private async logError(
    module: string,
    message: string,
    stackTrace?: string
  ): Promise<void> {
  
    await this.sp.web.lists
      .getByTitle("ErrorLogs")
      .items.add({
        Title: message,
        Module: module,
        StackTrace: stackTrace ?? ""
      });
  }
}

export interface IBETCalculationResult {
  totalBETCost: number;
  fileUrl: string;
}
