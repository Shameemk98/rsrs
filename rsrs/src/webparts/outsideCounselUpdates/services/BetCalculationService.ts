import { SPFI } from "@pnp/sp";
import * as XLSX from "xlsx";
import { LISTS } from "../../../common/constants";

export enum BETValidationStatus {
  Success = "Success",
  InvalidFile = "Invalid File",
  SiteNameMismatch = "Sitename not matched",
  InvalidContent = "Invalid file Content",
  CalculationFailed = "CalculationFailed",
  UploadFailed = "UploadFailed"
}

export interface IBETResult {
  success: boolean;
  total?: number;
  approvedFileUrl?: string;
  message?: string;
}

export class BetCalculationService {
  private static readonly SITE_NAME_CELL = "A4";
  private static readonly SITE_VALUE_CELL = "C4";
  private static readonly ADDRESS_CELL = "A5";
  private static readonly TOTAL_CELL = "L28";
  private static readonly APPROVED_LIBRARY = "BETDocsApproved";
  constructor(private readonly sp: SPFI) {}

  public async performBETCalculation(
    fileServerRelativeUrl: string,
    siteItemId: number
  ): Promise<IBETResult> {

    const siteName = await this.getSiteName(siteItemId);
      console.log("hii from bet.calservc");

    const workbook = await this.loadWorkbook(
      fileServerRelativeUrl
    );

    const worksheet =this.getWorksheet(workbook);
    if (!worksheet) {
    return {
        success: false,
        message: BETValidationStatus.InvalidFile
    };
    }

    const validationResult = this.validateWorksheet(
      worksheet,
      siteName
    );

    if (validationResult !== BETValidationStatus.Success) {
      return {
        success: false,
        message: validationResult
      };
    }

    const total = this.getTotalCost(worksheet);

    if (total === null) {
      return {
        success: false,
        message: BETValidationStatus.InvalidContent
      };
    }

    const destinationUrl =
      await this.moveToApprovedLibrary(
        fileServerRelativeUrl
      );

    await this.updateSiteItem(
      siteItemId,
      total,
      destinationUrl
    );

    return {
      success: true,
      total,
      approvedFileUrl: destinationUrl
    };
  }

  private async getSiteName(
    siteItemId: number
  ): Promise<string> {

    const item =
      await this.sp.web.lists
        .getByTitle(LISTS.SITE)
        .items
        .getById(siteItemId)();

    return String(
      item?.RSRSSiteName ?? ""
    ).trim();
  }

  private async loadWorkbook(
    fileServerRelativeUrl: string
  ): Promise<XLSX.WorkBook> {

    const fileBuffer =
      await this.sp.web
        .getFileByServerRelativePath(
          fileServerRelativeUrl
        )
        .getBuffer();

    return XLSX.read(fileBuffer, {
      type: "buffer"
    });
  }

  private getWorksheet(
    workbook: XLSX.WorkBook
  ): XLSX.WorkSheet |null{

    if (
      !workbook.SheetNames ||
      workbook.SheetNames.length === 0
    ) {
      return null
    }

    return workbook.Sheets[
      workbook.SheetNames[0]
    ];
  }

  private validateWorksheet(
    sheet: XLSX.WorkSheet,
    siteName: string
  ): BETValidationStatus {

    const a4 = this.getCellValue(
      sheet,
      BetCalculationService.SITE_NAME_CELL
    );

    const a5 = this.getCellValue(
      sheet,
      BetCalculationService.ADDRESS_CELL
    );

    const c4 = this.getCellValue(
      sheet,
      BetCalculationService.SITE_VALUE_CELL
    );

    if (a4 !== "site name:") {
      return BETValidationStatus.InvalidFile;
    }

    if (a5 !== "address:") {
      return BETValidationStatus.InvalidFile;
    }

    if (
      c4 !== siteName.trim().toLowerCase()
    ) {
      return BETValidationStatus.SiteNameMismatch;
    }

    return BETValidationStatus.Success;
  }

  private getTotalCost(
    sheet: XLSX.WorkSheet
  ): number | null {

    const cell =
      sheet[
        BetCalculationService.TOTAL_CELL
      ];

    if (!cell) {
      return null;
    }

    const total = Number(cell.v);

    return Number.isNaN(total)
      ? null
      : total;
  }

  private async moveToApprovedLibrary(
    fileServerRelativeUrl: string
  ): Promise<string> {

    const webInfo = await this.sp.web();

    const fileName =
      fileServerRelativeUrl
        .split("/")
        .pop();

    const destinationUrl =
      `${webInfo.ServerRelativeUrl}/${BetCalculationService.APPROVED_LIBRARY}/${fileName}`;

    await this.sp.web
      .getFileByServerRelativePath(
        fileServerRelativeUrl
      )
      .moveByPath(
        destinationUrl,
        true
      );

    return destinationUrl;
  }

  private async updateSiteItem(
    siteItemId: number,
    total: number,
    fileUrl: string
  ): Promise<void> {

    await this.sp.web.lists
      .getByTitle(LISTS.SITE)
      .items
      .getById(siteItemId)
      .update({
        BETEstimatedCost: total.toString(),
        BETFileUrl: fileUrl
      });
  }

  private getCellValue(
    sheet: XLSX.WorkSheet,
    address: string
  ): string {

    return String(
      sheet[address]?.v ?? ""
    )
      .trim()
      .toLowerCase();
  }
}