import { useCallback, useState } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { SPFI } from "@pnp/sp";

interface IExportParams {
  siteName: string;
  claimType: string;
  templateFileName: string;
}

export const useExportQuarterUpdate = (sp: SPFI) => {
  const [loading, setLoading] = useState(false);

  const exportWorkbook = useCallback(
    async ({
      siteName,
      claimType,
      templateFileName,
    }: IExportParams): Promise<void> => {
      setLoading(true);

      try {
        const escapedSiteName = siteName.replace(
          /'/g,
          "''"
        );

        const templateFile = sp.web
          .getFolderByServerRelativePath("Template")
          .files.getByUrl(templateFileName);

        const buffer =
          await templateFile.getBuffer();

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        // const overviewSheet =
        //   workbook.getWorksheet("SITE OVERVIEW");
        const overviewSheet =
          workbook.worksheets.find(s=>s.name.toUpperCase()==="SITE OVERVIEW");
        const updateSheet =
          workbook.worksheets.find(s=>s.name.toUpperCase()==="SITE UPDATE");
        // const updateSheet =
        //   workbook.getWorksheet("SITE UPDATE");

        if (!overviewSheet) {
          throw new Error(
            "SITE OVERVIEW worksheet not found"
          );
        }

        if (!updateSheet) {
          throw new Error(
            "SITE UPDATE worksheet not found"
          );
        }

        const sites = await sp.web.lists
          .getByTitle("Site")
          .items.select(
            "*",
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
            `RSRSSiteName eq '${escapedSiteName}'`
          )();

        if (!sites.length) {
          throw new Error("Site not found");
        }

        const site = sites[0];

        const siteUpdates = await sp.web.lists
          .getByTitle("SiteUpdate")
          .items.filter(
            `SiteNameId eq ${site.Id}`
          )
          .orderBy("Modified", false)
          .top(1)();

        const update =
          siteUpdates.length > 0
            ? siteUpdates[0]
            : null;

        const totalCost = Number(
          site.TotalProjectCost || 0
        );

        const totalSpending = Number(
          site.TotalSpendingToDate || 0
        );

        const reserveBalance =
          totalCost - totalSpending;

        //
        // SITE OVERVIEW
        //

        const overviewMappings: Record<
          string,
          any
        > = {
          B2: site.Title,
          B3: site.RSRSSiteName,
          B4: site.SiteLead?.Title || "",
          B5: site.SiteSupport?.Title || "",
          B6:
            site.SiteUpdateOwner?.Title || "",
          B7: site.City,
          B8: site.State,
          B9: site.Country,
          B10: site.Site_Type,
          B11: site.Site_Activity,
          B12: site.AccountProjectCode,
          B13: site.AccountCodeBlock,
          B14: site.Legacy_Company,
          B15: site.Legacy_Company_Code,
          B16: site.Claim_Type,
          B17: site.ObligationType,
          B18: site.OperatingGroup,
          B19: site.TransactionDate,
          B20: site.DateSiteAdded,
          B21: site.OriginalCompanyConnection,
        };

        Object.entries(
          overviewMappings
        ).forEach(([cell, value]) => {
          overviewSheet.getCell(cell).value =
            value ?? "";
        });

        //
        // SITE UPDATE HEADER
        //

        updateSheet.getCell("B3").value =
          site.IsQuarterReviewDone
            ? "Yes"
            : "No";

        updateSheet.getCell("B4").value =
          totalCost;

        updateSheet.getCell("B5").value =
          totalSpending;

        updateSheet.getCell("B6").value =
          reserveBalance;

        ["B4", "B5", "B6"].forEach(
          (address) => {
            updateSheet.getCell(address).numFmt =
              "#,##0.00";
          }
        );

        if (update) {
          if (
            claimType === "Indemnification"
          ) {
            const guaranteedLiability =
              Number(
                update.RecommendedChange || 0
              ) +
              totalCost -
              totalSpending;

            const indemnificationFields = {
              E7:
                update.MaxPotentialPayments,
              E8: guaranteedLiability,
              E9:
                update.RecourseProvisions,
              B10:
                update.RecourseProvisionsNA
                  ? "Yes"
                  : "No",
              E11:
                update.AssetsHeldCollateral,
              B12:
                update.AssetsHeldCollateralNA
                  ? "Yes"
                  : "No",
              B13: update.Term,
              B14: update.Expiration,
              B15:
                update.ExpirationNA
                  ? "Yes"
                  : "No",
              B16:
                update.HowIndemnificationArose,
              B17:
                update.EventsReqIndemnifierGuarantor,
              E18:
                update.AmountPaymentsIndemnification,
              E19:
                update.CumulativeAmountIndemnification,
              B20:
                update
                  .LikelihoodofthoseEventsOccurring,
              B21: update.Remarks,
              B22:
                update.RecommendedClosure
                  ? "Yes"
                  : "No",
              E23:
                update.RecommendedChange,
              B24:
                update.BasisForChange,
              B25:
                update.CompletedReview
                  ? "Yes"
                  : "No",
            };

            Object.entries(
              indemnificationFields
            ).forEach(([cell, value]) => {
              updateSheet.getCell(cell).value =
                value ?? "";
            });
          } else {
            const standardFields = {
              B7:
                update.KeySiteMilestones,
              B8:
                update.CurrentSiteStatus,
              E9:
                update.PercentageContributionLiability,
              E10:
                update.RecommendedChange,
              B11:
                update.RecommendedClosure
                  ? "Yes"
                  : "No",
              B12:
                update.BasisForChange,
              B13:
                update.CompletedReview
                  ? "Yes"
                  : "No",
            };

            Object.entries(
              standardFields
            ).forEach(([cell, value]) => {
              updateSheet.getCell(cell).value =
                value ?? "";
            });
          }
        }

        //
        // UNLOCK EDITABLE CELLS
        //

        const editableCells =
          claimType === "Indemnification"
            ? [
                "E7",
                "E9",
                "B10",
                "E11",
                "B12",
                "B13",
                "B14",
                "B15",
                "B16",
                "B17",
                "E18",
                "E19",
                "B20",
                "B21",
                "B22",
                "E23",
                "B24",
                "B25",
              ]
            : [
                "B7",
                "B8",
                "E9",
                "E10",
                "B11",
                "B12",
                "B13",
              ];

        editableCells.forEach((cell) => {
          updateSheet.getCell(
            cell
          ).protection = {
            locked: false,
          };
        });

        await overviewSheet.protect(
          "PfizerRSRS",
          {
            selectLockedCells: true,
            selectUnlockedCells: true,
          }
        );

        await updateSheet.protect(
          "PfizerRSRS",
          {
            selectLockedCells: true,
            selectUnlockedCells: true,
          }
        );

        const output =
          await workbook.xlsx.writeBuffer();
          const today = new Date();

const formattedDate =
  `${String(today.getDate()).padStart(2, "0")}-` +
  `${String(today.getMonth() + 1).padStart(2, "0")}-` +
  `${today.getFullYear()}`;

saveAs(
  new Blob([output]),
  `${siteName} ${formattedDate}.xlsx`
);

        /* saveAs(
          new Blob([output]),
          `${siteName}.xlsx`
        ); */
      } catch (error) {
        console.error(
          "Quarter Update Export Error",
          error
        );
        throw error;
      } finally {
        setLoading(false);
      }
    },
    [sp]
  );

  return {
    exportWorkbook,
    loading,
  };
};