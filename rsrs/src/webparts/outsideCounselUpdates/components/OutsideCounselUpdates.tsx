import * as React from 'react';
import styles from './OutsideCounselUpdates.module.scss';
import type { IOutsideCounselUpdatesProps } from './IOutsideCounselUpdatesProps';
import RsrsHeader from '../../rsrsHeader/components/RsrsHeader';
import {  Home32Regular } from "@fluentui/react-icons";
import {  DefaultButton,  Dropdown,  Link,  PrimaryButton,  SpinnerSize,  Dialog,  DialogType,  DialogFooter,  Spinner,  MessageBarType} from '@fluentui/react';
import {  CheckmarkCircle24Regular,  ErrorCircle24Regular} from "@fluentui/react-icons";
/* import { SpService } from '../../../services/spService'; */
import { useState } from "react";
import { SpService } from '../services/exportService';
import { useExportQuarterUpdate } from '../services/ExportQuarterUpdate';
import OutsideCounselImportService from "../services/OutsideCounselImportService";
import { BETService, BETValidationError } from "../services/BETService";


export default function (props: IOutsideCounselUpdatesProps): React.ReactElement<IOutsideCounselUpdatesProps> {
  const [action, setAction] = useState<'export' | 'import'>('export');
  const [sitesOptions, setSitesOptions] = useState<any[]>([]);
  const [selectedSiteId, setSelectedSiteId] = React.useState<string | number>();
  const [loading, setLoading] = useState(false);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [betFile, setBetFile] = useState<File | null>(null);
  const [showDialog, setShowDialog] =  useState(false);
const [message, setMessage] =  useState("");
  const excelFileRef = React.useRef<HTMLInputElement>(null);
  const betFileRef = React.useRef<HTMLInputElement>(null);


const [isErrorDialog, setIsErrorDialog] =
  useState(false);

  const spService = new SpService(props.context);
  const sp = spService.getSP();
  const { exportWorkbook } = useExportQuarterUpdate(sp);
  const betService = new BETService(props.context);

  const importService = new OutsideCounselImportService(spService.getSP());

  const loadSites = async (): Promise<void> => {
    try {
      const sitesData = await spService.getItems(
        "Site",
        ["Id", "Title", "RSRSSiteName", "Claim_Type", "BETFileUrl"],
        undefined,
        {
          field: "RSRSSiteName",
          ascending: true
        }
      );

      setSitesOptions(sitesData);

    } catch (error) {
      console.error(error);
    }
  };

const showDialogMessage = (
  text: string,
  isError: boolean
): void => {

  setMessage(text);
  setIsErrorDialog(isError);
  setShowDialog(true);

};



const handleImport = async (): Promise<void> => {
  try {

    // =============================
    // VALIDATIONS
    // =============================

    if (!selectedSiteId) {
      showDialogMessage(
        "Please select a site",
        true
      );
      return;
    }

    if (!excelFile) {
      showDialogMessage(
        "Please select a file to upload",
        true
      );
      return;
    }

    const excelExt = excelFile.name
      .split(".")
      .pop()
      ?.toLowerCase();

    if (
      excelExt !== "xls" &&
      excelExt !== "xlsx"
    ) {
      showDialogMessage(
  "Only .xls or .xlsx files are allowed",
  true
);
      return;
    }

    if (betFile) {

      const betExt = betFile.name
        .split(".")
        .pop()
        ?.toLowerCase();

      if (betExt !== "xlsm") {
        showDialogMessage(
  "Only .xlsm format is allowed for BET file",
  true
);
        return;
      }
    }

    const selectedSite =
      sitesOptions.find(
        site =>
          Number(site.Id) ===
          Number(selectedSiteId)
      );

    if (!selectedSite) {
      showDialogMessage(
  "Selected site not found",
  true
);
      return;
    }

    const claimType =
      selectedSite.Claim_Type;

    // =============================
    // STEP 1 : IMPORT EXCEL
    // =============================

    setLoading(true);

    try {

      if (
        claimType ===
        "Indemnification"
      ) {

        await importService
          .ImportOutsideCounselQuarterUpdatesID(
            excelFile,
            Number(selectedSiteId),
            selectedSite.RSRSSiteName
          );

      } else {

        await importService
          .ImportOutsideCounselQuarterUpdates(
            excelFile,
            Number(selectedSiteId),
            selectedSite.RSRSSiteName
          );
      }

      setLoading(false);

      showDialogMessage(
  "Site update completed successfully.",
  false
);

    } catch (error: any) {

      setLoading(false);

     showDialogMessage(
  error?.message ??
  "Site updates file validation failed. Please ensure that you are using the correct exported file.",
  true
);

      return;
    }

    // =============================
    // STEP 2 : BET PROCESSING
    // =============================

    if (betFile) {

      try {

        setLoading(true);

        const configItems =
          await spService.getItems(
            "ConfigurationSetting"
          );

        const quarter =
          configItems.find(
            (item: any) =>
              item.Title ===
              "CurrentQuarter"
          )?.Value0 ??
          configItems.find(
            (item: any) =>
              item.Title ===
              "CurrentQuarter"
          )?.Value ??
          "";

        const betFileName =
          betService.buildBETFileName(
            betFile.name,
            selectedSite.Title,
            quarter
          );
          console.log(betFileName,"betFileNamebetFileName");
          

        const uploadResult =
          await betService.uploadBETFile(
            betFile,
            betFileName
          );
          console.log(uploadResult,"uploadResultuploadResult");

       const bETCalculationReslt = await betService.performBETCalculation(
          uploadResult.serverRelativeUrl,
          Number(selectedSiteId)
        );

        console.log(  "BET Calculation Result:",  bETCalculationReslt);



        setLoading(false);

        showDialogMessage(
          "BET file processed successfully.",
          false
        );

      } catch (error: any) {

        setLoading(false);

        showDialogMessage(
  error?.message ??
  "BET file validation failed.",
  true
);

        return;
      }
    }

    // =============================
    // STEP 3 : CLEAR FORM
    // =============================

    setExcelFile(null);
    setBetFile(null);
    setSelectedSiteId(undefined);

    if (excelFileRef.current) {
      excelFileRef.current.value = "";
    }

    if (betFileRef.current) {
      betFileRef.current.value = "";
    }

  } catch (error: any) {

    setLoading(false);

    console.error(error);

    showDialogMessage(
  error?.message ??
  "Some error occurred while importing.",
  true
);
  }
};

  const handleExport = async (): Promise<void> => {
    try {
      if (!selectedSiteId) {
        showDialogMessage("Please select a site", true);
        return;
      }

      setLoading(true);

      const selectedSite = sitesOptions.find(
        site => site.Id === selectedSiteId
      );

      const claimType = selectedSite?.Claim_Type;


      // service call
      /*  const response = await spService.exportOutsideCounsel(selectedSite.Id,selectedSite.RSRSSiteName,   claimType,   template   );
    */

      console.log(selectedSite, "selectedSiteselectedSite");


      const buffer = await exportWorkbook({
        siteName: selectedSite.RSRSSiteName,
        claimType: claimType,
        templateFileName:
          claimType ===
            "Indemnification"
            ? "Site Update Template Indemnification.xlsx"
            : "Site Update Template.xlsx",
      });
      console.log(buffer, "bufferbufferbuffer");


      const today = new Date();

      const filename =
        `${selectedSite.RSRSSiteName} ${today.getDate()
        }-${today.getMonth() + 1}-${today.getFullYear()
        }.xlsx`;

      /*       const blob = new Blob([buffer],
              {
                type:
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              }
            );
      
            const link =
              document.createElement("a");
      
            link.href =
              URL.createObjectURL(blob);
      
            link.download = filename;
      
            link.click(); */
      showDialogMessage(
  "File exported successfully.",
  false
);
    } catch (error) {
      console.error(error);

      showDialogMessage(
  "Some error occurred while exporting the file.",
  true
);
    } finally {
      setLoading(false);
    }
  };
  React.useEffect(() => {
    loadSites().catch(console.error);
  }, []);
  const selectedSite = sitesOptions.find(site => site.Id === selectedSiteId);
  const primaryButtonStyles = { root: { borderRadius: "0.4rem" } };

  return (

    <>
      <RsrsHeader context={props.context} />
<Dialog
  hidden={!showDialog}
  onDismiss={() =>
    setShowDialog(false)
  }
 // modalProps={{    isBlocking: true  }}
>
  <div
    style={{
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      textAlign: "center",
      padding: "10px"
    }}
  >
    {isErrorDialog ? (
      <ErrorCircle24Regular
        style={{
          fontSize: "50px",
          color: "#d13438",
          marginBottom: "10px"
        }}
      />
    ) : (
      <CheckmarkCircle24Regular
        style={{
          fontSize: "50px",
          color: "#107c10",
          marginBottom: "10px"
        }}
      />
    )}

    <div
      style={{
        fontSize: "22px",
        fontWeight: 600,
        color: isErrorDialog
          ? "#d13438"
          : "#107c10",
        marginBottom: "15px"
      }}
    >
      {isErrorDialog
        ? "Error"
        : "Success"}
    </div>

    <div
      style={{
        fontSize: "14px",
        color: "#323130",
        lineHeight: "1.5",
        marginBottom: "20px"
      }}
    >
      {message}
    </div>
  </div>

  <DialogFooter
    styles={{
      actions: {
        justifyContent: "center"
      }
    }}
  >
    <PrimaryButton
      text="OK"
      onClick={() =>
        setShowDialog(false)
      }
    />
  </DialogFooter>
</Dialog>

      <section className={styles.outsideCounselUpdates}>
        {loading && (
          <div
            style={{
              position: "fixed",
              inset: 0,
              background: "rgba(255,255,255,0.5)",
              backdropFilter: "blur(4px)",
              zIndex: 99999,
              display: "flex",
              justifyContent: "center",
              alignItems: "center"
            }}
          >
            <Spinner
              size={SpinnerSize.large}
              label="Please Wait..."
            />
          </div>
        )}
        <div className={styles.header}>
          <Link href={`${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`}>
            <Home32Regular className={styles.icon} />
          </Link>

          <span className={styles.title}>
            Import/Export Excel for Outside Counsel
          </span>
        </div>
{/*         {message && (
          <MessageBar
            messageBarType={messageType}
            isMultiline={false}
            onDismiss={() => setMessage('')}
          >
            {message}
          </MessageBar>
        )} */}
        <div className={styles.formContainer}>
          <div className={styles.row}>
            <label className={styles.label}>Select Action:</label>

            <div className={styles.radioGroup}>
              <label>
                <input
                  type="radio"
                  name="action"
                  checked={action === 'export'}
                  onChange={() => setAction('export')}
                />
                Export Site Updates
              </label>

              <label>
                <input
                  type="radio"
                  name="action"
                  checked={action === 'import'}
                  onChange={() => setAction('import')}
                />
                Import Site Updates
              </label>
            </div>
          </div>
          <div className={styles.row}>
            <label className={styles.label}>Select Site:</label>

            <Dropdown
              placeholder="Select a Site"
              className={styles.dropdown}
              options={sitesOptions.map(site => ({
                key: site.Id,
                text: site.RSRSSiteName
              }))}
              selectedKey={selectedSiteId}
              onChange={(_, option) => {
                setSelectedSiteId(option?.key);
              }}
            />
            {action === 'export' && (

              <PrimaryButton text='Export' onClick={handleExport}
                disabled={loading} styles={primaryButtonStyles} />
            )}
          </div>

          {action === 'export' && (
            <>

              {selectedSite?.BETFileUrl && (
                <div className={styles.row}>
                  <label className={styles.label}>
                    Download Existing BET File:
                  </label>

                  <a href={`${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${selectedSite.BETFileUrl}`}
                    target="_blank"
                    rel="noreferrer"
                  >
                    {selectedSite.BETFileUrl.split('/').pop()}
                  </a>
                </div>
              )}
            </>
          )}

          {action === 'import' && (
            <>
              <div className={styles.row}>
                <label className={styles.label}>
                  Upload Updated Excel File:
                </label>

                <input
                  type="file"
                  ref={excelFileRef}
                  // accept=".xls,.xlsx"
                  onChange={(e) => {
                    setExcelFile(e.target.files?.[0] ?? null);
                  }}
                />
              </div>

              <div className={styles.row}>
                <label className={styles.label}>
                  Upload(BET) Excel File(Optional):
                </label>

                <input
                  type="file"
                  ref={betFileRef}
                  // accept=".xlsm"
                  onChange={(e) => {
                    setBetFile(e.target.files?.[0] ?? null);
                  }}
                />
              </div>
            </>

          )}

          <div className={styles.buttonRow}>
            {action === 'import' && (<PrimaryButton text='Upload' styles={primaryButtonStyles} onClick={handleImport} />)}

            <DefaultButton
              text="Cancel"
              styles={primaryButtonStyles}
              onClick={() => { window.location.href = `${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`; }}
            />

          </div>



        </div>

      </section>

    </>
  );

}

