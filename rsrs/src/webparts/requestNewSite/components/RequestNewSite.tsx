import * as React from 'react';
import styles from './RequestNewSite.module.scss';
import type { IRequestNewSiteProps } from './IRequestNewSiteProps';
import { SpService } from '../../../services/spService';
import moment from "moment";
import { useForm, Controller } from "react-hook-form";
import { Link, MessageBar, MessageBarType, Spinner, SpinnerSize } from "@fluentui/react";
import { useEffect, useState } from 'react';
import { DatePicker, IconButton, IStyle } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IComboBoxOption, ComboBox, Stack, Text, TextField, Dropdown, IDropdownOption, PrimaryButton, DefaultButton, Separator, Icon } from '@fluentui/react';
import RsrsHeader from '../../rsrsHeader/components/RsrsHeader';
import { Home32Regular } from "@fluentui/react-icons";
import { TooltipHost, DirectionalHint } from "@fluentui/react";

import {
  Dialog,
  DialogType,
  DialogFooter
} from "@fluentui/react";
import { commonDropdownStyles, customComboBoxStyles, customDropdownStyles, customPickerStyles, formStyles } from '../../../common/styles/controlStyles';

const siteActivityOptions: IDropdownOption[] = [
  { key: "", text: "Select Site Activity" },
  { key: 'Open', text: 'Open' },
  { key: 'Openmonitoronly', text: 'Open - Monitor Only' },
  { key: 'Closed', text: 'Closed' }
];

const claimTypeOptions: IDropdownOption[] = [
  { key: "", text: "Select Claim Type" },
  { key: 'Indemnification', text: 'Indemnification' },
  { key: 'Remediation', text: 'Remediation' },
  { key: 'ToxicTort', text: 'Toxic Tort' }
];

const roleOptions: IDropdownOption[] = [
  { key: "", text: "Select Role" },
  { key: "Assistant Project Manager", text: "Assistant Project Manager" },
  { key: "Outside Counsel", text: "Outside Counsel" },
  { key: "Technical Consultant", text: "Technical Consultant" },
  { key: "Other", text: "Other" }
];

const costEstimateMethodOptions: IDropdownOption[] = [
  { key: "", text: "Select Cost Estimation Method" },
  { key: "FAS 5", text: "FAS 5" },
  { key: "Fair Value", text: "Fair Value" }
];
const obligationTypeOptions: IDropdownOption[] = [
  { key: "", text: "Select Obligation Type" },
  { key: "Disclosure Only", text: "Disclosure Only" },
  { key: "Full Exception", text: "Full Exception" },
  { key: "Recognition", text: "Recognition" }
];


interface ISupportRow {
  id: number;
  support?: string | number;
  role?: string;
}

interface IFormData {
  siteLead: any[];
  siteSupport: any[];
  siteUpdateOwner: any[];
  siteName: string;
  claimType: string;
  country: string;
  city: string;
  state: string;
  siteType: string;
  siteActivity: string;
  legacyCompany: string;
  costEstimateMethod: string;
  originalCompanyConnection: string;
  obligationType?: string;
  operatingGroup?: string;
  transactionDate?: Date;
}

export default function RequestNewSite(props: IRequestNewSiteProps): React.ReactElement<IRequestNewSiteProps> {
  const [countryOptions, setCountryOptions] = useState<IDropdownOption[]>([]);
  // const [selectedCountry, setSelectedCountry] = useState<string | number>();
  const [siteTypeOptions, setSiteTypeOptions] = useState<IDropdownOption[]>([]);
  const [selectedSiteType, setSelectedSiteType] = useState<string | number>();
  const [legacyCompaniesOptions, setLegacyCompaniesOptions] = useState<IDropdownOption[]>([]);
  const [contactsOptions, setContactsOptions] = useState<IDropdownOption[]>([]);
  const [siteLead, setSiteLead] = useState<any[]>([]);
  const [siteSupport, setSiteSupport] = useState<any[]>([]);
  const [siteUpdateOwner, setSiteUpdateOwner] = useState<any[]>([]);
  const [operatingGroupOptions, setOperatingGroupOptions] = useState<IDropdownOption[]>([]);
  const [supportRows, setSupportRows] = useState<ISupportRow[]>([{ id: Date.now(), support: "", role: "" }]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isSubmitted, setIsSubmitted] = useState(false);
  const [message, setMessage] = useState("");
  const [messageType, setMessageType] = useState<MessageBarType | undefined>();
  const [isAdmin, setIsAdmin] = useState(false);
  const [isProcessed, setIsProcessed] = useState(false);
  const [pageMode, setPageMode] = useState<"create" | "view" | "approval">("create");
  const [generatedSiteId, setGeneratedSiteId] = useState("");
  const [defaultSiteLead, setDefaultSiteLead] = useState<string[]>([]);
  const [defaultSiteSupport, setDefaultSiteSupport] = useState<string[]>([]);
  const [defaultSiteUpdateOwner, setDefaultSiteUpdateOwner] = useState<string[]>([]);
  const [createdBy, setCreatedBy] = useState("");
  const [showDialog, setShowDialog] = useState(false);
  const [redirectUrl, setRedirectUrl] = useState("");
  const [isErrorDialog, setIsErrorDialog] = useState(false);
  const [securityLevel, setSecurityLevel] = useState<string>("");
  const isReadOnly = pageMode === "view" || isProcessed;
  const {
    control,
    handleSubmit,
    watch,
    reset,
    formState: { errors }
  } = useForm<IFormData>({
    defaultValues: {
      siteLead: [],
      siteSupport: [],
      siteUpdateOwner: [],
      siteName: "",
      claimType: "",
      country: "",
      city: "",
      state: "",
      siteType: "",
      siteActivity: "",
      legacyCompany: "",
      costEstimateMethod: "",
      originalCompanyConnection: "",
      obligationType: "",
      operatingGroup: ""
    }
  });
  const selectedClaimType = watch("claimType");
  const updateSupportRow = (
    id: number,
    field: "support" | "role",
    value: any
  ): void => {

    setSupportRows(prev =>
      prev.map(row =>
        row.id === id
          ? {
            ...row,
            [field]: value
          }
          : row
      )
    );
  };

  const spService = new SpService(props.context);
  const requestId = new URLSearchParams(window.location.search).get("RequestId");
  const textFieldStyles = isReadOnly ? formStyles.readOnlyField : formStyles.textField;
  const multiLineFieldStyles = isReadOnly ? formStyles.readOnlyMultiLineField : formStyles.multilineField;
  const dropdownStyles = isReadOnly ? customDropdownStyles : formStyles.dropdown;
  const comboBoxStyles = isReadOnly ? customComboBoxStyles : formStyles.comboBox;
  const peoplepickerstyles = isReadOnly ? customPickerStyles : {}
  const navigateToAccessDenied = (): void => {
    window.location.href =
      `${props.context.pageContext.web.absoluteUrl}/SitePages/AccessDenied.aspx`;
  };
  const validateAccess = async (): Promise<string> => {
    try {

      const currentUserEmail =
        props.context.pageContext.user.email;

      const contacts =
        await spService.getItems(
          props.ContactsListName,
          [
            "Id",
            "Title",
            "ContactEmail",
            "User_Security_Level"
          ],
          `ContactEmail eq '${currentUserEmail}'`
        );

      // No Contact Record
      if (!contacts || contacts.length === 0) {

        navigateToAccessDenied();

        return "";

      }

      const role =
        contacts[0].User_Security_Level;

      setSecurityLevel(role);

      // No Access
      if (
        role === "No Access"
      ) {

        navigateToAccessDenied();

        return "";

      }

      return role;

    }
    catch (error) {

      console.error(error);

      navigateToAccessDenied();

      return "";

    }
  };

  const checkAdmin = async (): Promise<void> => {

    try {

      const role =
        await validateAccess();

      if (!role) {
        return;
      }

      const userIsAdmin =
        role === "Administrator";

      setIsAdmin(userIsAdmin);

      if (!requestId) {

        setPageMode("create");

      }
      else if (
        requestId &&
        userIsAdmin
      ) {

        setPageMode("approval");

      }
      else {

        setPageMode("view");

      }

    }
    catch (error) {

      console.error(error);

      navigateToAccessDenied();

    }

  };

  const removeSupportRow = (id: number): void => {
    setSupportRows(prev => prev.filter(row => row.id !== id));
  };
  const navigateToHome = (): void => {
    window.location.href =
      `${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`;
  };

  const loadAdditionalSupport =
    async (
      requestId: number
    ): Promise<void> => {

      const items =
        await spService.getItems(
          "AdditionalSupport",
          [
            "Id",
            "AdditionalContactId",
            "AdditionalSupportRole"
          ],
          `SiteRequestIdId eq ${requestId}`
        );

      if (items.length === 0) {
        return;
      }

      setSupportRows(

        items.map(
          (item: any) => ({

            id: item.Id,
            support: Number(item.AdditionalContactId),
            // support: item.AdditionalContactId,

            role:
              item.AdditionalSupportRole

          })
        )

      );

    };

  const getSiteIdDetails = async (
    siteType: string
  ): Promise<void> => {

    const result = await spService.getItems(
      props.siteTypeListName,
      [
        "Id",
        "Title",
        "RSRSIDPrefix",
        "NextSiteID"
      ],
      `Title eq '${siteType}'`
    );

    if (result.length > 0) {

      const siteId =
        `${result[0].RSRSIDPrefix}-${result[0].NextSiteID}`;

      setGeneratedSiteId(siteId);
    }
  };
  const loadSiteAdditionalSupport = async (
    siteId: number
  ): Promise<void> => {
    const items = await spService.getItems(
      "SiteAdditionalSupport",
      [
        "Id",
        "AdditionalContactId",
        "AdditionalSupportRole"
      ],
      `SiteIDId eq ${siteId}`
    );
    if (items.length === 0) {
      return;
    }



    setSupportRows(
      items.map((item: any) => ({
        id: item.Id,
        support: Number(item.AdditionalContactId),
        //support: item.AdditionalContactId?.toString(),
        role: item.AdditionalSupportRole
      }))
    );
  };




  const loadRequest = async (id: number): Promise<void> => {

    const item =
      await spService.getItemByIdd(
        "NewSiteRequest",
        id
      );

    setCreatedBy(item.Author?.Title || "");

    if (item.SiteLead?.length > 0) {
      setDefaultSiteLead([
        item.SiteLead[0].EMail
      ]);
    }

    if (item.SiteSupport?.length > 0) {
      setDefaultSiteSupport([
        item.SiteSupport[0].EMail
      ]);
    }

    if (item.SiteUpdateOwner?.length > 0) {
      setDefaultSiteUpdateOwner([
        item.SiteUpdateOwner[0].EMail
      ]);
      setSiteUpdateOwner(
        item.SiteUpdateOwner
      );
    }


    setIsProcessed(
      item.IsCreated === true
    );



    reset({
      siteLead: item.SiteLead || [],
      siteSupport: item.SiteSupport || [],
      siteUpdateOwner: item.SiteUpdateOwner || [],
      siteName:
        item.RSRSSiteName,

      city:
        item.City,

      state:
        item.State,

      country:
        item.Country,

      claimType:
        item.Claim_Type,

      siteType:
        item.Site_Type,

      siteActivity:
        item.Site_Activity,

      legacyCompany:
        item.Legacy_Company,

      originalCompanyConnection:
        item.OriginalCompanyConnection,

      costEstimateMethod:
        item.Cost_Estimate_Method,

      obligationType:
        item.ObligationType,

      operatingGroup:
        item.OperatingGroup,


      transactionDate:
        item.TransactionDate
          ? new Date(
            item.TransactionDate
          )
          : undefined

    });

    if (
      item.IsCreated === true &&
      item.IsApproved === "Approved"
    ) {

      setGeneratedSiteId(
        item.RSRSSiteId?.Title || ""
      );
      if (item.RSRSSiteIdId) {
        await loadSiteAdditionalSupport(
          item.RSRSSiteIdId
        );

      }

    } else {

      await loadAdditionalSupport(id);

      if (item.Site_Type) {

        await getSiteIdDetails(
          item.Site_Type
        );

      }

    }

  };
  const navigateToDashboard = (): void => {
    window.location.href =
      `${props.context.pageContext.web.absoluteUrl}/SitePages/dashboard.aspx`;
  };


  const getSiteTypeData = async (
    siteType: string
  ): Promise<any> => {

    const items = await spService.getItems(
      props.siteTypeListName,
      [
        "Id",
        "Title",
        "RSRSIDPrefix",
        "NextSiteID"
      ],
      `Title eq '${siteType}'`
    );

    return items[0];

  };

  const getCompanyCodeData = async (
    company: string
  ): Promise<any> => {

    const items = await spService.getItems(
      "CompanyCodes",
      [
        "*"
      ],
      `Title eq '${company}'`
    );

    return items[0];

  };
  const validateUpdateOwnerOrOutsideCounselOnly = (): boolean => {

    const hasSiteUpdateOwner =
      siteUpdateOwner?.length > 0;

    const hasOutsideCounsel =
      supportRows.some(
        row =>
          row.support &&
          row.role === "Outside Counsel"
      );

    if (!hasSiteUpdateOwner && !hasOutsideCounsel) {

      setMessageType(MessageBarType.error);

      setMessage(
        "Please select either Site Update Owner or at least one Additional Support as Outside Counsel"
      );

      setIsErrorDialog(true);
      setShowDialog(true);

      return false;
    }

    return true;
  };
  const validateUpdateOwnerOrOutsideCounsel = (
    data: IFormData
  ): boolean => {



    const hasSiteUpdateOwner =
      data.siteUpdateOwner?.length > 0;

    const hasOutsideCounsel =
      supportRows.some(
        row =>
          row.support &&
          row.role === "Outside Counsel"
      );

    if (!hasSiteUpdateOwner && !hasOutsideCounsel) {

      setMessageType(MessageBarType.error);
      setMessage(
        "Please select either Site Update Owner or at least one Additional Support as Outside Counsel"
      );
      setIsErrorDialog(true);
      setShowDialog(true);

      return false;
    }

    return true;
  };

  const approveRequest = async (
    data: IFormData
  ): Promise<void> => {

    if (!requestId) {
      return;
    }
    try {

      setIsSubmitting(true);

      // Get original request (only for additional support copy/update)
      const request =
        await spService.getItemByIdd(
          "NewSiteRequest",
          Number(requestId)
        );
      const siteTypeText =
        siteTypeOptions.find(
          x => x.key === data.siteType
        )?.text ?? "";

      const legacyCompanyText =
        legacyCompaniesOptions.find(
          x => x.key === data.legacyCompany
        )?.text ?? "";

      const siteTypeData =
        await getSiteTypeData(siteTypeText);

      const companyData =
        await getCompanyCodeData(
          legacyCompanyText
        );

      const generatedSiteIdValue =
        `${siteTypeData.RSRSIDPrefix}-${siteTypeData.NextSiteID}`;

      const digit =
        companyData.Digit === 99
          ? 99
          : 98;

      const claimTypeText =
        claimTypeOptions.find(
          x => x.key === data.claimType
        )?.text ?? "";

      const claimPrefix =
        claimTypeText === "Indemnification"
          ? "IND"
          : "EPA";

      const thirdDigit =
        claimTypeText === "Indemnification"
          ? 0
          : companyData.Digit;

      const accountProjectCode =
        `BLLG${claimPrefix}${digit}${thirdDigit}${companyData.usct}`;

      const accountCodeBlock =
        `${companyData.LegacyCompanyCode}-11001190-66-8011-${accountProjectCode}-EFLEG`;

      // Resolve People Picker values

      const siteLeadUser =
        data.siteLead?.length > 0
          ? await spService.ensureUser(data.siteLead[0].loginName || data.siteLead[0].EMail) : null;

      const siteSupportUser =
        data.siteSupport?.length > 0
          ? await spService.ensureUser(
            data.siteSupport[0].loginName || data.siteSupport[0].EMail
          )
          : null;

      const siteUpdateOwnerUser =
        data.siteUpdateOwner?.length > 0
          ? await spService.ensureUser(
            data.siteUpdateOwner[0].loginName || data.siteUpdateOwner[0].EMail
          )
          : null;

      const sitePayload: any = {

        Title: generatedSiteIdValue,

        RSRSSiteName:
          data.siteName,

        City:
          data.city,

        State:
          data.state,

        Country:
          countryOptions.find(
            x => x.key === data.country
          )?.text ?? "",
        Site_Location: [data.city, data.state, countryOptions.find(x => x.key === data.country)?.text].filter(Boolean).join(", "),


        Claim_Type:
          claimTypeText,

        Site_Type:
          siteTypeText,

        Site_Activity:
          siteActivityOptions.find(
            x => x.key === data.siteActivity
          )?.text ?? "",

        Legacy_Company:
          legacyCompanyText,

        Legacy_Company_Code:
          companyData.LegacyCompanyCode,

        OriginalCompanyConnection:
          data.originalCompanyConnection,

        Cost_Estimate_Method:
          costEstimateMethodOptions.find(
            x => x.key === data.costEstimateMethod
          )?.text ?? "",

        Corporate_Environmental_Res_Acco:
          "11001190",

        AccountProjectCode:
          accountProjectCode,

        AccountCodeBlock:
          accountCodeBlock,

        RequestedBy:
          request.Author?.Title,

        NewSiteRequest: "Yes",

        DateSiteAdded:
          new Date(),

        ObligationType:
          data.claimType === "Indemnification"
            ? obligationTypeOptions.find(
              x => x.key === data.obligationType
            )?.text ?? ""
            : "",

        OperatingGroup:
          data.claimType === "Indemnification"
            ? operatingGroupOptions.find(
              x => x.key === data.operatingGroup
            )?.text ?? ""
            : "",

        TransactionDate:
          data.transactionDate
            ? moment(
              data.transactionDate
            ).toISOString()
            : null
      };

      if (siteLeadUser) {
        sitePayload.SiteLeadId = [
          siteLeadUser.Id
        ];
      }

      if (siteSupportUser) {
        sitePayload.SiteSupportId = [
          siteSupportUser.Id
        ];
      }

      if (siteUpdateOwnerUser) {
        sitePayload.SiteUpdateOwnerId = [
          siteUpdateOwnerUser.Id
        ];
      }

      const site =
        await spService.createItem(
          "Site",
          sitePayload
        );

      for (const row of supportRows) {

        if (row.support && row.role) {

          await spService.createItem(
            "SiteAdditionalSupport",
            {
              SiteIDId: site.Id,
              AdditionalContactId: row.support,
              AdditionalSupportRole: row.role,
              IsActive: "true"
            }
          );

        }

      }

      await spService.updateItem(
        props.siteTypeListName,
        siteTypeData.Id,
        {
          NextSiteID:
            (
              Number(
                siteTypeData.NextSiteID
              ) + 1
            ).toString()
        }
      );

      await spService.updateItem(
        "CompanyCodes",
        companyData.Id,
        {
          usct:
            (
              Number(companyData.usct) + 1
            ).toString()
        }
      );

      await spService.updateItem(
        "NewSiteRequest",
        Number(requestId),
        {
          IsCreated: true,
          IsApproved: "Approved",
          RSRSSiteIdId: site.Id
        }
      );

      setIsProcessed(true);

      setMessageType(MessageBarType.success);

      setMessage(
        "Request Approved Successfully"
      );

      setRedirectUrl(
        `${props.context.pageContext.web.absoluteUrl}/SitePages/Dashboard.aspx`
      );

      setIsErrorDialog(false);
      setShowDialog(true);

    } catch (error) {

      console.error(error);

      setMessageType(
        MessageBarType.error
      );

      setMessage("Error occurred during approval."
      );

      setRedirectUrl("");
      setIsErrorDialog(true);
      setShowDialog(true);
    } finally {

      setIsSubmitting(false);

    }
  };


  /*   const approveRequest = async (): Promise<void> => {
  
      if (!requestId) {
        return;
      }
  
      try {
  
        setIsSubmitting(true);
  
        const request =
          await spService.getItemByIdd(
            "NewSiteRequest",
            Number(requestId)
          );
  
        const siteTypeData =
          await getSiteTypeData(
            request.Site_Type
          );
  
        const companyData =
          await getCompanyCodeData(
            request.Legacy_Company
          );
  
        // Generate Site Id
  
        const generatedSiteIdValue =
          `${siteTypeData.RSRSIDPrefix}-${siteTypeData.NextSiteID}`;
  
        // Generate Account Project Code
  
        const digit =
          companyData.Digit === 99
            ? 99
            : 98;
  
        const claimPrefix =
          request.Claim_Type === "Indemnification"
            ? "IND"
            : "EPA";
  
        const thirdDigit =
          request.Claim_Type === "Indemnification"
            ? 0
            : companyData.Digit;
  
        const accountProjectCode =
          `BLLG${claimPrefix}${digit}${thirdDigit}${companyData.usct}`;
  
        // Generate Account Code Block
  
        const accountCodeBlock =
          `${companyData.LegacyCompanyCode}-11001190-66-8011-${accountProjectCode}-EFLEG`;
  
        // Create Site
  
        const sitePayload = {
  
          Title: generatedSiteIdValue,
  
          RSRSSiteName:
            request.RSRSSiteName,
  
          SiteLeadId:
            request.SiteLeadId,
  
          SiteSupportId:
            request.SiteSupportId,
  
          SiteUpdateOwnerId:
            request.SiteUpdateOwnerId,
  
          City:
            request.City,
  
          State:
            request.State,
  
          Country:
            request.Country,
  
          Claim_Type:
            request.Claim_Type,
  
          Site_Type:
            request.Site_Type,
  
          Site_Activity:
            request.Site_Activity,
  
          Legacy_Company:
            request.Legacy_Company,
  
          Legacy_Company_Code:
            companyData.LegacyCompanyCode,
  
          OriginalCompanyConnection:
            request.OriginalCompanyConnection,
  
          Cost_Estimate_Method:
            request.Cost_Estimate_Method,
  
          Corporate_Environmental_Res_Acco:
            "11001190",
  
          AccountProjectCode:
            accountProjectCode,
  
          AccountCodeBlock:
            accountCodeBlock,
  
          RequestedBy:
            props.context.pageContext.user.displayName,
  
          DateSiteAdded:
            new Date(),
  
          ObligationType:
            request.ObligationType,
  
          OperatingGroup:
            request.OperatingGroup,
  
          TransactionDate: request.TransactionDate ? moment(request.TransactionDate, "MM-DD-YYYY").toISOString() : null
  
        };
        const site =
          await spService.createItem(
            "Site",
            sitePayload
          );
  
        // Copy Additional Support
  
        const supports =
          await spService.getItems(
            "AdditionalSupport",
            ["*"],
            `SiteRequestIdId eq ${requestId}`
          );
        for (const support of supports) {
  
          await spService.createItem(
            "SiteAdditionalSupport",
            {
  
              SiteIDId:
                site.Id,
  
              AdditionalContactId:
                support.AdditionalContactId,
  
              AdditionalSupportRole:
                support.AdditionalSupportRole,
  
              IsActive: "true"
  
            }
          );
  
        }
  
        // Update SiteType Counter
  
        await spService.updateItem(
          props.siteTypeListName,
          siteTypeData.Id,
          {
  
            NextSiteID:
              (Number(
                siteTypeData.NextSiteID
              ) + 1).toString()
  
          }
        );
  
        // Update Company Counter
  
        await spService.updateItem(
          "CompanyCodes",
          companyData.Id,
          {
  
            usct:
              Number(
                companyData.usct
              ) + 1
  
          }
        );
  
        // Update Request
  
        await spService.updateItem(
          "NewSiteRequest",
          Number(requestId),
          {
  
            IsCreated: true,
  
            IsApproved: "Approved",
  
            RSRSSiteIdId:
              site.Id
  
          }
        );
  
        setIsProcessed(true);
        setMessageType(
          MessageBarType.success
        );
  
        setMessage(
          "Request Approved Successfully"
        );
  
        setTimeout(() => {
  
          navigateToDashboard();
  
        }, 1500);
      }
      catch (error) {
  
        console.error(error);
  
        setMessageType(
          MessageBarType.error
        );
  
        setMessage(
          "Error occurred during approval."
        );
        setTimeout(() => {
          setMessage("");
        }, 3000);
  
      }
      finally {
  
        setIsSubmitting(false);
  
      }
  
    }; */

  const rejectRequest =
    async (): Promise<void> => {

      if (!requestId) {
        return;
      }

      try {

        await spService.updateItem(
          "NewSiteRequest",
          Number(requestId),
          {
            IsCreated: true,
            IsApproved: "Rejected"
          }
        );

        setMessageType(
          MessageBarType.success
        );

        setMessage(
          "Request Rejected Successfully"
        );

        setRedirectUrl(
          `${props.context.pageContext.web.absoluteUrl}/SitePages/Dashboard.aspx`
        );

        setIsErrorDialog(false);
        setShowDialog(true);


      }
      catch (error) {

        console.error(error);
        setMessageType(
          MessageBarType.error
        );

        setMessage("Error occurred during Reject."
        );

        setRedirectUrl("");
        setIsErrorDialog(true);
        setShowDialog(true);

      }

    };

  const loadOperatingGroups = async (): Promise<void> => {
    try {
      const items = await spService.getItems(props.OperatingGroupListName, ["Id", "Title"], undefined, { field: "Title", ascending: true });
      const options: IDropdownOption[] = [
        {
          key: "",
          text: "Select Operating Group"
        },
        ...items.map(item => ({
          key: item.Title,
          text: item.Title
        }))
      ];

      setOperatingGroupOptions(options);
    } catch (error) {
      console.error("Error loading countries:", error);
    }
  };

  const loadContacts = async (): Promise<void> => {
    try {
      const items = await spService.getItems(props.ContactsListName, ["Id", "Title"], undefined, { field: "Title", ascending: true });
      const options: IDropdownOption[] = [
        {
          key: "",
          text: "Select Name"
        },
        ...items.map(item => ({
          key: item.Id,
          text: item.Title
        }))
      ]; setContactsOptions(options);
    } catch (error) {
      console.error("Error loading countries:", error);
    }
  };

  const loadLegacyCompanies = async (): Promise<void> => {
    try {
      const items = await spService.getItems(props.LegacyCompanyListName, ["Id", "Title"], undefined, { field: "Title", ascending: true });
      const options: IDropdownOption[] = [
        {
          key: "",
          text: "Select Legacy Company "
        },
        ...items.map(item => ({
          key: item.Title,
          text: item.Title
        }))
      ];
      setLegacyCompaniesOptions(options);
    } catch (error) {
      console.error("Error loading countries:", error);
    }
  };

  async function loadCountries(): Promise<void> {
    try {
      const items = await spService.getItems(props.countryListName, ["Id", "Title"], undefined, { field: "Title", ascending: true });
      const options: IDropdownOption[] = [
        {
          key: "",
          text: "Select Country "
        },
        ...items.map(item => ({
          key: item.Title,
          text: item.Title
        }))
      ];
      setCountryOptions(options);
    } catch (error) {
      console.error("Error loading countries:", error);
    }
  }



  const loadSiteTypes = async (): Promise<void> => {
    try {

      const items = await spService.getItems(props.siteTypeListName, ["Id", "Title"]);
      const options: IDropdownOption[] = [
        {
          key: "",
          text: "Select Site Type"
        },
        ...items.map(item => ({
          key: item.Title,
          text: item.Title
        }))
      ];
      setSiteTypeOptions(options);
    } catch (error) {
      console.error("Error loading SiteTypes:", error);
    }
  };


  const onSubmit = async (data: IFormData): Promise<void> => {
    /* setIsSubmitting(true);
    setMessage("");
    const hasSiteUpdateOwner =data.siteUpdateOwner?.length > 0;
const hasOutsideCounsel =supportRows.some( row =>  row.support &&   row.role === "Outside Counsel" );
if (!hasSiteUpdateOwner && !hasOutsideCounsel) {
  setOwnerValidationError(
    "Please select either Site Update Owner or at least one Additional Support as Outside Counsel"
  );
  setIsSubmitting(false);

  return;
}

setOwnerValidationError(""); */

    setIsSubmitting(true);
    try {
      const siteLeadUser = data.siteLead?.length > 0 ? await spService.ensureUser(data.siteLead[0].loginName) : null;

      const siteSupportUser = siteSupport?.length > 0 ? await spService.ensureUser(siteSupport[0].loginName) : null;

      const siteUpdateOwnerUser = data.siteUpdateOwner?.length > 0 ? await spService.ensureUser(data.siteUpdateOwner[0].loginName) : null;



      const payload: any = {
        RSRSSiteName: data.siteName,
        City: data.city,
        State: data.state,
        // Site_Location:`${data.city || ""}${data.city && data.state ? ", " : ""}${data.state || ""}`,
        Site_Location: [data.city, data.state, countryOptions.find(x => x.key === data.country)?.text].filter(Boolean).join(", "),
        Country: countryOptions.find(x => x.key === data.country)?.text ?? "",
        Claim_Type: claimTypeOptions.find(x => x.key === data.claimType)?.text ?? "",
        Site_Type: siteTypeOptions.find(x => x.key === data.siteType)?.text ?? "",
        //Site_Type: data.siteType,
        Site_Activity: siteActivityOptions.find(x => x.key === data.siteActivity)?.text ?? "",
        Legacy_Company: legacyCompaniesOptions.find(x => x.key === data.legacyCompany)?.text ?? "",
        OriginalCompanyConnection: data.originalCompanyConnection,
        Cost_Estimate_Method: costEstimateMethodOptions.find(x => x.key === data.costEstimateMethod)?.text ?? "",
        IsCreated: false,
        //IsApproved: "",
        ObligationType: data.claimType === "Indemnification"
          ? obligationTypeOptions.find(x => x.key === data.obligationType)?.text ?? ""
          : "",
        OperatingGroup:
          data.claimType === "Indemnification"
            ? operatingGroupOptions.find(x => x.key === data.operatingGroup)?.text ?? ""
            : "",
        TransactionDate:
          data.claimType === "Indemnification" &&
            data.transactionDate
            ? moment(data.transactionDate).format("MM-DD-YYYY")
            : ""
      };
      if (siteLeadUser) {
        payload.SiteLeadId = [siteLeadUser.Id];
      }

      if (siteSupportUser) {
        payload.SiteSupportId = [siteSupportUser.Id];
      }

      if (siteUpdateOwnerUser) {
        payload.SiteUpdateOwnerId = [siteUpdateOwnerUser.Id];
      }

      const response = await spService.createItem(
        "NewSiteRequest",
        payload
      );
      const requestId = response.Id;


      // Additional Support Save
      for (const row of supportRows) {

        if (row.support && row.role) {

          await spService.createItem(
            "AdditionalSupport",
            {
              SiteRequestIdId: requestId,
              AdditionalContactId: row.support,
              AdditionalSupportRole: row.role,
              RecordType: "Initial"
            }
          );
        }
      }
      setMessageType(MessageBarType.success);
      setMessage("Site request submitted successfully.");
      setIsSubmitted(true);

      setRedirectUrl(
        isAdmin
          ? `${props.context.pageContext.web.absoluteUrl}/SitePages/Dashboard.aspx`
          : `${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`
      );

      setIsErrorDialog(false);
      setShowDialog(true);

    } catch (error) {

      console.error(error);

      setMessageType(MessageBarType.error);

      setMessage("Error occurred while saving the request.");

      setRedirectUrl("");
      setIsErrorDialog(true);
      setShowDialog(true);
    }
    finally {
      setIsSubmitting(false);
    }
  };



  const peoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory,
    spHttpClient: props.context.spHttpClient
  } as any;



  const addSupportRow = (): void => {
    setSupportRows(prev => [
      ...prev,
      {
        id: Date.now(),
        support: "",
        role: ""
      }
    ]);
  };








  /*   const cardStyle = {
      background: '#ffffff',
      borderRadius: '16px',
      padding: '24px',
      width: '100%',
      marginBottom: '20px',
      boxSizing: 'border-box' as const,
      boxShadow: '0px 4px 12px rgba(18, 93, 154)'
      // boxShadow: '0px 4px 12px rgba(102, 84, 206, 0.08)'
    }; */

  /*   const cardStyle = {
    background: '#ffffff',
    borderRadius: '16px',
    padding: '24px',
    width: '100%',
    marginBottom: '20px',
    boxSizing: 'border-box' as const,
  // borderBottom: '4px solid #125D9A',
    boxShadow: '0 2px 8px rgba(0,0,0,0.08)'
  }; */
  const sectionTitleStyle = {
    fontWeight: "600",
    color: '#0F6CBD'
  };

  useEffect(() => {
    void checkAdmin();
    if (requestId) { void loadRequest(Number(requestId)); }

    loadCountries().catch(console.error);
    loadSiteTypes().catch(console.error);
    loadContacts().catch(console.error);
    loadLegacyCompanies().catch(console.error);
    loadOperatingGroups().catch(console.error);


  }, []);

  // useEffect(() => {

  //   if (
  //     pageMode === "approval"
  //     && requestId
  //   ) {

  //     const loadSiteId =
  //       async () => {

  //         const request =
  //           await spService.getItemByIdd(
  //             "NewSiteRequest",
  //             Number(requestId)
  //           );

  //         await getSiteIdDetails(
  //           request.Site_Type
  //         );

  //       };

  //     void loadSiteId();

  //   }

  // }, [pageMode, requestId]);
  const handleApproveClick = (): void => {

    const currentData = watch();

    validateUpdateOwnerOrOutsideCounsel(currentData);

    void handleSubmit(
      async (data) => {

        if (
          !validateUpdateOwnerOrOutsideCounsel(data)
        ) {
          return;
        }

        await approveRequest(data);

      }
    )();

  };
  const handleCreateClick = (): void => {

    const currentData = watch();

    validateUpdateOwnerOrOutsideCounsel(currentData);

    void handleSubmit(
      async (data) => {

        if (
          !validateUpdateOwnerOrOutsideCounsel(data)
        ) {
          return;
        }

        await onSubmit(data);

      }
    )();

  };







  return (
    <>
      {/* <RsrsHeader context={props.context} /> */}
      <Dialog
        hidden={!showDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: isErrorDialog ? "Error" : "Success"
        }}
      >
        <MessageBar
          messageBarType={messageType}
          isMultiline={true}
        >
          {message}
        </MessageBar>

        <DialogFooter>
          <PrimaryButton
            text="OK"
            onClick={() => {
              setShowDialog(false);

              if (!isErrorDialog && redirectUrl) {
                window.location.href = redirectUrl;
              }
            }}
          />
        </DialogFooter>
      </Dialog>
      <section >
        {isSubmitting && (
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
              label="Submitting Request..."
            />
          </div>
        )}

        {/*         {message && (
          <div
            style={{
              position: "fixed",
              top: 20,
              right: 20,
              zIndex: 100000,
              width: 450
            }}
          >
            <MessageBar
              messageBarType={messageType}
              isMultiline={false}
            >
              {message}
            </MessageBar>
          </div>
        )} */}
        <div className={styles.header}>
          <Link href={`${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`}>
            <Home32Regular className={styles.icon} />
          </Link>

          <span className={styles.title}>
            Request New Site
          </span>
        </div>
        {/*  <div className={styles.header}>
        </div> */}

        {/* <div className={styles.headerDivider}></div> */}
        <div className={styles.requiredFields}>
          [Required Fields]
          <span className={styles.required}>*</span>
        </div>
        <div className={styles.formContainer}>

          <table className={styles.formTable}>
            <tbody>

              {pageMode === "approval" && (
                <tr>
                  <td className={styles.labelCell}>
                    Site ID :
                    <span className={styles.required}>*</span>
                  </td>
                  <td className={styles.fieldCell}>
                    <TextField value={generatedSiteId} readOnly styles={textFieldStyles} />

                  </td>
                </tr>
              )}

              {/* Site Name */}
              <tr>
                <td className={styles.labelCell}>
                  RSRS Site Name :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="siteName"
                    control={control}
                    rules={{
                      required: "RSRS Site Name is required"
                    }}
                    render={({ field }) => (
                      <TextField
                        readOnly={isReadOnly}
                        value={field.value}
                        onChange={(_, value) =>
                          field.onChange(value)
                        }
                        errorMessage={errors.siteName?.message}
                        styles={textFieldStyles}
                      />
                    )}
                  />
                </td>
              </tr>
              {pageMode === "approval" && (
                <tr>
                  <td className={styles.labelCell}>
                    Requested By :
                  </td>
                  <td className={styles.fieldCell}>
                    <TextField
                      // disabled={isReadOnly}
                      readOnly
                      value={createdBy}
                      styles={textFieldStyles}
                    />
                  </td>
                </tr>
              )}

              {/* Site Lead */}
              <tr>
                <td className={styles.labelCell}>
                  Site Lead :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="siteLead"
                    control={control}
                    rules={{
                      validate: value =>
                        value?.length > 0 || "Site Lead is required"
                    }}
                    render={({ field }) => (
                      <>
                        <PeoplePicker
                          placeholder='Enter a name or email address'
                          context={peoplePickerContext}
                          personSelectionLimit={1}
                          disabled={isReadOnly}
                          defaultSelectedUsers={defaultSiteLead}
                          principalTypes={[PrincipalType.User]}
                          required
                          onChange={(items: any) => {
                            field.onChange(items);
                            setSiteLead(items);
                          }}
                          styles={peoplepickerstyles}
                        />

                        {errors.siteLead && (
                          <Text
                            styles={{
                              root: {
                                color: "#d13438",
                                fontSize: 12
                              }
                            }}
                          >
                            {errors.siteLead.message}

                          </Text>
                        )}
                      </>
                    )}
                  />
                </td>
              </tr>

              {/* Site Support */}
              {<tr>
                <td className={styles.labelCell}>
                  Site Support :
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="siteSupport"
                    control={control}
                    render={({ field }) => (
                      <PeoplePicker
                        placeholder='Enter a name or email address'
                        context={peoplePickerContext}
                        personSelectionLimit={1}
                        disabled={isReadOnly}
                        defaultSelectedUsers={defaultSiteSupport}
                        principalTypes={[PrincipalType.User]}
                        onChange={(items) => {
                          field.onChange(items);
                          setSiteSupport(items);
                        }}
                        styles={peoplepickerstyles}
                      />
                    )}
                  />
                </td>
              </tr>}

              {/* Update Owner */}
              <tr>
                <td className={styles.labelCell}>
                  RSRS Update Owner :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="siteUpdateOwner"
                    control={control}
                    /* rules={{
                      validate: value =>
                        value?.length > 0 || "RSRS Update Owner is required"
                    }} */
                    render={({ field }) => (
                      <>
                        <TooltipHost
                          content="Please select a site update owner Or add an outside ounsel"
                          directionalHint={DirectionalHint.topCenter}
                          tooltipProps={{
                            styles: {
                              content: {
                                backgroundColor: "#323130",
                                color: "#ffffff",
                                border: "none"
                              }
                            }
                          }}>

                          <PeoplePicker
                            placeholder='Enter a name or email address'
                            context={peoplePickerContext}
                            personSelectionLimit={1}
                            disabled={isReadOnly}
                            defaultSelectedUsers={defaultSiteUpdateOwner}
                            principalTypes={[PrincipalType.User]}
                            onChange={(items: any) => {
                              field.onChange(items);
                              setSiteUpdateOwner(items);
                            }}
                            styles={peoplepickerstyles}
                          />
                        </TooltipHost>
                      </>
                    )}
                  />
                </td>
              </tr>

              {/* Additional Support */}
              <>
                {supportRows.map((row, index) => (
                  <tr key={row.id}>
                    <td className={styles.labelCell}>
                      {index === 0 ? "Additional Support :" : ""}
                    </td>

                    <td className={styles.fieldCell}>
                      <div className={styles.additionalSupportContainer}>
                        <Dropdown
                          className={styles.supportDropdown}
                          placeholder="Select Name"
                          options={contactsOptions}
                          disabled={isReadOnly}
                          selectedKey={row.support}
                          onChange={(_, option) =>
                            updateSupportRow(
                              row.id,
                              "support",
                              option?.key
                            )
                          }
                          styles={dropdownStyles}
                        />

                        <Dropdown
                          className={styles.roleDropdown}
                          placeholder="Select Role"
                          options={roleOptions}
                          disabled={isReadOnly}
                          selectedKey={row.role}
                          onChange={(_, option) =>
                            updateSupportRow(
                              row.id,
                              "role",
                              option?.key
                            )
                          }
                          styles={dropdownStyles}
                        />

                        {index === 0 ? (
                          <IconButton
                            disabled={isReadOnly}
                            iconProps={{ iconName: "Add" }}
                            title="Add"
                            onClick={addSupportRow}
                          />
                        ) : (
                          <IconButton
                            disabled={isReadOnly}
                            iconProps={{ iconName: "Delete" }}
                            title="Remove"
                            styles={{
                              root: {
                                color: "red"
                              }
                            }}
                            onClick={() => removeSupportRow(row.id)}
                          />
                        )}

                      </div>
                    </td>
                  </tr>
                ))}
              </>

              {/* Claim Type */}
              <tr>
                <td className={styles.labelCell}>
                  Claim Type :
                  <span className={styles.required}>*</span>
                </td>

                <td className={styles.fieldCell}>
                  <Controller
                    name="claimType"
                    control={control}
                    rules={{
                      required: "Claim Type is required"
                    }}
                    render={({ field }) => (
                      <Dropdown
                        placeholder="Select Claim Type"
                        options={claimTypeOptions}
                        disabled={isReadOnly}
                        selectedKey={claimTypeOptions.find(x => x.text === field.value)?.key}
                        onChange={(_, option) =>
                          field.onChange(option?.key)
                        }
                        errorMessage={errors.claimType?.message}
                        styles={dropdownStyles}
                      />
                    )}
                  />
                </td>
              </tr>

              {/* City */}
              <tr>
                <td className={styles.labelCell}>
                  City :
                </td>

                <td className={styles.fieldCell}>
                  <Controller
                    name="city"
                    control={control}
                    render={({ field }) => (
                      <TextField
                        readOnly={isReadOnly}
                        value={field.value}
                        onChange={(_, value) =>
                          field.onChange(value)
                        }
                        styles={textFieldStyles}
                      />
                    )}
                  />
                </td>
              </tr>

              {/* State */}
              <tr>
                <td className={styles.labelCell}>
                  State :
                </td>

                <td className={styles.fieldCell}>
                  <Controller
                    name="state"
                    control={control}
                    render={({ field }) => (
                      <TextField
                        readOnly={isReadOnly}
                        value={field.value}
                        onChange={(_, value) =>
                          field.onChange(value)
                        }
                        styles={textFieldStyles}
                      />
                    )}
                  />
                </td>
              </tr>

              {/* Country */}
              <tr>
                <td className={styles.labelCell}>
                  Country :
                  <span className={styles.required}>*</span>
                </td>

                <td className={styles.fieldCell}>
                  <Controller
                    name="country"
                    control={control}
                    rules={{
                      required: "Country is required"
                    }}
                    render={({ field }) => (
                      <Dropdown
                        placeholder='Select Country'
                        disabled={isReadOnly}
                        options={countryOptions}
                        selectedKey={countryOptions.find(x => x.text === field.value)?.key}
                        onChange={(_, option) =>
                          field.onChange(option?.key)
                        }
                        errorMessage={errors.country?.message}
                        styles={dropdownStyles}
                      />
                    )}
                  />
                </td>
              </tr>
              <tr>
                <td className={styles.labelCell}>
                  Site Type :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="siteType"
                    control={control}
                    rules={{
                      required: "Site Type is required"
                    }}
                    render={({ field }) => (
                      <Dropdown
                        placeholder='Select Site Type '
                        options={siteTypeOptions}
                        disabled={isReadOnly}
                        selectedKey={siteTypeOptions.find(x => x.text === field.value)?.key}
                        onChange={(_, option) =>
                          field.onChange(option?.key)
                        }
                        errorMessage={errors.siteType?.message}
                        styles={dropdownStyles}
                      />
                    )}
                  />
                </td>
              </tr>
              <tr>
                <td className={styles.labelCell}>
                  Site Activity :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="siteActivity"
                    control={control}
                    rules={{
                      required: "Site Activity is required"
                    }}
                    render={({ field }) => (
                      <Dropdown
                        placeholder="Select Site Activity"
                        disabled={isReadOnly}
                        options={siteActivityOptions}
                        selectedKey={siteActivityOptions.find(x => x.text === field.value)?.key}
                        onChange={(_, option) =>
                          field.onChange(option?.key)
                        }
                        errorMessage={errors.siteActivity?.message}
                        styles={dropdownStyles}
                      />
                    )}
                  />
                </td>
              </tr>
              <tr>
                <td className={styles.labelCell}>
                  Legacy Company :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="legacyCompany"
                    control={control}
                    rules={{
                      required: "Legacy Company is required"
                    }}
                    render={({ field }) => (
                      <Dropdown
                        placeholder="Select Legacy Company"
                        disabled={isReadOnly}
                        options={legacyCompaniesOptions}
                        selectedKey={legacyCompaniesOptions.find(x => x.text === field.value)?.key}
                        onChange={(_, option) =>
                          field.onChange(option?.key)
                        }
                        errorMessage={errors.legacyCompany?.message}
                        styles={dropdownStyles}
                      />
                    )}
                  />
                </td>
              </tr>
              {selectedClaimType === "Indemnification" && (
                <>
                  <tr>
                    <td className={styles.labelCell}>
                      Obligation Type :
                      <span className={styles.required}>*</span>
                    </td>

                    <td className={styles.fieldCell}>
                      <Controller
                        name="obligationType"
                        control={control}
                        rules={{
                          validate: value =>
                            selectedClaimType !== "Indemnification"
                              ? true
                              : !!value || "Obligation Type is required"
                        }}
                        render={({ field }) => (
                          <Dropdown
                            placeholder="Select Obligation Type"
                            options={obligationTypeOptions}
                            disabled={isReadOnly}
                            selectedKey={
                              obligationTypeOptions.find(
                                x => x.text === field.value
                              )?.key
                            }
                            onChange={(_, option) =>
                              field.onChange(option?.key)
                            }
                            errorMessage={
                              errors.obligationType?.message
                            }
                            styles={dropdownStyles}
                          />
                        )}
                      />
                    </td>
                  </tr>

                  <tr>
                    <td className={styles.labelCell}>
                      Operating Group :
                      <span className={styles.required}>*</span>
                    </td>

                    <td className={styles.fieldCell}>
                      <Controller
                        name="operatingGroup"
                        control={control}
                        rules={{
                          validate: value =>
                            selectedClaimType !== "Indemnification"
                              ? true
                              : !!value || "Operating Group is required"
                        }}
                        render={({ field }) => (
                          <Dropdown
                            placeholder="Select Operating Group"
                            options={operatingGroupOptions}
                            disabled={isReadOnly}
                            selectedKey={
                              operatingGroupOptions.find(
                                x => x.text === field.value
                              )?.key
                            }
                            onChange={(_, option) =>
                              field.onChange(option?.key)
                            }
                            errorMessage={
                              errors.operatingGroup?.message
                            }
                            styles={dropdownStyles}
                          />
                        )}
                      />
                    </td>
                  </tr>

                  <tr>
                    <td className={styles.labelCell}>
                      Transaction Date (MM-DD-YYYY) :
                      <span className={styles.required}>*</span>
                    </td>

                    <td className={styles.fieldCell}>
                      <Controller
                        name="transactionDate"
                        control={control}
                        rules={{
                          validate: value =>
                            selectedClaimType !== "Indemnification"
                              ? true
                              : !!value ||
                              "Transaction Date is required"
                        }}
                        render={({ field }) => (
                          <>
                            <DatePicker
                              placeholder="Select Transaction Date"
                              disabled={isReadOnly}
                              value={field.value}
                              onSelectDate={date =>
                                field.onChange(date)
                              }
                              formatDate={(date?: Date) =>
                                date
                                  ? moment(date).format(
                                    "MM-DD-YYYY"
                                  )
                                  : ""
                              }
                            />

                            {errors.transactionDate && (
                              <Text
                                styles={{
                                  root: {
                                    color: "#d13438",
                                    fontSize: 12,
                                    marginTop: 4
                                  }
                                }}
                              >
                                {
                                  errors.transactionDate
                                    .message
                                }
                              </Text>
                            )}
                          </>
                        )}
                      />
                    </td>
                  </tr>
                </>
              )}
              <tr>
                <td className={styles.labelCell}>
                  Date Site Added (MM-DD-YYYY) :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <TextField
                    readOnly
                    value={moment().format("MM-DD-YYYY")}
                    styles={textFieldStyles}
                  />
                </td>
              </tr>
              <tr>
                <td className={styles.labelCell}>
                  Cost Estimation Method :
                  <span className={styles.required}>*</span>
                </td>
                <td className={styles.fieldCell}>
                  <Controller
                    name="costEstimateMethod"
                    control={control}
                    rules={{
                      required: "Cost Estimation Method is required"
                    }}
                    render={({ field }) => (
                      <Dropdown
                        disabled={isReadOnly}
                        options={costEstimateMethodOptions}
                        placeholder="Select Cost Estimation Method"
                        selectedKey={costEstimateMethodOptions.find(x => x.text === field.value)?.key}
                        onChange={(_, option) =>
                          field.onChange(option?.key)
                        }
                        errorMessage={errors.costEstimateMethod?.message}
                        styles={dropdownStyles}
                      />
                    )}
                  />
                </td>
              </tr>

              {/* Original Company Connection */}
              <tr>
                <td className={styles.labelCell}>
                  Original Company Connection :
                  <span className={styles.required}>*</span>
                </td>

                <td className={styles.fieldCell}>
                  <Controller
                    name="originalCompanyConnection"
                    control={control}
                    rules={{
                      required: "Original Company Connection is required"
                    }}
                    render={({ field }) => (
                      <>
                        <TooltipHost
                          content="Includes site history up to the point it became recognized as a potential liability, including the nexus to this site nature of first notice of liability - when and how Pfizer was first notified of this potential liability (Limit 2000 characters)."
                          directionalHint={DirectionalHint.topCenter}
                          tooltipProps={{
                            styles: {
                              content: {
                                backgroundColor: "#323130",
                                color: "#ffffff",
                                border: "none"
                              }
                            }
                          }}
                        >
                          <TextField
                            readOnly={isReadOnly}
                            multiline
                            rows={5}
                            value={field.value}
                            onChange={(_, value) => field.onChange(value)}
                            errorMessage={
                              errors.originalCompanyConnection?.message
                            }
                            styles={multiLineFieldStyles}
                          />
                        </TooltipHost>
                      </>
                    )}
                  />

                </td>
              </tr>

            </tbody>
          </table>

          <div className={styles.actionBar}>
            <Stack
              horizontal
              wrap
              horizontalAlign="center"
              tokens={{ childrenGap: 12 }}
            >
              {pageMode === "create" && (
                <PrimaryButton
                  text={
                    isSubmitting
                      ? "Submitting..."
                      : isSubmitted
                        ? "Submitted"
                        : "Submit"
                  }
                  disabled={isSubmitting || isSubmitted}
                  iconProps={{
                    iconName: isSubmitted
                      ? "Completed"
                      : "CheckMark"
                  }}
                  onClick={handleCreateClick}
                // onClick={handleSubmit(onSubmit)}
                />
              )}
              {pageMode === "approval" && !isProcessed && (
                <>

                  <PrimaryButton
                    text="Approve"
                    iconProps={{
                      iconName: "Completed"
                    }}
                    onClick={handleApproveClick}
                  //  onClick={handleSubmit(approveRequest)}
                  />
                  <PrimaryButton
                    text="Reject"
                    iconProps={{
                      iconName: "Blocked"
                    }}
                    styles={{
                      root: {
                        backgroundColor: "#d13438",
                        borderColor: "#d13438"
                      },
                      rootHovered: {
                        backgroundColor: "#e47373",
                        borderColor: "#a4262c"
                      }
                    }}
                    onClick={() => void rejectRequest()}
                  />
                </>
              )}

              <DefaultButton
                text="Cancel"
                iconProps={{
                  iconName: "Cancel"
                }}
                onClick={() => {
                  if (pageMode === "create") {
                    navigateToHome();
                  } else {
                    navigateToDashboard();
                  }
                }}
              />


            </Stack>
          </div>
        </div>
      </section>
    </>
  );
}

