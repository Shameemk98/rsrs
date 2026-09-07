import * as React from 'react';
import styles from './SiteOverview.module.scss';
import type { ISiteOverviewProps } from './ISiteOverviewProps';
import { ComboBox, DatePicker, DefaultButton, Dropdown, Icon, IconButton, IDropdownOption, Link, PrimaryButton, Spinner, SpinnerSize, Stack, TextField } from '@fluentui/react';
import { Controller, useForm } from 'react-hook-form';
import { PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { SpService } from '../../../services/spService';
import { useEffect, useState } from 'react';
import moment from 'moment';
import {  Home32Regular } from "@fluentui/react-icons";
import { MessageBar, MessageBarType } from "@fluentui/react";
import RsrsHeader from '../../rsrsHeader/components/RsrsHeader';
import {
  Dialog,
  DialogType,
  DialogFooter
} from "@fluentui/react";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";


interface IFormData {
  RSRSSiteID: any;
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
  dateSiteAdded: string;
}

interface ISupportRow {
  id: number;
  support?: string;
  role?: string;
}

export default function SiteOverview(props: ISiteOverviewProps): React.ReactElement<ISiteOverviewProps> {
  const [supportRows, setSupportRows] = useState<ISupportRow[]>([{ id: Date.now(), support: "", role: "" }]);
  const [contactsOptions, setContactsOptions] = useState<IDropdownOption[]>([]);
  const [operatingGroupOptions, setOperatingGroupOptions] = useState<IDropdownOption[]>([]);
  const [countryOptions, setCountryOptions] = useState<IDropdownOption[]>([]);
  const [siteTypeOptions, setSiteTypeOptions] = useState<IDropdownOption[]>([]);
  const [legacyCompaniesOptions, setLegacyCompaniesOptions] = useState<IDropdownOption[]>([]);
  const [jumpSiteOptions, setJumpSiteOptions] = useState<IDropdownOption[]>([]);
  const [selectedSiteId, setSelectedSiteId] = useState<number>();
  const [siteNumber, setSiteNumber] = useState<string>("");
  const [selectedView, setSelectedView] = useState("all");
  const [siteLeadUsers, setSiteLeadUsers] = useState<any[]>([]);
  const [siteSupportUsers, setSiteSupportUsers] = useState<any[]>([]);
  const [siteUpdateOwnerUsers, setSiteUpdateOwnerUsers] = useState<any[]>([]);
  const [isAdmin, setIsAdmin] = useState<boolean>(false);
  //const [securityLevel, setSecurityLevel] = useState<string>("");
  // const [assignedSites, setAssignedSites] = useState<number[]>([]);
  const [currentUserId, setCurrentUserId] = useState<number>(0);
  const [message, setMessage] = useState("");
  const [messageType, setMessageType] = useState<MessageBarType>(MessageBarType.success);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [isSiteLoading, setIsSiteLoading] =  useState<boolean>(false);
  const [showDialog, setShowDialog] = useState(false);
const [isErrorDialog, setIsErrorDialog] = useState(false);
const [redirectUrl, setRedirectUrl] = useState("");


  const formTopRef = React.useRef<HTMLDivElement>(null);
  const spService = new SpService(props.context);
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

  const removeSupportRow = (
    id: number
  ): void => {

    setSupportRows(prev =>
      prev.filter(
        row => row.id !== id
      )
    );

  };
const parseHtmlToPlainText = (html: string): string => {
  if(html)
  {
  return html
      .replace(/<\/p><p>/gi, '\n')      // Turn separate paragraphs into line breaks
      .replace(/<br\s*\/?>/gi, '\n')    // Turn explicit <br> tags into line breaks
      .replace(/<\/?[^>]+(>|$)/g, ''); 
  }
  else return ""
};
const validateUpdateOwnerOrOutsideCounsel = (): boolean => {

  const hasSiteUpdateOwner =
    siteUpdateOwnerUsers?.length > 0;

  const hasOutsideCounsel =
    supportRows.some(
      row =>
        row.support &&
        row.role === "Outside Counsel"
    );

  return (
    hasSiteUpdateOwner ||
    hasOutsideCounsel
  );
};

  const onSubmit = async (data: IFormData): Promise<void> => {
    if (!validateUpdateOwnerOrOutsideCounsel()) {
  return;
}
    setIsSubmitting(true);

    try {

      if (!selectedSiteId) {

        formTopRef.current?.scrollIntoView({
          behavior: "smooth",
          block: "start"
        });

        setMessage("Please select a site.");
        setMessageType(MessageBarType.warning);
        return;
      }
      // const siteLeadUser = siteLeadUsers?.length > 0 ? await spService.ensureUser(siteLeadUsers[0].EMail || siteLeadUsers[0].secondaryText) : null;
      // const siteSupportUser = siteSupportUsers?.length > 0 ? await spService.ensureUser(siteSupportUsers[0].EMail || siteSupportUsers[0].secondaryText) : null;
      // const siteUpdateOwnerUser = siteUpdateOwnerUsers?.length > 0 ? await spService.ensureUser(siteUpdateOwnerUsers[0].EMail || siteUpdateOwnerUsers[0].secondaryText) : null;
      const getUserId = async (user: any): Promise<number | null> => {
        if (!user) return null;
        if (user.Id) {
          return user.Id;
        }

        // PeoplePicker user, resolve to SharePoint ID
        const ensuredUser = await spService.ensureUser(
          user.EMail ||
          user.secondaryText ||
          user.loginName
        );

        return ensuredUser?.Id?? null;
      };

      const siteLeadUserId =
        siteLeadUsers?.length > 0
          ? await getUserId(siteLeadUsers[0])
          : null;

      const siteSupportUserId =
        siteSupportUsers?.length > 0
          ? await getUserId(siteSupportUsers[0])
          : null;

      const siteUpdateOwnerUserId =
        siteUpdateOwnerUsers?.length > 0
          ? await getUserId(siteUpdateOwnerUsers[0])
          : null;
      const payload: any = {

        RSRSSiteName: data.siteName,
        Claim_Type: data.claimType,
        Country: countryOptions.find(x => x.key === data.country)?.text || "",
        City: data.city,
        State: data.state,
        Site_Type: siteTypeOptions.find(x => x.key === data.siteType)?.text || "",
        Site_Activity: data.siteActivity,
        Legacy_Company: legacyCompaniesOptions.find(x => x.key === data.legacyCompany)?.text || "",
        Cost_Estimate_Method: costEstimateMethodOptions.find(x => x.key === data.costEstimateMethod)?.text || "",
        OriginalCompanyConnection: data.originalCompanyConnection,
        // SiteLeadId: siteLeadUser?.Id ? [siteLeadUser.Id] : [],
        // SiteSupportId: siteSupportUser?.Id ? [siteSupportUser?.Id] : [],
        // SiteUpdateOwnerId: siteUpdateOwnerUser?.Id ? [siteUpdateOwnerUser?.Id] : [],
        SiteLeadId: siteLeadUserId?[siteLeadUserId]:[],
        SiteSupportId: siteSupportUserId?[siteSupportUserId]:[],
        SiteUpdateOwnerId: siteUpdateOwnerUserId?[siteUpdateOwnerUserId]:[],

      };
      console.log(payload, "payloadpayload");

      if (data.claimType === "Indemnification") {

        payload.ObligationType =
          data.obligationType;

        payload.OperatingGroup =
          operatingGroupOptions.find(
            x => x.key === data.operatingGroup
          )?.text || "";

        payload.TransactionDate =
          data.transactionDate;
      }

      await spService.updateItem(
        props.SiteListName,
        selectedSiteId,
        payload
      );
      setJumpSiteOptions(prev =>
  prev.map(item =>
    item.key === selectedSiteId
      ? {
          ...item,
          text: data.siteName
        }
      : item
  )
);
      const existingSupports =
        await spService.getItems(
          props.SiteAdditionalSupportListName,
          ["Id"],
          `SiteIDId eq ${selectedSiteId}`
        );

      for (const item of existingSupports) {

        await spService.deleteItem(
          props.SiteAdditionalSupportListName,
          item.Id
        );

      }

      for (const row of supportRows) {

        if (!row.support) continue;

        await spService.createItem(
          props.SiteAdditionalSupportListName,
          {
            SiteIDId: selectedSiteId,
            AdditionalContactId: row.support,
            AdditionalSupportRole: row.role
          }
        );
      }

      formTopRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });

      setMessage("Site updated successfully.");
setMessageType(MessageBarType.success);

setRedirectUrl(
  `${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`
);

setIsErrorDialog(false);
setShowDialog(true);


    }
    catch (error) {

      console.error(error);

      formTopRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });

      setMessage(
  "An error occurred while updating the site."
);

setMessageType(
  MessageBarType.error
);

setRedirectUrl("");
setIsErrorDialog(true);
setShowDialog(true);
    }
    finally {

      setIsSubmitting(false);

    }
  };

  const loadUserSecurity = async (): Promise<boolean> => {

    const currentUserEmail =
      props.context.pageContext.user.email;

    const users = await spService.getItems(
      props.ContactsListName,
      [
        "User_Security_Level",
        "ContactEmail"
      ],
      `ContactEmail eq '${currentUserEmail}'`
    );
    console.log(users,"usersusersusers");
    

    if (!users.length) {
      return false;
    }

    const admin =
      users[0].User_Security_Level ===
      "Administrator";

       console.log(admin,"adminadminadmin");

    setIsAdmin(admin);

    return admin;
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
          key: item.Id,
          text: item.Title
        }))
      ];
      setLegacyCompaniesOptions(options);
    } catch (error) {
      console.error("Error loading countries:", error);
    }
  };

  const loadSiteTypes = async (): Promise<void> => {
    try {

      const items = await spService.getItems(props.siteTypeListName, ["Id", "Title"]);
      const options: IDropdownOption[] = [
        {
          key: "",
          text: "Select Site Type"
        },
        ...items.map(item => ({
          key: item.Id,
          text: item.Title
        }))
      ];
      setSiteTypeOptions(options);
    } catch (error) {
      console.error("Error loading SiteTypes:", error);
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
          key: item.Id,
          text: item.Title
        }))
      ];
      setCountryOptions(options);
    } catch (error) {
      console.error("Error loading countries:", error);
    }
  }


  const loadOperatingGroups = async (): Promise<void> => {
    try {
      const items = await spService.getItems(props.OperatingGroupListName, ["Id", "Title"], undefined, { field: "Title", ascending: true });
      const options: IDropdownOption[] = [
        {
          key: "",
          text: "Select Operating Group"
        },
        ...items.map(item => ({
          key: item.Id,
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
          text: "Select Additional Support"
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
/*   const {
    control,
    handleSubmit,
    trigger,
    watch,
    reset,
    formState: { errors }
  } = useForm<IFormData>({}) */
   const {
  control,
  handleSubmit,
  trigger,
  watch,
  reset,
  formState: { errors }
} = useForm<IFormData>();

  const selectedClaimType = watch("claimType");
 
/*   const decodeHtml = (html: string): string => {
  const txt = document.createElement("textarea");
  txt.innerHTML = html;
  return txt.value;
  
  
};
console.log(decodeHtml,"decodeHtml"); */
const handleUpdateClick = async (): Promise<void> => {

  const formValid = await trigger();

  const updateOwnerValid =
    validateUpdateOwnerOrOutsideCounsel();

  if (!updateOwnerValid) {

    setMessage(
      "Please select either Site Update Owner or at least one Additional Support as Outside Counsel"
    );

    setMessageType(
      MessageBarType.error
    );

    setIsErrorDialog(true);

    setShowDialog(true);
  }

  if (!formValid || !updateOwnerValid) {
    return;
  }

  await handleSubmit(onSubmit)();
};

    const loadAdditionalSupport = async (
    requestId: number
  ): Promise<void> => {

    const items = await spService.getItems(
      props.SiteAdditionalSupportListName,
      [
        "Id",
        "AdditionalContactId",
        "AdditionalSupportRole"
      ],
      `SiteIDId eq ${requestId}`
    );


    if (items.length > 0) {

      setSupportRows(
        items.map((item: any) => ({
          id: item.Id,
          support: item.AdditionalContactId,
          role: item.AdditionalSupportRole
        }))
      );

    } else {

      setSupportRows([
        {
          id: Date.now(),
          support: "",
          role: ""
        }
      ]);

    }
  };
     const loadSiteById = async (siteId: number): Promise<void> => {
    setIsSiteLoading(true);

    try {

      const site =
        await spService.getListItemById(
          props.SiteListName,
          siteId
        );
      setSiteNumber(site.Title);
      setSiteLeadUsers(site.SiteLead || []);

      setSiteSupportUsers(
        site.SiteSupport || []
      );

      setSiteUpdateOwnerUsers(
        site.SiteUpdateOwner || []
      );
      const countryKey =
        countryOptions.find(
          x => x.text === site.Country
        )?.key;

      const siteTypKey =
        siteTypeOptions.find(
          x => x.text === site.Site_Type
        )?.key;
         console.log(siteTypKey, "siteTypKeysiteTypKey");

      const legacyCmpanyKey =
        legacyCompaniesOptions.find(
          x => x.text === site.Legacy_Company
        )?.key;

      const costEstimaMethd =
        costEstimateMethodOptions.find(
          x => x.text === site.Cost_Estimate_Method
        )?.key;

      const obligtionTyp =
        obligationTypeOptions.find(
          x => x.text === site.ObligationType)?.key;

      const opertnGrp =
        operatingGroupOptions.find(
          x => x.text === site.OperatingGroup)?.key;




      reset({
        RSRSSiteID: site.Title,
        siteName: site.RSRSSiteName,
        city:
          site.City,

        state:
          site.State,

        country: countryKey as string,


        claimType:
          site.Claim_Type,

        siteType:
          siteTypKey as string,

        siteActivity:
          site.Site_Activity,

        legacyCompany:
          legacyCmpanyKey as string,

        costEstimateMethod:
          costEstimaMethd as string,

        originalCompanyConnection: site.OriginalCompanyConnection,
       //originalCompanyConnection:decodeHtml(site.OriginalCompanyConnection || ""),
        dateSiteAdded: site.DateSiteAdded,

        obligationType: obligtionTyp as string,
        operatingGroup: opertnGrp as string,
        transactionDate: site.TransactionDate ? new Date(site.TransactionDate) : undefined,
        siteLead: site.SiteLead || []
      });

      await loadAdditionalSupport(siteId);

    } catch (error) {

      console.error(error);

    }
    finally {

        setIsSiteLoading(false);

    }

  };

/* const loadSites = async (
  filterType: string = "Open",
  admin?: boolean
): Promise<void> => {

  const sites = await spService.getItems(
    props.SiteListName,
    [
      "Id",
      "RSRSSiteName",
      "Claim_Type",
      "Site_Activity",
      "SiteLead/Id",
      "SiteLead/Title",
      //"SiteLead/EMail",
      "SiteSupport/Id",
      "SiteSupport/Title",
     // "SiteSupport/EMail",
      "SiteUpdateOwner/Id",
      "SiteUpdateOwner/Title",
      //"SiteUpdateOwner/EMail"
    ],
    undefined,
    { field: "RSRSSiteName", ascending: true },
    2000,
    ["SiteLead", "SiteSupport", "SiteUpdateOwner"]
  );
console.log(sites,"sitessitessites");

  const currentUserId =props.context.pageContext.legacyPageContext.userId;

  let filteredSites = [...sites] ;

  // Legacy user filter
  if (!admin) {

    filteredSites = filteredSites.filter((site: any) => {

      const isSiteLead =
        site.SiteLead?.some(
          (u: any) => u.Id === currentUserId
        );

      const isSiteSupport =
        site.SiteSupport?.some(
          (u: any) => u.Id === currentUserId
        );

      const isSiteUpdateOwner =
        site.SiteUpdateOwner?.some(
          (u: any) => u.Id === currentUserId
        );

      return (
        isSiteLead ||
        isSiteSupport ||
        isSiteUpdateOwner
      );
    });
  }
 

  // View Sites = My Open Sites
  if (filterType === "Open") {

    filteredSites = filteredSites.filter(
      (site: any) =>
        site.Site_Activity === "Open" ||
        site.Site_Activity === "Open - Monitor Only"
    );
  }

  // View Sites = My Update Sites
  else if (filterType === "Updated") {

    filteredSites = filteredSites.filter(
      (site: any) =>
        site.SiteUpdateOwner?.some(
          (u: any) => u.Id === currentUserId
        )
    );

    filteredSites = filteredSites.filter(
      (site: any) =>
        site.Site_Activity === "Open" ||
        site.Site_Activity === "Open - Monitor Only"
    );
  }

  // View Sites = All My Sites
  else if (filterType === "all") {
    // No additional filtering
    // Admin sees all
    // Non-admin already filtered above
  }
 console.log(filteredSites,"filteredSites----");
  setJumpSiteOptions(
    filteredSites.map((site: any) => ({
      key: site.Id,
      text: site.RSRSSiteName
    }))
  );
  const params = new URLSearchParams(window.location.search);
const sid = params.get("sid");

if (
  !sid &&
  filteredSites.length > 0 &&
  !selectedSiteId
) {
  const firstSiteId = filteredSites[0].Id;

  setSelectedSiteId(firstSiteId);
  void loadSiteById(firstSiteId);
}
}; */
const loadSites = async (
  filterType: string = "Open",
  admin?: boolean
): Promise<void> => {

  const sites = await spService.getItems(
    props.SiteListName,
    [
      "Id",
      "RSRSSiteName",
      "Claim_Type",
      "Site_Activity",
      "SiteLead/Id",
      "SiteLead/Title",
      "SiteSupport/Id",
      "SiteSupport/Title",
      "SiteUpdateOwner/Id",
      "SiteUpdateOwner/Title"
    ],
    undefined,
    { field: "RSRSSiteName", ascending: true },
    2000,
    ["SiteLead", "SiteSupport", "SiteUpdateOwner"]
  );

  const currentUserId =
    props.context.pageContext.legacyPageContext.userId;

  let filteredSites = [...sites];

  // Legacy user filter
  if (!admin) {
    filteredSites = filteredSites.filter((site: any) => {

      const isSiteLead =
        site.SiteLead?.some(
          (u: any) => u.Id === currentUserId
        );

      const isSiteSupport =
        site.SiteSupport?.some(
          (u: any) => u.Id === currentUserId
        );

      const isSiteUpdateOwner =
        site.SiteUpdateOwner?.some(
          (u: any) => u.Id === currentUserId
        );

      return (
        isSiteLead ||
        isSiteSupport ||
        isSiteUpdateOwner
      );
    });
  }

  // My Open Sites
  if (filterType === "Open") {

    filteredSites = filteredSites.filter(
      (site: any) =>
        site.Site_Activity === "Open" ||
        site.Site_Activity === "Open - Monitor Only"
    );

  }

  // My Update Sites
  else if (filterType === "Updated") {

    filteredSites = filteredSites.filter(
      (site: any) =>
        site.SiteUpdateOwner?.some(
          (u: any) => u.Id === currentUserId
        )
    );

    filteredSites = filteredSites.filter(
      (site: any) =>
        site.Site_Activity === "Open" ||
        site.Site_Activity === "Open - Monitor Only"
    );

  }

  const options = filteredSites.map((site: any) => ({
    key: site.Id,
    text: site.RSRSSiteName
  }));

  setJumpSiteOptions(options);

  const params = new URLSearchParams(
    window.location.search
  );

  const sid = params.get("sid");

  // If page opened from URL parameter
  if (sid) {
    return;
  }

  // Legacy behavior:
  // Keep current site if it still exists
  const currentSiteStillExists =
    selectedSiteId &&
    filteredSites.some(
      (site: any) =>
        site.Id === selectedSiteId
    );

  if (currentSiteStillExists) {

    void loadSiteById(selectedSiteId!);

  } else if (filteredSites.length > 0) {

    // Select first site automatically
    const firstSiteId = filteredSites[0].Id;

    setSelectedSiteId(firstSiteId);

    // Load first site's details
    void loadSiteById(firstSiteId);

  } else {

    // No sites after filter
    setSelectedSiteId(undefined);

    reset();
  }
};






  // const initialize = async (): Promise<void> => {

  //   const admin = await loadUserSecurity();

  //   await Promise.all([
  //     loadCountries(),
  //     loadSiteTypes(),
  //     loadLegacyCompanies(),
  //     loadOperatingGroups(),
  //     loadContacts()
  //   ]);

  //   await loadSites("all", admin);

  // };
  const initialize = async (): Promise<void> => {
    const admin = await loadUserSecurity();

    await Promise.all([
      loadCountries(),
      loadSiteTypes(),
      loadLegacyCompanies(),
      loadOperatingGroups(),
      loadContacts()
    ]);

    setIsAdmin(admin);
  };
  useEffect(() => {
    if (
      countryOptions.length > 0 &&
      siteTypeOptions.length > 0 &&
      legacyCompaniesOptions.length > 0 &&
      operatingGroupOptions.length > 0&&
      contactsOptions.length>0
    ) {
      void loadSites("all", isAdmin);
    }
  }, [
    countryOptions,
    siteTypeOptions,
    legacyCompaniesOptions,
    operatingGroupOptions,
    contactsOptions,
    isAdmin
  ]);
  useEffect(() => {
    void initialize();
    const currentUserEmail = props.context.pageContext.user.email;
    setCurrentUserId(props.context.pageContext.legacyPageContext.userId);

    console.log(props.context.pageContext.legacyPageContext.userId, "props.context.pageContext.legacyPageContext.userId");

  }, []);
  useEffect(() => {
  const params = new URLSearchParams(window.location.search);

  const sid = params.get("sid");
  const claimType = params.get("Claim_type");

  if (sid) {
    setSelectedSiteId(Number(sid));
  }

  console.log("sid:", sid);
  console.log("Claim Type:", claimType);
}, []);
useEffect(() => {
  const params = new URLSearchParams(window.location.search);

  const sid = params.get("sid");

  if (sid && jumpSiteOptions.length > 0) {
    const siteId = Number(sid);

    setSelectedSiteId(siteId);

    void loadSiteById(siteId);
  }
}, [jumpSiteOptions]);


  const peoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory,
    spHttpClient: props.context.spHttpClient
  } as any;

  const viewSitesOptions: IDropdownOption[] = [
    { key: "Open", text: "My Open Sites" },
    { key: "all", text: "All My Sites" },
    { key: "Updated", text: "My Update Sites" }
  ];
  /* 
 
   const jumpSiteOptions: IDropdownOption[] = [
   { key: "Pfizer", text: "Pfizer" },
   { key: "Infosys", text: "Infosys" }
 ]; */

  const siteActivityOptions: IDropdownOption[] = [
    { key: "", text: "Select Site Activity" },
    { key: 'Open', text: 'Open' },
    { key: 'Open - Monitor Only', text: 'Open - Monitor Only' },
    { key: 'Closed', text: 'Closed' }
  ];

  const claimTypeOptions: IDropdownOption[] = [
    { key: "", text: "Select Claim Type" },
    { key: 'Indemnification', text: 'Indemnification' },
    { key: 'Remediation', text: 'Remediation' },
    { key: 'ToxicTort', text: 'Toxic Tort' }
  ];

  const roleOptions: IDropdownOption[] = [
    { key: "", text: "Select A Role" },
    { key: "Assistant Project Manager", text: "Assistant Project Manager" },
    { key: "Outside Counsel", text: "Outside Counsel" },
    { key: "Technical Consultant", text: "Technical Consultant" },
    { key: "Other", text: "Other" }
  ];



  const cardStyle = {
    //background: '#ffffff',
    borderRadius: '16px',
    padding: '0 3rem',
    width: '100%',
    marginBottom: '20px',
    boxSizing: 'border-box' as const,
    borderBottom: '4px solid #135891',
    boxShadow: '0 2px 8px rgba(0,0,0,0.08)'
  };

  return (
    <>
    {/* <RsrsHeader context={props.context}  /> */}
    <Dialog
  hidden={!showDialog}
  dialogContentProps={{
    type: DialogType.largeHeader,
    title: isErrorDialog ? "Error" : "Success",
    subText:message
  }}
>
  {/* <MessageBar
    messageBarType={messageType}
    isMultiline={true}
  >
    {message}
  </MessageBar> */}

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
    <section className={styles.siteOverview} >

      {(isSubmitting || isSiteLoading) && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: "rgba(255,255,255,0.7)",
            zIndex: 9999,
            display: "flex",
            alignItems: "center",
            justifyContent: "center"
          }}
        >
          <Spinner
            label="Please Wait..."
            size={SpinnerSize.large}
          />
        </div>
      )}
{/*       {message && (
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
          Site Overview
        </span>
      </div>
      <div ref={formTopRef} />
      <div
        style={{
          background: "#fff",
          //border: "1px solid #e1dfdd",
          //borderRadius: "4px",
          padding: "0 10rem",
          marginTop:'10px'
          //maxWidth: "1100px",
          //margin: "0 auto"
        }}
      >
        {/* View Sites */}
        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            View Sites :<span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Dropdown
              options={viewSitesOptions}
              selectedKey={selectedView}
              onChange={(_, option) => {
                const value = option?.key as string;
                setSelectedView(value);
                void loadSites(value, isAdmin);

              }}

            />
          </div>
        </div>

        {/* Jump To Site */}

        <div style={{ display: "flex", marginBottom: 40 }}>
          <div className={styles.label}>
            Jump To Site :<span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
              <ComboBox
                options={jumpSiteOptions}
                selectedKey={selectedSiteId}
                onChange={(_, option) => {
                  const siteId = option?.key as number;
                  setSelectedSiteId(siteId);
                  void loadSiteById(siteId);
                }}
              />

          </div>
        </div>

        {/* Site ID */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            RSRS Site ID :<span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <TextField
              value={siteNumber}
              readOnly
              disabled={!!selectedSiteId}
            />
          </div>
        </div>

        {/* Site Name */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            RSRS Site Name : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="siteName"
              control={control}
              rules={{
                required: "RSRS Site Name is required"
              }}
              render={({ field }) => (
                <TextField
                  value={field.value}
                  errorMessage={errors.siteName?.message}
                  onChange={(_, value) =>
                    field.onChange(value)
                  }
                />
              )}
            />
          </div>
        </div>

        {/* Site Lead */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Site Lead : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="siteLead"
              control={control}
              rules={{
                validate: (value) =>
                  (value && value.length > 0) ||
                  "Site Lead is required"
              }}
              render={({ field }) => (
                <>
                  <PeoplePicker
                    context={peoplePickerContext}
                    personSelectionLimit={1}
                     defaultSelectedUsers={
                      siteLeadUsers.map(
                        (u: any) => u.secondaryText?u.secondaryText: u.Title
                      )
                    } 
                    onChange={(items) => {
                      setSiteLeadUsers(items);
                      console.log(items, "itemsitems");

                      field.onChange(items);
                    }}
                  />

                  {errors.siteLead && (
                    <div
                      style={{
                        color: "#a1272b",
                        fontSize: 12,
                        marginTop: 4
                      }}
                    >
                      {errors.siteLead.message}
                    </div>
                  )}
                </>
              )}
            />
          </div>
        </div>

        {/* Site Support */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Site Support :
          </div>

          <div style={{ flex: 1 }}>
            <PeoplePicker
              context={peoplePickerContext}
              personSelectionLimit={1}
              placeholder='Enter a Name Or Email'
              defaultSelectedUsers={siteSupportUsers.map((u: any) =>u.secondaryText?u.secondaryText: u.Title)}
              onChange={(items) => {
                setSiteSupportUsers(items);
                console.log(items,"setSiteSupportUsers");

                //field.onChange(items);
              }}
            />
          </div>
        </div>

        {/* Update Owner */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            RSRS Update Owner :
          </div>

          <div style={{ flex: 1 }}>
            <PeoplePicker
              context={peoplePickerContext}
              personSelectionLimit={1}
              placeholder='Enter a Name Or Email'
              defaultSelectedUsers={siteUpdateOwnerUsers.map((u: any) => u.secondaryText?u.secondaryText: u.Title)}
              //defaultSelectedUsers={siteUpdateOwnerUsers}
              onChange={(items) => {
                setSiteUpdateOwnerUsers(items);
                //field.onChange(items);
              }}
            />
          </div>
        </div>

        {/* Additional Support */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Additional Support :
          </div>

          <div style={{ flex: 1 }}>

            {supportRows.map((row, index) => (

              <Stack
                key={row.id}
                horizontal
                tokens={{ childrenGap: 10 }}
                style={{ marginBottom: 10 }}
              >

                <Dropdown
                  styles={{ root: { flex: 1 } }}
                  options={contactsOptions}
                  selectedKey={row.support}
                  placeholder="Select Name"
                  onChange={(_, option) =>
                    updateSupportRow(
                      row.id,
                      "support",
                      option?.key
                    )
                  }
                />

                <Dropdown
                  styles={{ root: { width: 200 } }}
                  options={roleOptions}
                  selectedKey={row.role}
                  placeholder="Select Role"
                  onChange={(_, option) =>
                    updateSupportRow(
                      row.id,
                      "role",
                      option?.key
                    )
                  }
                />

                <IconButton
                  iconProps={{ iconName: "Add" }}
                  onClick={addSupportRow}
                />

                {supportRows.length > 1 && (

                  <IconButton
                    iconProps={{
                      iconName: "Delete"
                    }}
                    onClick={() =>
                      removeSupportRow(row.id)
                    }
                  />

                )}

              </Stack>

            ))}

          </div>
        </div>

        {/* Claim Type */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Claim Type : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="claimType"
              control={control}
              rules={{
                required: "Claim Type is required"
              }}
              render={({ field }) => (
                <Dropdown
                  options={claimTypeOptions}
                  placeholder="Select Claim Type"
                  selectedKey={field.value}
                  errorMessage={errors.claimType?.message}
                  onChange={(_, option) =>
                    field.onChange(option?.key)
                  }
                />
              )}
            />
          </div>
        </div>

        {/* City */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            City :
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="city"
              control={control}
              render={({ field }) => (
                <TextField
                  value={field.value || ""}
                  onChange={(_, value) =>
                    field.onChange(value)
                  }
                />
              )}
            />
          </div>
        </div>

        {/* State  */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            State :
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="state"
              control={control}
              render={({ field }) => (
                <TextField
                  value={field.value || ""}
                  onChange={(_, value) =>
                    field.onChange(value)
                  }
                />
              )}
            />
          </div>
        </div>

        {/* Country */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Country :<span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="country"
              control={control}
              rules={{
                required: "Country is required"
              }}
              render={({ field }) => (
                <Dropdown
                  options={countryOptions}
                  placeholder="Select Country"
                  selectedKey={field.value}
                  errorMessage={errors.country?.message}
                  onChange={(_, option) =>
                    field.onChange(option?.key)
                  }
                />
              )}
            />

          </div>
        </div>

        {/* Site Type */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Site Type : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="siteType"
              control={control}
              rules={{
                required: "Site Type is required"
              }}
              render={({ field }) => (
                <Dropdown
                  options={siteTypeOptions}
                  placeholder="Select Site Type"
                  errorMessage={errors.siteType?.message}
                  selectedKey={field.value}
                  onChange={(_, option) =>
                    field.onChange(option?.key)
                  }
                />
              )}
            />

          </div>
        </div>

        {/* Site Activity */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Site Activity  : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="siteActivity"
              control={control}
              rules={{
                required: "Site Activity is required"
              }}
              render={({ field }) => (
                <Dropdown
                  options={siteActivityOptions}
                  placeholder="Select Site Activity"
                  errorMessage={errors.siteActivity?.message}
                  selectedKey={field.value}
                  onChange={(_, option) =>
                    field.onChange(option?.key)
                  }
                />
              )}
            />

          </div>
        </div>

        {/* Legacy Company */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Legacy Company : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="legacyCompany"
              control={control}
              rules={{
                required: "Legacy Company is required"
              }}
              render={({ field }) => (
                <Dropdown
                  options={legacyCompaniesOptions}
                  placeholder="Select Legacy Company"
                  errorMessage={errors.legacyCompany?.message}
                  selectedKey={field.value}
                  onChange={(_, option) =>
                    field.onChange(option?.key)
                  }
                />
              )}
            />

          </div>
        </div>
        {selectedClaimType === "Indemnification" && (
          <>
            {/* Obligation Type */}
            <div style={{ display: "flex", marginBottom: 16 }}>
              <div className={styles.label}>
                Obligation Type :
                <span style={{ color: "red" }}>*</span>
              </div>

              <div style={{ flex: 1 }}>
                <Controller
                  name="obligationType"
                  control={control}
                  rules={{
                    validate: value =>
                      selectedClaimType !== "Indemnification"
                        ? true
                        : !!value ||
                        "Obligation Type is required"
                  }}
                  render={({ field }) => (
                    <Dropdown
                      options={obligationTypeOptions}
                      selectedKey={field.value}
                      onChange={(_, option) =>
                        field.onChange(option?.key)
                      }
                      errorMessage={
                        errors.obligationType?.message
                      }
                    />
                  )}
                />
              </div>
            </div>

            {/* Operating Group */}
            <div style={{ display: "flex", marginBottom: 16 }}>
              <div className={styles.label}>
                Operating Group :
                <span style={{ color: "red" }}>*</span>
              </div>

              <div style={{ flex: 1 }}>
                <Controller
                  name="operatingGroup"
                  control={control}
                  rules={{
                    validate: value =>
                      selectedClaimType !== "Indemnification"
                        ? true
                        : !!value ||
                        "Operating Group is required"
                  }}
                  render={({ field }) => (
                    <Dropdown
                      options={operatingGroupOptions}
                      selectedKey={field.value}
                      onChange={(_, option) =>
                        field.onChange(option?.key)
                      }
                      errorMessage={
                        errors.operatingGroup?.message
                      }
                    />
                  )}
                />
              </div>
            </div>

            {/* Transaction Date */}
            <div style={{ display: "flex", marginBottom: 16 }}>
              <div className={styles.label}>
                Transaction Date :
                <span style={{ color: "red" }}>*</span>
              </div>

              <div style={{ flex: 1 }}>
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
                        value={field.value}
                        formatDate={(date?: Date) => date ? moment(date).format("MM-DD-YYYY") : ""}
                        onSelectDate={(date) =>
                          field.onChange(date)
                        }
                      />
                      {errors.transactionDate && (
                        <div
                          style={{
                            color: "red",
                            fontSize: 12,
                            marginTop: 4
                          }}
                        >
                          {errors.transactionDate.message}
                        </div>
                      )}
                    </>
                  )}
                />
              </div>
            </div>
          </>
        )}

        {/* Date Site Added */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label} >
            Date Site Added (MM-DD-YYYY) : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="dateSiteAdded"
              control={control}
              render={({ field }) => (
                <TextField
                  value={field.value ? moment(field.value).format("MM-DD-YYYY") : ""}
                  readOnly
                  disabled={!!selectedSiteId}
                />
              )}
            />
          </div>
        </div>

        {/* Cost Estimation Method */}

        <div style={{ display: "flex", marginBottom: 16 }}>
          <div className={styles.label}>
            Cost Estimation Method : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="costEstimateMethod"
              control={control}
              rules={{
                required: "Cost Estimation Method is required"
              }}

              render={({ field }) => (
                <Dropdown
                  options={costEstimateMethodOptions}
                  placeholder="Select Cost Estimation Method"
                  errorMessage={errors.costEstimateMethod?.message}
                  selectedKey={field.value}
                  onChange={(_, option) =>
                    field.onChange(option?.key)
                  }
                />
              )}
            />
          </div>
        </div>

        {/* Original Company Connection */}

        <div style={{ display: "flex", marginBottom: 20 }}>
          <div className={styles.label}>
            Original Company Connection : <span style={{ color: "red" }}>*</span>
          </div>

          <div style={{ flex: 1 }}>
            <Controller
              name="originalCompanyConnection"
              control={control}
              rules={{
                required: "Original Company Connection is required"
              }}

              render={({ field }) => (
                <TextField
                  multiline
                  rows={6}
                  value={parseHtmlToPlainText(field.value)}
                  errorMessage={errors.originalCompanyConnection?.message}
                  onChange={(_, value) =>
                    field.onChange(value)
                  }
                />
              )}
            />
           {/*  <Controller
  name="originalCompanyConnection"
  control={control}
  rules={{
    required: "Original Company Connection is required"
  }}
  render={({ field }) => (
    <>
      <RichText
        value={field.value || ""}
        onChange={(value) => {
          field.onChange(value);
          return value;
        }}
      />

      {errors.originalCompanyConnection && (
        <div
          style={{
            color: "#a4262c",
            marginTop: 5
          }}
        >
          {errors.originalCompanyConnection.message}
        </div>
      )}
    </>
  )}
/> */}

          </div>
        </div>

        {/* Buttons */}

        <div
          style={{
            display: "flex",
            justifyContent: 'flex-end',
            gap: 10,
            marginTop: 24
          }}
        >

          
          <PrimaryButton
            text="Update"
            onClick={handleUpdateClick}
            //onClick={handleSubmit(onSubmit)}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => {
              window.location.href =
                `${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`;
            }}
          />
        </div>
      </div>
    </section>
    </>
  );

}

