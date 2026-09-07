import * as React from 'react';
import styles from './SummaryOfReserveMatters.module.scss';
import type { ISummaryOfReserveMattersProps } from './ISummaryOfReserveMattersProps';
import { useEffect, useState, useMemo } from 'react';
import {
  DetailsList,
  IColumn,
  SelectionMode,
  Dropdown,
  IDropdownOption,
  Link,
  SearchBox,
  mergeStyles,
  DetailsListLayoutMode,
  PrimaryButton,
  Icon,
  ConstrainMode
} from '@fluentui/react';
import { Home32Regular } from '@fluentui/react-icons';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import * as XLSX from 'xlsx-js-style';
import RsrsHeader from '../../rsrsHeader/components/RsrsHeader';

export default function SummaryOfReserveMatters(props: ISummaryOfReserveMattersProps): React.ReactElement<ISummaryOfReserveMattersProps> {
  const sp = useMemo(() => spfi().using(SPFx(props.context)), []);

  const [items, setItems] = useState<any[]>([]);
  const [filteredItems, setFilteredItems] = useState<any[]>([]);
  const [searchText, setSearchText] = useState('');
  const [activityFilter, setActivityFilter] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [sortKey, setSortKey] = useState<string>('');
  const [isSortedDescending, setIsSortedDescending] = useState<boolean>(false);

  const itemsPerPage = 10;
  const loadData = async () => {

    // ✅ Fetch Site list
    const sites = await sp.web.lists
      .getByTitle(props.sitelistname)
      .items.select(
        "ID", "Title", "RSRSSiteName",
        "SiteLead/Title", "SiteSupport/Title", "SiteUpdateOwner/Title",
        "City", "State", "Country", "Legacy_Company",
        "AccountCodeBlock", "Site_Activity", "Site_Type", "Claim_Type",
        "TotalProjectCost", "TotalSpendingToDate", "DateSiteAdded", "QuarterSiteClosedRSRS"
      )
      .expand("SiteLead", "SiteSupport", "SiteUpdateOwner")
      .top(4999)();



   
    const updates = await sp.web.lists
      .getByTitle("SiteUpdate")
      .items.select(
        "SiteName/Title",   // lookup to Site list
        "BETFlag",
        "TotalBETCost",
        "BETDiffrence",
        "BETDiffrencePercentage"
      )
      .expand("SiteName")
      .top(4999)();

    // ✅ Convert Siteupdate to lookup map for fast matching
    const updateMap: any = {};
    updates.forEach((u: any) => {
      const key = u.SiteName?.Title;
      if (key) {
        updateMap[key] = u;
      }
    });

    // ✅ Merge both datasets
    const finalData = sites.map((item: any) => {

      const update = updateMap[item.Title];



      const total = item.TotalProjectCost || 0;
      const spent = item.TotalSpendingToDate || 0;
      const rb = total - spent;


      return {
        ID: item.ID,
        id: item.Title,
        siteName: item.RSRSSiteName,
        siteLead: item.SiteLead?.[0]?.Title || "",
        siteSupport: item.SiteSupport?.[0]?.Title || "",
        rsrsUpdateOwner: item.SiteUpdateOwner?.[0]?.Title || "",
        city: item.City,
        state: item.State,
        country: item.Country,
        legacyCompany: item.Legacy_Company,
        accountCode: item.AccountCodeBlock,
        siteactivity: item.Site_Activity,
        siteType: item.Site_Type,
        claimType: item.Claim_Type,
        totalProjectCost: total === 0 ? 0 : total.toFixed(2),
        totalspendingtodate: spent === 0 ? 0 : spent.toFixed(2),
        reservebalance: rb === 0 ? 0 : rb.toFixed(2),
        betflag: update?.BETFlag === "true" ? "Yes" : update?.BETFlag === "false" ? "No" : "",
        totalbetcost: update?.TotalBETCost || 0,
        betdifference: update?.BETDiffrence || 0,
        betdifferencepercent: update?.BETDiffrencePercentage || 0,
        datesiteadded: item.DateSiteAdded ? item.DateSiteAdded.split("T")[0] : "",
        quarterclosed: item.QuarterSiteClosedRSRS
      };
    });
    console.log(finalData, "=========");

    setItems(finalData);
  };

  useEffect(() => {
    void loadData();
  }, [])

  const sortData = (data: any[], key: string, desc: boolean) => {
    return [...data].sort((a, b) => {
      let valA = a[key];
      let valB = b[key];

      if (valA == null) return 1;
      if (valB == null) return -1;

      if (typeof valA === 'string') {
        return desc
          ? valB.localeCompare(valA)
          : valA.localeCompare(valB);
      }

      return desc ? valB - valA : valA - valB;
    });
  };

  const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
    const newDesc = sortKey === column.fieldName ? !isSortedDescending : false;

    setSortKey(column.fieldName!);
    setIsSortedDescending(newDesc);

    const sorted = sortData(filteredItems, column.fieldName!, newDesc);
    setFilteredItems(sorted);
  };

  useEffect(() => {

    if (!activityFilter) {
      setFilteredItems([]);
      return;
    }

    let temp = [...items];

    if (activityFilter === 'OpenSites') {
      temp = temp.filter(x => x.siteactivity === 'Open' || x.siteactivity === 'Open - Monitor Only');
    }

    if (searchText) {
      const search = searchText.toLowerCase();

      temp = temp.filter(x =>
        Object.keys(x).some(key => {
          const val = (x as any)[key];
          return val && String(val).toLowerCase().includes(search);
        })
      );
    }

    if (sortKey) {
      temp = sortData(temp, sortKey, isSortedDescending);
    }

    setFilteredItems(temp);
    setCurrentPage(1);

  }, [activityFilter, searchText, items, sortKey, isSortedDescending]);

  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const paginatedItems = filteredItems.slice(startIndex, endIndex);

  const baseColumns: IColumn[] = [
    createColumn('id', 'RSRS Site ID',90,100),
    createColumn('siteName', 'Site Name',120,150),
    createColumn('siteLead', 'Site Lead',120,150),
    createColumn('siteSupport', 'Site Support',120,150),
    createColumn('rsrsUpdateOwner', 'RSRS Update Owner',140,150),
    createColumn('city', 'City',120,150),
    createColumn('state', 'State',120,150),
    createColumn('country', 'Country',120,150),
    createColumn('legacyCompany', 'Legacy Company',120,150),
    createColumn('accountCode', 'Account Code',120,150),
    createColumn('siteactivity', 'Site Activity',120,150),
    createColumn('siteType', 'Site Type',120,150),
    createColumn('claimType', 'Claim Type',120,150),
    createColumn('totalProjectCost', 'Total Project Cost', 120,150,true,),
    createColumn('totalspendingtodate', 'Total Spending to Date', 150,180,true,),
    createColumn('reservebalance', 'Reserve Balance', 120,150,true,)
  ];

  const openSiteExtraColumns: IColumn[] = [
    createColumn('betflag', 'BET Flag',120,150,true,),
    createColumn('totalbetcost', 'BET Estimated Cost',150,150,true),
    createColumn('betdifference', 'Difference', 80,150,true),
    createColumn('betdifferencepercent', '% Difference', 80,150,true)
  ];

  const tailColumns: IColumn[] = [
    createColumn('datesiteadded', 'Date Site Added',100,150,true),
    createColumn('quarterclosed', 'Quarter Closed',150,150,true)
  ];

  // ✅ FINAL columns
  const columns: IColumn[] =
    activityFilter === 'OpenSites'
      ? [...baseColumns, ...openSiteExtraColumns, ...tailColumns] : [...baseColumns, ...tailColumns];

  function createColumn(field: string, name: string, minWidth:number, maxWidth:number,isNumber?: boolean,): IColumn {
    return {
      key: field,
      name: name,
      fieldName: field,
      minWidth: minWidth,
      maxWidth:maxWidth,

      isSorted: sortKey === field,
      isResizable: true,
      isSortedDescending: sortKey === field ? isSortedDescending : false,
      onColumnClick,

      // ✅ Custom render for Site Name
      onRender: (item: any) => {
        if (field === "id") {
          console.log(item, "itemitemitemitem");

          const url =
            props.context.pageContext.web.absoluteUrl +
            "/SitePages/SiteOverview.aspx?sid=" +
            item.ID +
            "&Claim_type=" +
            (item.claimType ? item.claimType.split(' ')[0] : "");
          return (
            <span
              style={{ color: "#0078d4", cursor: "pointer" }}
              onClick={() => window.open(url, "_blank")}> {item.id}
            </span>
          );
        }

        return item[field];
      }
    };
  }
  const exportToExcel = () => {

    if (!filteredItems.length) {
      alert("No data to export");
      return;
    }

    // ✅ Prepare data
    const exportData = filteredItems.map(item => {
      const row: any = {};
      columns.forEach(col => {
        row[col.name] = item[col.fieldName as string];
      });
      return row;
    });

    // ✅ Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);

    // ✅ Apply header styling (row 1)
    const range = XLSX.utils.decode_range(worksheet['!ref']!);

    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });

      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: {
            bold: true,
            color: { rgb: "000000" } // black
          }
        };
      }
    }

    // ✅ Auto width
    worksheet['!cols'] = columns.map(() => ({ wch: 20 }));

    // ✅ Workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Reserve Matters");

    // ✅ Dynamic file name
    let fileName = "Summary_of_Reserve_Matters.xlsx";

    if (activityFilter === "AllSites") {
      fileName = "Summary Of All Reserve Matters.xlsx";
    } else if (activityFilter === "OpenSites") {
      fileName = "Summary Of Open Reserve Matters.xlsx";
    }

    // ✅ Download
    XLSX.writeFile(workbook, fileName);
  };
  ``
  const dropdownOptions: IDropdownOption[] = [
    { key: '', text: '--Select--' },
    { key: 'AllSites', text: 'All Sites' },
    { key: 'OpenSites', text: 'Open Sites' }
  ];

const headerClass = mergeStyles({
  backgroundColor: '#0078d4',
  selectors: {
    '.ms-DetailsHeader-cell': {
      backgroundColor: '#0078d4'
    },

    '.ms-DetailsHeader-cell:hover': {
      backgroundColor: '#0a4a7e !important'
    },

    '.ms-DetailsHeader-cell.is-actionable:hover': {
      backgroundColor: '#0a4a7e !important'
    },

    '.ms-DetailsHeader-cellTitle': {
      color: 'white',
      fontWeight: 600
    },

    '.ms-DetailsHeader-cell:hover .ms-DetailsHeader-cellTitle': {
      color: 'black !important'
    },

    '.ms-DetailsHeader-cell:hover .ms-DetailsHeader-cellName': {
      color: 'black !important'
    }
  }
});

  return (
    <>
      <RsrsHeader context={props.context} />
      <section className={styles.summaryOfReserveMatters}>
        <div className={styles.header}>
          <Link href={`${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`}>
            <Home32Regular className={styles.icon} />
          </Link>
          <span className={styles.title}>Summary of Reserve Matters</span>
        </div>

        <div className={styles.activity}>
          <label className={styles.activitylabel} >Select Site Activity:</label>
          <Dropdown
            selectedKey={activityFilter}
            options={dropdownOptions}
            onChange={(e, option) => setActivityFilter(option?.key as string)}
            styles={{ dropdown: { width: 200 } }}
          />
        </div>

        {activityFilter && (
          <>
            <div className={styles.topBar}>

              <SearchBox
                className={styles.searchBox}
                placeholder="Type here....."
                onChange={(_, val) => setSearchText(val || '')}
              />



              <PrimaryButton
               disabled={filteredItems.length===0}
                onClick={exportToExcel}
                className={styles.exportBtn}>
                  
                <Icon iconName="ExcelDocument" style={{ marginRight: 6 }} />
                Export to Excel
              </PrimaryButton>




            </div>

            {filteredItems.length > 0 ? (
              <DetailsList
                items={paginatedItems}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                constrainMode={ConstrainMode.horizontalConstrained}
                onRenderDetailsHeader={(props, defaultRender) =>
                  defaultRender ? (
                    <div className={headerClass}>
                      {defaultRender(props)}
                    </div>
                  ) : null
                }
              />
            ) : (
              <div
                style={{
                  textAlign: "center",
        padding: "20px",
        border: "1px solid #d1d1d1",
        borderTop: "none",
        backgroundColor: "#fff"
                }}
              >
                No matching records found
              </div>
            )}

            <div className={styles.paginationContainer}>
              <PrimaryButton
              text='Previous'
                onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                disabled={currentPage === 1}
              />

              <span>
                Page {currentPage} of {Math.ceil(filteredItems.length / itemsPerPage)}
              </span>

              <PrimaryButton
                text="Next"
                onClick={() =>
                  setCurrentPage(p =>
                    endIndex < filteredItems.length ? p + 1 : p
                  )
                }
                disabled={endIndex >= filteredItems.length}
              />

            </div>
          </>
        )}
      </section>
    </>
  );
}
