import * as React from 'react';
import styles from './QueryFilter.module.scss';
import type { IQueryFilterProps } from './IQueryFilterProps';
import {
  Dropdown,
  PrimaryButton,
  Spinner,
  MessageBar,
  Link,
  IDropdownOption,
  SearchBox,
  Icon,
  mergeStyles,
  MessageBarType
} from "@fluentui/react";

import {
  DetailsList,
  SelectionMode,
  DetailsListLayoutMode,
  IColumn
} from "@fluentui/react/lib/DetailsList";

import * as XLSX from "xlsx-js-style";
import { saveAs } from "file-saver";
import { SpService } from '../../../services/spService';
import { Home32Regular } from "@fluentui/react-icons";
import { useState, useCallback } from "react";
import RsrsHeader from '../../rsrsHeader/components/RsrsHeader';



export default function QueryFilter(props: IQueryFilterProps): React.ReactElement<IQueryFilterProps> {
  const [query, setQuery] = useState("");
  const [items, setItems] = useState<any[]>([]);
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [searchText, setSearchText] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [sortColumn, setSortColumn] = useState("");
  const [isDescending, setIsDescending] = useState(false);
  const [warning, setWarning] = useState("");
  const itemsPerPage = 10;
  const spService = new SpService(props.context);
  const formatLegacyNumber = (value: any): string => {
    if (value === null || value === undefined || value === "") {
      return "";
    }

    const numParts = value.toString().split(".");

    numParts[0] = numParts[0].replace(
      /\B(?=(\d{3})+(?!\d))/g,
      ","
    );

    if (numParts[1] && numParts[1] !== "00") {
      return numParts.join(".");
    }

    return numParts[0];
  };
  const onColumnClick = useCallback(
    (
      ev?: React.MouseEvent<HTMLElement>,
      column?: IColumn
    ): void => {

      if (!column || !column.fieldName) {
        return;
      }

      const fieldName = column.fieldName;

      const desc =
        sortColumn === fieldName
          ? !isDescending
          : false;

      const sortedItems = [...items].sort((a, b) => {
        const aValue = a[fieldName];
        const bValue = b[fieldName];

        // Handle null/undefined
        if (aValue == null && bValue == null) return 0;
        if (aValue == null) return 1;
        if (bValue == null) return -1;

        // Numeric sort
        const numA = Number(
          String(aValue).replace(/,/g, "")
        );

        const numB = Number(
          String(bValue).replace(/,/g, "")
        );

        if (!isNaN(numA) && !isNaN(numB)) {
          return desc
            ? numB - numA
            : numA - numB;
        }

        // String sort
        return desc
          ? String(bValue).localeCompare(
            String(aValue)
          )
          : String(aValue).localeCompare(
            String(bValue)
          );
      });

      const updatedColumns = columns.map((col) => ({
        ...col,
        isSorted: col.fieldName === fieldName,
        isSortedDescending:
          col.fieldName === fieldName
            ? desc
            : false
      }));

      setItems(sortedItems);
      setColumns(updatedColumns);
      setSortColumn(fieldName);
      setIsDescending(desc);
    },
    [items, columns, sortColumn, isDescending]
  );
  const displayColumns = React.useMemo(() => {
    return columns.map(col => ({
      ...col,
      onColumnClick: onColumnClick
    }));
  }, [columns, onColumnClick]);
  /* const onColumnClick = (
    ev?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ): void => {
    console.log("hiii",items);
    
    if (!column || !column.fieldName) return;
  
    const fieldName = column.fieldName;
  
    const desc =
      sortColumn === fieldName
        ? !isDescending
        : false;
  
    const sortedItems = [...items].sort((a, b) => {
      const aValue = a[fieldName];
      const bValue = b[fieldName];
  
      // Handle null / undefined
      if (aValue == null && bValue == null) return 0;
      if (aValue == null) return 1;
      if (bValue == null) return -1;
  
      // Remove commas and convert to number
      const numA = Number(
        String(aValue).replace(/,/g, "")
      );
  
      const numB = Number(
        String(bValue).replace(/,/g, "")
      );
  
      // Numeric sorting
      if (!isNaN(numA) && !isNaN(numB)) {
        return desc
          ? numB - numA
          : numA - numB;
      }
  
      // String sorting
      return desc
        ? String(bValue).localeCompare(
            String(aValue)
          )
        : String(aValue).localeCompare(
            String(bValue)
          );
    });
  
    const updatedColumns = columns.map((col) => ({
      ...col,
      isSorted: col.fieldName === fieldName,
      isSortedDescending:
        col.fieldName === fieldName
          ? desc
          : false
    }));
    console.log(sortedItems);
    
  
    setItems(sortedItems);
    setColumns(updatedColumns);
    setSortColumn(fieldName);
    setIsDescending(desc);
  }; */



  const runQuery = async () => {
    if (!query) {
      setWarning("Please select a query.");
      setTimeout(() => {
        setWarning("");
      }, 3000);
      return;
    }

    try {
      setLoading(true);
      setError("");

      let data: any[] = [];
      let cols: IColumn[] = [];

      switch (query) {
        case "legal":
          data = await spService.getLegalExport(props.context);
          cols = [
            {
              key: "1",
              name: "RSRS Site ID",
              fieldName: "Title",
              isResizable: true,
              minWidth: 100
            },
            {
              key: "2",
              name: "RSRS Site Name",
              fieldName: "RSRSSiteName",
              isResizable: true,
              minWidth: 120
            },
            {
              key: "3",
              name: "City",
              fieldName: "City",
              isResizable: true,
              minWidth: 100
            },
            {
              key: "4",
              name: "State",
              fieldName: "State",
              isResizable: true,
              minWidth: 100
            },
            {
              key: "5",
              name: "Country",
              fieldName: "Country",
              isResizable: true,
              minWidth: 100
            },
            {
              key: "6",
              name: "Site Type",
              isResizable: true,
              fieldName: "Site_Type",
              minWidth: 120
            },
            {
              key: "7",
              name: "Site Activity",
              isResizable: true,
              fieldName: "Site_Activity",
              minWidth: 120
            },
            {
              key: "8",
              name: "Legacy Company",
              isResizable: true,
              fieldName: "Legacy_Company",
              minWidth: 120
            },
            {
              key: "9",
              name: "Site Lead",
              isResizable: true,
              fieldName: "SiteLead",
              minWidth: 120
            },
            {
              key: "10",
              name: "Site Support",
              isResizable: true,
              fieldName: "SiteSupport",
              minWidth: 120
            },
            {
              key: "11",
              name: "RSRS Update Owner",
              isResizable: true,
              fieldName: "SiteUpdateOwner",
              minWidth: 120
            },
            {
              key: "12",
              name: "Outside Counsel",
              isResizable: true,
              fieldName: "OutsideCounsel",
              minWidth: 120
            },
            {
              key: "13",
              name: "Account Project Code",
              isResizable: true,
              fieldName: "AccountProjectCode",
              minWidth: 120
            },
            {
              key: "14",
              name: "Claim Type",
              isResizable: true,
              fieldName: "Claim_Type",
              minWidth: 120
            },
            {
              key: "15",
              name: "Total Project Cost",
              fieldName: "TotalProjectCost",
              minWidth: 120,
              onRender: (item) => formatLegacyNumber(item.TotalProjectCost)
            },
            {
              key: "16",
              name: "Total Spending to Date",
              fieldName: "TotalSpendingToDate",
              minWidth: 120,
              onRender: (item) => formatLegacyNumber(item.TotalSpendingToDate)
            },
            {
              key: "17",
              name: "Reserve",
              fieldName: "ReserveBalance",
              minWidth: 120,
              onRender: (item) => formatLegacyNumber(item.ReserveBalance)
            }
          ];
          break;

        case "active":
          data = await spService.getActiveContacts(props.context);
          console.log(data, "datadatadata");


          cols = [
            {
              key: "1",
              name: "First Name",
              fieldName: "Title",
              isResizable: true,
              minWidth: 120
            },
            {
              key: "2",
              name: "Last Name",
              fieldName: "LastName",
              isResizable: true,
              minWidth: 120
            },
            {
              key: "3",
              name: "Contact Type",
              isResizable: true,
              fieldName: "Contact_Type",
              minWidth: 100
            },
            {
              key: "4",
              name: "Contact Email",
              isResizable: true,
              fieldName: "Email",
              minWidth: 150
            },
            {
              key: "5",
              name: "Telephone",
              isResizable: true,
              fieldName: "Phone",
              minWidth: 150
            },
            {
              key: "6",
              name: "Pfizer Network User Name",
              isResizable: true,
              fieldName: "Username",
              minWidth: 120
            },
            {
              key: "7",
              name: "User Security Level",
              isResizable: true,
              fieldName: "Security",
              minWidth: 120
            }
            ,
            {
              key: "8",
              name: "Date Contact Added",
              isResizable: true,
              fieldName: "Date",
              minWidth: 150,
              maxWidth: 150,
            },
            {
              key: "9",
              name: "Contact Company",
              isResizable: true,
              fieldName: "Company",
              minWidth: 150
            }
          ];
          break;

        case "negative":
          data = await spService.getNegativeReserve(props.context);
          console.log("datataaa",data);
          
          cols = [
            {
              key: "1",
              name: "RSRS Site ID",
              fieldName: "Title",
              isResizable: true,
              minWidth: 200,

            },
            {
              key: "2",
              name: "RSRS Site Name",
              fieldName: "RSRSSiteName",
              isResizable: true,
              minWidth: 350,

            },
            {
              key: "3",
              name: "Reserve Balance",
              fieldName: "Reserve",
              isResizable: true,
              minWidth: 250,
              onRender: (item) =>
                formatLegacyNumber(
                  Number(item.Reserve).toFixed(2)
                )
            },
            {
              key: "4",
              name: "Site Activity",
              fieldName: "Site_Activity",
              isResizable: true,
              minWidth: 250,

            },

          ];
          break;

        case "quarter":
          data = await spService.getQuarterReview(props.context);

          cols = [
            {
              key: "1",
              name: "Claim Type",
              fieldName: "Claim_Type",
              isResizable: true,
              minWidth: 120
            },
            {
              key: "2",
              name: "RSRS Site ID",
              fieldName: "Title",
              isResizable: true,
              minWidth: 120
            },
            {
              key: "3",
              name: "RSRS Site Name",
              fieldName: "RSRSSiteName",
              isResizable: true,
              minWidth: 120
            },
            {
              key: "4",
              name: "Site Activity",
              fieldName: "Site_Activity",
              isResizable: true,
              minWidth: 120
            },
            {
              key: "5",
              name: "Site Type",
              isResizable: true,
              fieldName: "Site_Type",
              minWidth: 100
            },
            {
              key: "6",
              name: "Legacy Company",
              isResizable: true,
              fieldName: "Legacy_Company",
              minWidth: 130
            },
            {
              key: "7",
              name: "Cost Estimation Method",
              isResizable: true,
              fieldName: "Cost_Estimate_Method",
              minWidth: 170
            },
            {
              key: "8",
              name: "Stage",
              isResizable: true,
              fieldName: "Project_Stage",
              minWidth: 100
            },
            {
              key: "9",
              name: "Site Lead",
              isResizable: true,
              fieldName: "Sitelead",
              minWidth: 100
            },
            {
              key: "10",
              name: "Site Support",
              isResizable: true,
              fieldName: "SiteSupport",
              minWidth: 100
            },
            {
              key: "11",
              name: "RSRS Update Owner",
              isResizable: true,
              fieldName: "SiteUpdateOwner",
              minWidth: 130
            },
            {
              key: "12",
              name: "Outside Counsel",
              isResizable: true,
              fieldName: "outsideC",
              minWidth: 130
            },
            {
              key: "13",
              name: "Assistant Project Manager",
              isResizable: true,
              fieldName: "assistantPM",
              minWidth: 170
            },
            {
              key: "14",
              name: "Technical Consultant",
              isResizable: true,
              fieldName: "technicalC",
              minWidth: 150
            },
            {
              key: "15",
              name: "Other Additional Support",
              isResizable: true,
              fieldName: "otherAdditionalS",
              minWidth: 170
            },
            {
              key: "16",
              name: "RSRS Total Project Cost",
              fieldName: "TotalProjectCost",
              minWidth: 150,
              onRender: (item) =>
                formatLegacyNumber(item.TotalProjectCost)
            },
            {
              key: "17",
              name: "BET Project Cost",
              fieldName: "BETEstimatedCost",
              minWidth: 150,
              onRender: (item) =>
                formatLegacyNumber(item.BETEstimatedCost)
            }




          ];



          break;

        case "rfts":
          data = await spService.getRFTS(props.context);

          cols = [
            {
              key: "1",
              name: "RSRS Site ID",
              fieldName: "Title",
              minWidth: 120
            },
            {
              key: "2",
              name: "Site Name",
              fieldName: "RSRSSiteName",
              minWidth: 180
            },
            {
              key: "3",
              name: "Site Type",
              fieldName: "Site_Type",
              minWidth: 140
            },
            {
              key: "4",
              name: "Site Activity",
              fieldName: "Site_Activity",
              minWidth: 140
            },
            {
              key: "5",
              name: "Company",
              fieldName: "Legacy_Company",
              minWidth: 140
            },
            {
              key: "6",
              name: "Technical Project Manager",
              fieldName: "TechnicalProjectManager",
              minWidth: 180
            },
            {
              key: "7",
              name: "PGE Role",
              fieldName: "PGERole",
              minWidth: 140
            },
            {
              key: "8",
              name: "Account",
              fieldName: "AccountProjectCode",
              minWidth: 150
            },
            {
              key: "9",
              name: "Total Project Cost",
              fieldName: "TotalProjectCost",
              minWidth: 150,
              onRender: (item) =>
                formatLegacyNumber(item.TotalProjectCost)
            }
          ];
          break;
      }

      // After switch block

      setSortColumn("");
      setIsDescending(false);

      const sortableColumns = cols.map((col) => ({
        ...col,
        isSorted: false,
        isSortedDescending: false
      }));

      setColumns(sortableColumns);
      console.log(sortableColumns, "sortableColumnssortableColumnssortableColumns");
      console.log(cols, "colscols");


      setItems(data);
      //setColumns(sortableColumns);
      setColumns(cols);

      setSearchText("");
      setCurrentPage(1);
    } catch (err) {
      console.error(err);
      setError("Error loading data. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  // Search All Columns
  const filteredItems = items.filter((item) =>
    columns.some((column) => {
      const value = item[column.fieldName || ""];
      return (
        value &&
        value
          .toString()
          .toLowerCase()
          .includes(searchText.toLowerCase())
      );
    })
  );

  const totalPages = Math.ceil(
    filteredItems.length / itemsPerPage
  );

  const paginatedItems = filteredItems.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  // Export To Excel
  const exportToExcel = () => {
    if (!filteredItems.length) {
      alert("No data available to export.");
      return;
    }

    const exportData = filteredItems.map((item) => {
      const row: any = {};

      columns.forEach((col) => {
        row[col.name] = item[col.fieldName || ""];
      });

      return row;
    });

    //const worksheet = XLSX.utils.json_to_sheet(exportData);
    const worksheet = XLSX.utils.json_to_sheet(exportData);

    // Make header row bold
    const range = XLSX.utils.decode_range(worksheet["!ref"] || "");

    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({
        r: 0,
        c: C
      });

      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: {
            bold: true,
            color: { rgb: "FFFFFF" }
          },
          fill: {
            fgColor: { rgb: "0078D4" }
          }
        };
      }
    }

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Query Results"
    );
    //test
    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array"
    });

    const blob = new Blob([excelBuffer], {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8"
    });

    saveAs(
      blob,
      `${query || "Query"}_Results.xlsx`
    );
  };
  const dropdownOptions: IDropdownOption[] = [
    { key: "", text: "--Select--" },
    { key: "active", text: "Active Contacts Query" },
    { key: "legal", text: "Legal Export Query" },
    { key: "negative", text: "Negative Reserve Query" },
    { key: "quarter", text: "Quarter Review Query" },
    { key: "rfts", text: "RFTS Export Query" }
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
      {/* <RsrsHeader context={props.context} /> */}
      <section className={`${styles.queryFilter} `}>
        <div className={styles.header}>
          <Link href={`${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`}>
            <Home32Regular className={styles.icon} />
          </Link>

          <span className={styles.title}>
            Query Filter
          </span>
        </div>

        <div className={styles.query}>
          <label className={styles.querylabel}>
            Select Query:  </label>

          <Dropdown
            options={dropdownOptions}
            selectedKey={query}
            onChange={(e, option) => {
              const selectedValue = option?.key as string;

              setQuery(selectedValue);
              // Clear grid
              setItems([]);
              setColumns([]);
              setCurrentPage(1);
              setSearchText("");

              if (!selectedValue) {
                // Show warning
                setWarning("Please select a query.");
                setTimeout(() => {
                  setWarning("");
                }, 3000);
              }
            }}
            styles={{
              dropdown: { width: 300 }
            }}
          />

          `

          <PrimaryButton
            className={styles.runButton}
            onClick={runQuery}
          // disabled={!query}
          >
            Run Query
          </PrimaryButton>
        </div>
        {warning && (
          <MessageBar
            messageBarType={MessageBarType.warning} // Warning
            isMultiline={false}
          >
            {warning}
          </MessageBar>
        )}

        {loading && (
          <Spinner label="Loading data..." />
        )}

        {error && (
          <MessageBar messageBarType={3}>
            {error}
          </MessageBar>
        )}

        {!loading && !error && items.length > 0 && (
          <>
            <div className={styles.topBar}>
              <SearchBox
                className={styles.searchBox}
                placeholder="Type here..."
                value={searchText}
                onChange={(_, value) => {
                  setSearchText(value || ""); setCurrentPage(1);
                }}
              />

              <PrimaryButton
                onClick={exportToExcel}
                className={styles.exportBtn}
                disabled={filteredItems.length === 0}
              >
                <Icon
                  iconName="ExcelDocument"
                  style={{ marginRight: 6 }}
                />
                Export to Excel
              </PrimaryButton>
            </div>
            <div className={styles.gridDivider}></div>

            {filteredItems.length > 0 && (
              <DetailsList
                items={paginatedItems}
                columns={displayColumns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                onRenderDetailsHeader={(props, defaultRender) => defaultRender ? <div className={headerClass}>{defaultRender(props)}</div> : null}
              />
            )}
            {filteredItems.length === 0 && (
              <div
                style={{
                  textAlign: "center",
                  padding: "20px",
                  border: "1px solid #d1d1d1",
                  borderTop: "none",
                  color: "#323130"
                }}
              >
                No matching records found
              </div>
            )}
            {filteredItems.length > 0 && (
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                  marginTop: "15px"
                }}
              >
                <PrimaryButton
                  text="Previous"
                  disabled={currentPage === 1}
                  onClick={() =>
                    setCurrentPage((prev) => prev - 1)
                  }
                />

                <span>
                  Page {currentPage} of {totalPages}
                </span>

                <PrimaryButton
                  text="Next"
                  disabled={currentPage === totalPages}
                  onClick={() =>
                    setCurrentPage((prev) => prev + 1)
                  }
                />
              </div>
            )}
          </>
        )}

      </section>
    </>
  );
}

