import * as React from 'react';
import styles from './NewSiteRequestsDashboard.module.scss';
import type { INewSiteRequestsDashboardProps } from './INewSiteRequestsDashboardProps';
import { useEffect, useState } from "react";
import { SpService } from '../../../services/spService';
import { Home32Regular } from '@fluentui/react-icons';
import {
  DetailsList,
  IColumn,
  Link,
  mergeStyles,
  DetailsListLayoutMode,
  SearchBox,
  PrimaryButton
} from '@fluentui/react';
import RsrsHeader from '../../rsrsHeader/components/RsrsHeader';
import { Log } from '@microsoft/sp-core-library';


export default function NewSiteRequestsDashboard(props: INewSiteRequestsDashboardProps): React.ReactElement<INewSiteRequestsDashboardProps> {


  const [items, setItems] = useState<any[]>([]);
  const [filteredItems, setFilteredItems] = useState<any[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [sortKey, setSortKey] = useState<string>('');
  const [isSortedDescending, setIsSortedDescending] = useState(false);
  const [isAuthorized, setIsAuthorized] = useState(false);
  const spService = new SpService(props.context);

  const itemsPerPage = 10;

  const onSearch = (text: string) => {
    if (!text) {
      setFilteredItems(items);
      setCurrentPage(1);
      return;
    }

    const filtered = items.filter((item) =>
      (
        item.Request +
        item.SiteName +
        item.SiteRequestedBy +
        item.RequestedDate +
        item.IsActionTaken +
        item.Status
      )
        .toLowerCase()
        .includes(text.toLowerCase())
    );

    setFilteredItems(filtered);
    setCurrentPage(1);
  };
  const getNewSiteRequestData = async (): Promise<void> => {
    try {
      const data = await spService.getItems(
        "NewSiteRequest",
        [
          "ID",
          "RSRSSiteName",
          "Created",
          "IsCreated",
          "IsApproved",
          "Author/Title"
        ],
        undefined,
        {
          field: "ID",
          ascending: false
        },
        4999,
        ["Author"]
      );

      const formattedItems = data.map((item: any) => {

        const createdDate = item.Created
          ? item.Created.split("T")[0].split("-")
          : [];

        const formattedDate =
          createdDate.length === 3
            ? `${createdDate[1]}-${createdDate[2]}-${createdDate[0]}`
            : "";

        return {
          Request: "REQUEST" + item.ID,
          SiteName: item.RSRSSiteName || "",
          SiteRequestedBy: item.Author?.Title || "",
          RequestedDate: formattedDate,
          IsActionTaken: item.IsCreated ? "Yes" : "No",
          Status: item.IsApproved || ""
        };
      });
      console.log(formattedItems, "formattedItemsformattedItemsformattedItems");


      setItems(formattedItems);
      setFilteredItems(formattedItems);

    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };
  const validateDashboardAccess = async (): Promise<boolean> => {

    try {

      const currentUserEmail =
        props.context.pageContext.user.email;

      const contacts =
        await spService.getItems(
          "Contacts",
          [
            "ContactEmail",
            "User_Security_Level"
          ],
          `ContactEmail eq '${currentUserEmail}'`
        );

      if (!contacts || contacts.length === 0) {

        window.location.href =
          `${props.context.pageContext.web.absoluteUrl}/SitePages/AccessDenied.aspx`;

        return false;
      }

      const role =
        contacts[0].User_Security_Level;

      if (role !== "Administrator") {

        window.location.href =
          `${props.context.pageContext.web.absoluteUrl}/SitePages/AccessDenied.aspx`;

        return false;
      }

      return true;

    } catch {

      window.location.href =
        `${props.context.pageContext.web.absoluteUrl}/SitePages/AccessDenied.aspx`;

      return false;
    }
  };

  /*       useEffect(() => {
     void getNewSiteRequestData();
    }, []); */
  useEffect(() => {

    const initializeDashboard = async () => {

      const hasAccess =
        await validateDashboardAccess();

      if (!hasAccess) {
        return;
      }

      await getNewSiteRequestData();
    };

    void initializeDashboard();

  }, []);

  const sortData = (data: any[], key: string, desc: boolean) => {
    return [...data].sort((a, b) => {
      let valA = a[key];
      let valB = b[key];

      if (!valA) return 1;
      if (!valB) return -1;

      if (typeof valA === "string") {
        return desc
          ? valB.localeCompare(valA)
          : valA.localeCompare(valB);
      }

      return desc ? valB - valA : valA - valB;
    });
  };

  const onColumnClick = (ev?: React.MouseEvent<HTMLElement>, column?: IColumn) => {

    if (!column) return;

    const newDesc =
      sortKey === column.fieldName ? !isSortedDescending : false;

    setSortKey(column.fieldName!);
    setIsSortedDescending(newDesc);

    const sorted = sortData(filteredItems, column.fieldName!, newDesc);
    setFilteredItems(sorted);
  };

  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const paginatedItems = filteredItems.slice(startIndex, endIndex);

  const columns: IColumn[] = [

    {
      key: 'Request',
      name: 'Request #',
      fieldName: 'Request',
      minWidth: 108,
      isResizable: true,
      isSorted: sortKey === 'Request',
      isSortedDescending: sortKey === 'Request' ? isSortedDescending : false,
      onColumnClick,

      onRender: (item) => {
        const link = `${props.context.pageContext.web.absoluteUrl}/SitePages/NewSite.aspx?RequestId=${item.Request.replace("REQUEST", "")}`;

        return (
          <a
            className={styles.requestLink}
            onClick={() => window.open(link, "_blank")}
          >
            {item.Request}
          </a>
        );
      }
    },

    {
      key: 'SiteName',
      name: 'Site Name',
      fieldName: 'SiteName',
      minWidth: 250,
      isResizable: true,
      isSorted: sortKey === 'SiteName',
      isSortedDescending: sortKey === 'SiteName' ? isSortedDescending : false,
      onColumnClick
    },

    {
      key: 'SiteRequestedBy',
      name: 'Site Requested By',
      fieldName: 'SiteRequestedBy',
      minWidth: 250,
      isResizable: true,
      isSorted: sortKey === 'SiteRequestedBy',
      isSortedDescending: sortKey === 'SiteRequestedBy' ? isSortedDescending : false,
      onColumnClick
    },

    {
      key: 'RequestedDate',
      name: 'Requested Date',
      fieldName: 'RequestedDate',
      minWidth: 116,
      isResizable: true,
      isSorted: sortKey === 'RequestedDate',
      isSortedDescending: sortKey === 'RequestedDate' ? isSortedDescending : false,
      onColumnClick
    },

    {
      key: 'IsActionTaken',
      name: 'Is Action Taken?',
      fieldName: 'IsActionTaken',
      minWidth: 105,
      isResizable: true,
      isSorted: sortKey === 'IsActionTaken',
      isSortedDescending: sortKey === 'IsActionTaken' ? isSortedDescending : false,
      onColumnClick,

      onRender: (item) => {
        const cls = item.IsActionTaken === 'Yes'
          ? styles.statusApproved
          : styles.statusRejected;

        return <span className={cls}>{item.IsActionTaken}</span>;
      }
    },

    {
      key: 'Status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 120,
      isResizable: true,
      isSorted: sortKey === 'Status',
      isSortedDescending: sortKey === 'Status' ? isSortedDescending : false,
      onColumnClick,

      onRender: (item) => {
        const status = item.Status?.toLowerCase();
        let cls = styles.statusDefault;

        if (status === 'approved') cls = styles.statusApproved;
        else if (status === 'pending') cls = styles.statusPending;
        else if (status === 'rejected') cls = styles.statusRejected;

        return <span className={cls}>{item.Status}</span>;
      }
    }
  ];

  // ✅ Header style
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
      {/* <RsrsHeader context={props.context}  /> */}
      <section className={`${styles.newSiteRequestsDashboard}`}>

        {/* HEADER */}
        <div className={styles.header}>
          <Link href={`${props.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx`}>
            <Home32Regular className={styles.icon} />
          </Link>
          <span className={styles.title}>New Site Requests Dashboard</span>
        </div>



        {/* TABLE */}
        <div className={styles.listContainer}>
          {/* SEARCH */}
          <SearchBox
            // className={styles.searchContainer}
            placeholder="Search..."
            onChange={(_, val) => onSearch(val || '')}
          />
          <DetailsList
            items={paginatedItems}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={0}
            onRenderDetailsHeader={(props, defaultRender) =>
              defaultRender ? (
                <div className={headerClass}>
                  {defaultRender(props)}
                </div>
              ) : null
            }
          />

          {filteredItems.length === 0 && (
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
        </div>

        {/* PAGINATION */}
        {filteredItems.length > 0 && (
          <div className={styles.paginationContainer}>
            <PrimaryButton
              text="Previous"
              onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
              disabled={currentPage === 1}
            />


            <span>
              Page {currentPage} of {Math.ceil(filteredItems.length / itemsPerPage)}
            </span>

            <PrimaryButton
              text="Next"
              disabled={endIndex >= filteredItems.length}
              onClick={() =>
                setCurrentPage(p =>
                  endIndex < filteredItems.length ? p + 1 : p
                )
              }

            />
          </div>
        )}

      </section>
    </>
  );

}

