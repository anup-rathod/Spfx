import * as React from 'react';
import styles from './MaterialGrid.module.scss';
import type { IMaterialGridProps } from './IMaterialGridProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  MaterialReactTable,
} from 'material-react-table';
import { Box, Button } from '@mui/material';
import IMaterialGridState from './MaterialGridState';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { MdDelete } from "react-icons/md";
import { MdEdit } from "react-icons/md";

export default class MaterialGrid extends React.Component<IMaterialGridProps, IMaterialGridState> {

  constructor(props : IMaterialGridProps){
    super(props)
    const column = [
      {
        header: 'Edit',
        accessorKey: 'IsApproved',
        size: 120,
        Cell: ({ row }: any) => (
          <MdEdit onClick={() => this.handleEdit(row.original.ID)} />
        ),
      },
      {
        header: 'ID', 
        accessorKey: 'ID',
        size: 120,
      },
      {
        header: 'Invoice No',
        accessorKey: 'InvoiceNo',
        size: 120,
      },
      {
        header: 'Company Name',  
        accessorKey: 'CompanyName',
        size: 120,
      },
      {
        header: 'Invoice Details',  
        accessorKey: 'InvoiceDetails',
        size: 120,
      },
      {
        header: 'Company Code',  
        accessorKey: 'CompanyCode',
        size: 120,
      },
      {
        header: 'Invoice Amount',  
        accessorKey: 'InvoiceAmount',
        size: 120,
      },
      {
        header: 'Basic Value',  
        accessorKey: 'BasicValue',
        size: 120,
      },
      {
        header: 'Approver',
        accessorKey: 'Approver',
        size: 120,
        Cell: ({ row }: any) => (
          <div style={{ fontWeight: 'bold' }}>
            {row.original.Approver ? (
              row.original.Approver.map((approver: any, index: number) => (
                <div key={index}>{approver.Title}</div>
              ))
            ) : (
              <div>No Approver Assigned</div>
            )}
          </div>
        )
        
      },    
      {
        header: 'IsApproved',  
        accessorKey: 'IsApproved',
        size: 120,
        Cell: ({ row }: any) => (
          <div style={{ fontWeight: 'bold' }}>
            {row.original.IsApproved.toString()}
          </div>
        )
      },
      {
        header: 'Country',  
        accessorKey: 'Country',
        size: 120,
      },
      {
        header: 'isDeleted',  
        accessorKey: 'isDeleted',
        size: 120,
      },
      {
        header: 'Delete',
        accessorKey: 'IsApproved',
        size: 120,
        Cell: ({ row }: any) => (
          <MdDelete onClick={() => this.handleDelete(row.original.ID)} />
        ),
      },
    ]
    
    this.state = {
      data : [],
      columns: column,
      isDeleted: true,
    } 
  }

  public async componentDidMount(): Promise<void> {
    try {
      this.getList();
    } catch (error) {
      console.log("ComponentDidmount : error ", error);
    }  
  }
  //Get List 
  getList = async (): Promise<void> => {  
    try {
      const sp = spfi().using(SPFx(this.props.context));
      // const deleted:boolean = false;
            
      const items = await sp.web.lists.getByTitle("Task1").items.select("ID","InvoiceNo","CompanyName","InvoiceDetails","CompanyCode","InvoiceAmount","BasicValue","Approver/Title","IsApproved","Country","isDeleted")
      .expand("Approver")
      .filter("isDeleted eq false")()
        
      console.log("Retrieved items:", items);
  
      this.setState({
        data: items
      });
    } catch (error) {
      console.error("Error in getList:", error);
    }
  }; 
  //Edit Form Path
  handleEdit = (id: number) => {
    const { context } = this.props;
    const editPageUrl = `${context.pageContext.web.absoluteUrl}/SitePages/Form-Update.aspx?itemID=${id}`;
    window.location.href = editPageUrl;
  };
  
  //To handle Delete
  handleDelete = async (id: number) => {
    const {isDeleted} = this.state as {
      isDeleted : boolean;
    }
    try {
      const sp = spfi().using(SPFx(this.props.context));

      await sp.web.lists.getByTitle("Task1").items.getById(id).update({
        'isDeleted': isDeleted,
      });

      this.getList();
      alert("Delete Successfully")
    } catch (error) {
      console.error("Error in handleDelete:", error);
    }
  };


  public render(): React.ReactElement<IMaterialGridProps> {
    // console.log(this.state.data,"mydata")
    return (
      <>
      <MaterialReactTable
          displayColumnDefOptions={{
            'mrt-row-actions': {
              muiTableHeadCellProps: {
                align: 'center',
              },
              size: 120,
            },
          }}
          columns={this.state.columns}
          data={this.state.data}
          // state={{ isLoading: true }}
          enableColumnResizing
          initialState={{ density: 'compact', pagination: { pageIndex: 0, pageSize: 100 }, showColumnFilters: true }}
          columnResizeMode="onEnd"
          positionToolbarAlertBanner="bottom"
          enablePinning
          // enableRowActions
          // onEditingRowSave={this.handleSaveRowEdits}
          // onEditingRowCancel={this.handleCancelRowEdits}
          enableGrouping
          enableStickyHeader
          enableStickyFooter
          enableDensityToggle={false}
          enableExpandAll={false}
          renderTopToolbarCustomActions={({ table }) => (
            <Box
              sx={{ display: 'flex', gap: '1rem', p: '0.5rem', flexWrap: 'wrap' }}
            >
            </Box>
          )}
        />
      </>
    );
  }
}
