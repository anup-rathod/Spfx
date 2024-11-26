import * as React from "react";

export interface IWebpart2State {
      ID: React.Key;
      Title: string;
      Lookup: string;
      data: [];
      lookupOptions: any[];
      columns: any;
      selectedUsers: any;
  }