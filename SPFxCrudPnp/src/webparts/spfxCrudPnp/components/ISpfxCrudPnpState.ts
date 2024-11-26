import * as React from "react";

export interface ISpfxCrudState {
    ID: React.Key;
    Title: string;
    Description: string;
    Address:string;
    data:any;
    Choice: string;
    ChoiceOptions: [];
    selectedUsers: any[];
    Lookup: string,
  }