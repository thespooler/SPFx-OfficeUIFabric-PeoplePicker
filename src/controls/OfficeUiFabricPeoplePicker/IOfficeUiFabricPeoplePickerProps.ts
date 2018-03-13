import { SPHttpClient } from '@microsoft/sp-http';
import { IPersonaProps } from 'office-ui-fabric-react';

export interface IOfficeUiFabricPeoplePickerProps {
  description: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  typePicker: string;
  principalTypeUser: boolean;
  principalTypeSharePointGroup: boolean;
  principalTypeSecurityGroup: boolean;
  principalTypeDistributionList: boolean;
  numberOfItems: number;
  selectedItems?: IPersonaProps[];
  onChange?: (users: IPersonaProps[]) => void;
}
