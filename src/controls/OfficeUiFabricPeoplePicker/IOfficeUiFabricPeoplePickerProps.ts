import { SPHttpClient } from '@microsoft/sp-http';
import { SharePointUserPersona } from '../../index';

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
  users?: SharePointUserPersona[];
  onChange?: (users: SharePointUserPersona[]) => void;
}
