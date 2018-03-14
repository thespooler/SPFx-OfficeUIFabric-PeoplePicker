import { SPHttpClient } from '@microsoft/sp-http';
import { IPersonaProps } from 'office-ui-fabric-react';
import { PrincipalType, TypePicker } from '.';

export interface IOfficeUiFabricPeoplePickerProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  typePicker: TypePicker;
  principalType: PrincipalType;
  numberOfItems: number;
  selectedItems?: IPersonaProps[];
  onChange?: (users: IPersonaProps[]) => void;
}
