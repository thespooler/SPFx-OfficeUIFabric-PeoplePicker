import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneToggle, PropertyPaneSlider } from "@microsoft/sp-webpart-base";
import { Version } from "@microsoft/sp-core-library";
import { IOfficeUiFabricPeoplePickerWebPartProps } from "./IOfficeUiFabricPeoplePickerWebPartProps";
import { PrincipalType, IOfficeUiFabricPeoplePickerProps, OfficeUiFabricPeoplePicker, TypePicker } from "../..";
import React = require("react");
import * as ReactDom from 'react-dom';
import * as strings from 'OfficeUiFabricPeoplePickerWebPartStrings';

export default class OfficeUiFabricPeoplePickerWebPart extends BaseClientSideWebPart<IOfficeUiFabricPeoplePickerWebPartProps> {

  public render(): void {
    let principalType: PrincipalType = PrincipalType.None;
    principalType |= this.properties.principalTypeUser ? PrincipalType.User : PrincipalType.None;
    principalType |= this.properties.principalTypeDistributionList ? PrincipalType.DistributionList : PrincipalType.None;
    principalType |= this.properties.principalTypeSecurityGroup ? PrincipalType.SecurityGroup : PrincipalType.None;
    principalType |= this.properties.principalTypeSharePointGroup ? PrincipalType.SharePointGroup : PrincipalType.None;

    const element: React.ReactElement<IOfficeUiFabricPeoplePickerProps> = React.createElement(
      OfficeUiFabricPeoplePicker,
      {
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        typePicker: this.properties.typePicker == TypePicker.Compact ? TypePicker.Compact: TypePicker.Normal,
        principalType: principalType,
        numberOfItems: this.properties.numberOfItems
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('typePicker', {
                  label: strings.TypePickerLabel,
                  selectedKey: TypePicker.Normal,
                  options: [
                    { key: TypePicker.Normal, text: 'Normal' },
                    { key: TypePicker.Compact, text: 'Compact' }
                  ]
                }),
                PropertyPaneToggle('principalTypeUser', {
                    label: strings.principalTypeUserLabel,
                    checked: true,
                  }
                ),
                PropertyPaneToggle('principalTypeSharePointGroup', {
                    label: strings.principalTypeSharePointGroupLabel,
                    checked: true,
                  }
                ),
                PropertyPaneToggle('principalTypeSecurityGroup', {
                    label: strings.principalTypeSecurityGroupLabel,
                    checked: false,
                  }
                ),
                PropertyPaneToggle('principalTypeDistributionList', {
                    label: strings.principalTypeDistributionListLabel,
                    checked: false,
                  }
                ),
                PropertyPaneSlider('numberOfItems', {
                  label: strings.numberOfItemsFieldLabel,
                  min: 1,
                  max: 20,
                  step: 1
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
