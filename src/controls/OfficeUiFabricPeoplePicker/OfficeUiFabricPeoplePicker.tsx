import * as strings from 'OfficeUiFabricPeoplePickerStrings';
import React = require('react');
import {
  IOfficeUiFabricPeoplePickerProps,
  IOfficeUiFabricPeoplePickerState,
  IClientPeoplePickerSearchUser,
  SharePointSearchUserPersona,
  IEnsurableSharePointUser,
  IEnsureUser,
  TypePicker
} from '.';
import { NormalPeoplePicker, IPersonaProps, CompactPeoplePicker, IBasePickerSuggestionsProps, autobind } from 'office-ui-fabric-react';
import { people } from '../../webparts/officeUiFabricPeoplePicker/PeoplePickerExampleData';
import { EnvironmentType, Environment } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientBatch } from '@microsoft/sp-http';
import * as lodash from 'lodash';

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: strings.suggestions,
  noResultsFoundText: strings.noResults,
  loadingText: strings.loading
};

export class OfficeUiFabricPeoplePicker extends React.Component<IOfficeUiFabricPeoplePickerProps, IOfficeUiFabricPeoplePickerState> {

  constructor(props: IOfficeUiFabricPeoplePickerProps, context?: any) {
    super(props, context);
    this.state = {
      selectedItems: props.selectedItems
    };
  }

  public render(): React.ReactElement<IOfficeUiFabricPeoplePickerProps> {
    if (this.props.typePicker === TypePicker.Normal) {
      return (
        <NormalPeoplePicker
          onChange={this._onChange.bind(this)}
          onResolveSuggestions={this._onFilterChanged}
          getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
          pickerSuggestionsProps={suggestionProps}
          className={'ms-PeoplePicker'}
          selectedItems={this.state.selectedItems}
          key={'normal'}
        />
      );
    } else {
      return (
        <CompactPeoplePicker
          onChange={this._onChange.bind(this)}
          onResolveSuggestions={this._onFilterChanged}
          getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
          pickerSuggestionsProps={suggestionProps}
          selectedItems={this.state.selectedItems}
          className={'ms-PeoplePicker'}
          key={'normal'}
        />
      );
    }
  }

  private _onChange(items: any[]) {
    this.setState({
      selectedItems: items
    });
    if (this.props.onChange) {
      this.props.onChange(items);
    }
  }

  @autobind
  private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
    if (!filterText || filterText.length < 3) return Promise.resolve([] as IPersonaProps[]);

    return this._searchPeople(filterText);
  }

  /**
   * @function
   * Returns people results after a REST API call
   */
  private _searchPeople(terms: string) {
    // If the running environment is local, load the data from the mock
    if (DEBUG && Environment.type === EnvironmentType.Local) {
      return Promise.resolve(people);
    } 

    const userRequestUrl = `${this.props.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const ensureUserUrl = `${this.props.siteUrl}/_api/web/ensureUser`;
    const userQueryParams = {
      'queryParams': {
        'AllowEmailAddresses': true,
        'AllowMultipleEntities': false,
        'AllUrlZones': false,
        'MaximumEntitySuggestions': this.props.numberOfItems,
        'PrincipalSource': 15,
        'PrincipalType': this.props.principalType,
        'QueryString': terms
      }
    };

    return this.props.spHttpClient.post(userRequestUrl,
      SPHttpClient.configurations.v1, { body: JSON.stringify(userQueryParams) })
      .then((httpResponse: SPHttpClientResponse) => {
        return httpResponse.json();
      })
      .then((response: {value: string}) => {
        const batch = this.props.spHttpClient.beginBatch();
        let userQueryResults: IClientPeoplePickerSearchUser[] = JSON.parse(response.value);
        const batchPromises = userQueryResults.map(p =>
          batch.post(ensureUserUrl, SPHttpClientBatch.configurations.v1, 
            { 
              body: JSON.stringify({ logonName: p.Key })
            })
          .then(httpResponse => httpResponse.json())
          .then((user: IEnsureUser) => ({ ...p, ...user } as IEnsurableSharePointUser))
        );

        return batch.execute().then(() => 
          Promise.all(batchPromises).then(users => users.map(u => SharePointSearchUserPersona(u))));
      });
  }
}
