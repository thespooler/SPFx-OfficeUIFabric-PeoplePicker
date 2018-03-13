import { IPersonaProps, IPersona } from "office-ui-fabric-react";

export interface IOfficeUiFabricPeoplePickerState {
    currentPicker?: number | string;
    delayResults?: boolean;
    selectedItems: any[];
}
export interface IPeopleSearchProps {
    JobTitle: string;
    PictureURL: string;
    PreferredName: string;
}

export interface IUserEntityData {
    IsAltSecIdPresent: string;
    ObjectId: string;
    Title: string;
    Email: string;
    MobilePhone: string;
    OtherMails: string;
    Department: string;
}

export interface IClientPeoplePickerSearchUser {
    Key: string;
    Description: string;
    DisplayText: string;
    EntityType: string;
    ProviderDisplayName: string;
    ProviderName: string;
    IsResolved: boolean;
    EntityData: IUserEntityData;
    MultipleMatches: any[];
}

export interface IEnsureUser {
    Email: string;
    Id: number;
    IsEmailAuthenticationGuestUser: boolean;
    IsHiddenInUI: boolean;
    IsShareByEmailGuestUser: boolean;
    IsSiteAdmin: boolean;
    LoginName: string;
    PrincipalType: number;
    Title: string;
    UserId: {
        NameId: string;
        NameIdIssuer: string;
    };
}

export interface ISPDataUserInfoItem {
    Department: string;
    FirstName: string;
    Id: number;
    IsSiteAdmin: boolean;
    JobTitle: string;
    LastName: string;
    Name: string;
    Title: string;
    UserName: string;
}

export interface IEnsurableSharePointUser 
    extends IClientPeoplePickerSearchUser, IEnsureUser {}

export interface ISharePointUserPersona extends IPersonaProps {
    user: IEnsurableSharePointUser;
}

export const SharePointUserInfoPersona = (user: ISPDataUserInfoItem) => ({
    primaryText: user.Title,
    secondaryText: user.JobTitle,
    tertiaryText: user.Department,
    imageShouldFadeIn: true,
    imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${user.UserName}`
} as IPersona);

export const SharePointSearchUserPersona = (user: IEnsurableSharePointUser) => ({
    user: user,
    primaryText: user.Title,
    secondaryText: user.EntityData.Title,
    tertiaryText: user.EntityData.Department,
    imageShouldFadeIn: true,
    imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${user.Key.substr(user.Key.lastIndexOf('|') + 1)}`
} as ISharePointUserPersona);
