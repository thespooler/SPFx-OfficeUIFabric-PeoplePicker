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

export interface IUserListItem {
    Id: number;
    Title: string;
}

export interface IEnsureUser extends IUserListItem {
    Email: string;
    IsEmailAuthenticationGuestUser: boolean;
    IsHiddenInUI: boolean;
    IsShareByEmailGuestUser: boolean;
    IsSiteAdmin: boolean;
    LoginName: string;
    PrincipalType: number;
    UserId: {
        NameId: string;
        NameIdIssuer: string;
    };
}

export interface ISPDataUserInfoItem extends IUserListItem {
    Department: string;
    FirstName: string;
    IsSiteAdmin: boolean;
    JobTitle: string;
    LastName: string;
    Name: string;
    UserName: string;
}

export interface IEnsurableSharePointUser 
    extends IClientPeoplePickerSearchUser, IEnsureUser {}

export interface ISharePointSearchUserPersona extends IPersonaProps {
    user: IEnsurableSharePointUser;
}

export interface ISharePointUserPersona extends IPersonaProps {
    user: IUserListItem;
}

export const SharePointUserInfoPersona = (user: ISPDataUserInfoItem) => ({
    user,
    primaryText: user.Title,
    secondaryText: user.JobTitle,
    tertiaryText: user.Department,
    imageShouldFadeIn: true,
    imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${user.UserName}`
} as ISharePointUserPersona);

export const SharePointSearchUserPersona = (user: IEnsurableSharePointUser) => ({
    user,
    primaryText: user.Title,
    secondaryText: user.EntityData.Title,
    tertiaryText: user.EntityData.Department,
    imageShouldFadeIn: true,
    imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${user.Key.substr(user.Key.lastIndexOf('|') + 1)}`
} as ISharePointUserPersona);
