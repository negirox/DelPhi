import { SearchUserModel } from "../../../models/SearchUserModel";

export interface IDelphiSearchState{
    items: Array<SearchUserModel>,
    errors?: [],
    searchText: string,
    searchResults: Array<SearchUserModel>,
}