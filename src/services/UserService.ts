import "@pnp/graph/users";  
import { IUser } from '@pnp/graph/users';
import { getGraph, getSP } from "../pnpjsConfig";
import { SearchQueryInit, SearchResults } from "@pnp/sp/search";
import axios, { AxiosResponse } from "axios";
export class UserService{
    public static async GetAllUsers(): Promise<IUser[]> {
        const graph = getGraph();
        console.log(graph);
        const users = await graph.users();
        console.log(users);
        return await graph.users();
    }
    public static getSearchItems(k: string, startRow: number, rowLimit: number): Promise<SearchResults> {
        const sp = getSP();  
        return new Promise<SearchResults>((resolve, reject) => {  
            const query: SearchQueryInit = {  
                Querytext: k,  
                TrimDuplicates: true,  
                EnableInterleaving: true,  
                StartRow: startRow,  
                RowLimit: rowLimit,
            };  
  
            sp.search(query)  
                .then((r: SearchResults) => {  
                    resolve(r);  
                }, e => { console.error(e); reject(e); });  
        });  
    }
    public static async getUserFromSearch(searchUrl:string):Promise<AxiosResponse<any,any>>{
        return await axios.get(searchUrl);
    }
}