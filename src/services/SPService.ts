import { spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";

import "@pnp/sp/lists";
import "@pnp/sp/lists/web";

import "@pnp/sp/fields";
import "@pnp/sp/fields/select";

import "@pnp/sp/items";
import "@pnp/sp/items/select";

import "@pnp/sp/folders";
import "@pnp/sp/folders/files";

import "@pnp/sp/files";
import "@pnp/sp/files/item";
import "@pnp/sp/files/select";

import { ISPService } from "./ISPService";
import { IDocumentDto } from "../dtos/IDocumentDto";
import { IListItemDto } from "../dtos/IListItemDto";
import { IFieldInfoDto } from "../dtos/IFieldInfoDto";

import { WebPartContext } from "@microsoft/sp-webpart-base";

export class SPService implements ISPService {
    private sp;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    // -------------------------------------------------------
    // DOCUMENT LIBRARY: Strongly typed IDocument[]
    // -------------------------------------------------------
    public async getDocuments(listTitle: string): Promise<IDocumentDto[]> {
        const items = await this.sp.web.lists
            .getByTitle(listTitle)
            .items
            .select(
                "Id",
                "FileLeafRef",
                "FileRef",
                "Modified",
                "Editor/Title",
                "File_x0020_Size",
                "File_x0020_Type",
                "*"
            )
            .expand("Editor")();

        return items.map(i => this.mapToDocument(i));
    }

    private mapToDocument(i: any): IDocumentDto {
        return {
            id: i.Id,
            name: i.FileLeafRef,
            url: i.FileRef,
            modified: new Date(i.Modified),
            modifiedBy: i.Editor?.Title,
            fileSize: i.File_x0020_Size,
            fileType: i.File_x0020_Type,

            // dynamic bag
            fields: { ...i }
        };
    }

    // -------------------------------------------------------
    // GENERIC LIST ITEMS: IListItem[]
    // -------------------------------------------------------
    public async getListItems(listTitle: string): Promise<IListItemDto[]> {
        const items = await this.sp.web.lists
            .getByTitle(listTitle)
            .items
            .select("*")();

        return items.map(i => this.mapToListItem(i));
    }

    private mapToListItem(i: any): IListItemDto {
        return {
            id: i.Id,
            title: i.Title,
            fields: { ...i }
        };
    }

    // -------------------------------------------------------
    // LIST FIELDS: IFieldInfo[]
    // -------------------------------------------------------
    public async getListFields(listTitle: string): Promise<IFieldInfoDto[]> {
        const fields = await this.sp.web.lists
            .getByTitle(listTitle)
            .fields
            .select("InternalName", "Title", "TypeAsString", "Hidden", "ReadOnlyField")(); // execute

        return fields.map(f => ({
            internalName: f.InternalName,
            title: f.Title,
            type: f.TypeAsString,
            hidden: f.Hidden,
            readOnly: f.ReadOnlyField
        }));
    }

    // -------------------------------------------------------
    // CAML QUERY: IListItem[]
    // -------------------------------------------------------
    public async getItemsByCaml(listTitle: string, viewXml: string): Promise<IListItemDto[]> {
        const result = await this.sp.web.lists
            .getByTitle(listTitle)
            .renderListDataAsStream({ ViewXml: viewXml });

        return result.Row.map((i: any) => this.mapToListItem(i));
    }

    // -------------------------------------------------------
    // FOLDER ITEMS: IDocument[]
    // -------------------------------------------------------
    public async getFolderItems(serverRelativeFolderPath: string): Promise<IDocumentDto[]> {
        const files = await this.sp.web
            .getFolderByServerRelativePath(serverRelativeFolderPath)
            .files
            .expand("ListItemAllFields")
            .select(
                "Name",
                "ServerRelativeUrl",
                "TimeLastModified",
                "Length",
                "UniqueId",
                "ListItemAllFields/Id"
            )();

        return files.map((f: any) => ({
            id: f.ListItemAllFields?.Id ?? 0,
            name: f.Name,
            url: f.ServerRelativeUrl,
            modified: new Date(f.TimeLastModified),
            modifiedBy: "",
            fileSize: f.Length,
            fileType: f.Name.split(".").pop() ?? "",
            fields: { ...f }
        }));
    }
}
