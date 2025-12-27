import { spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/files/item";

import { ISPService } from "./ISPService";
import { IDocumentDto } from "../dtos/IDocumentDto";
import { IFieldInfoDto } from "../dtos/IFieldInfoDto";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFieldInfo } from "@pnp/sp/fields";

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
                "File/Length",
                "File/Name",
                "*"
            )
            .expand("Editor", "File")();

        return items.map(i => this.mapToDocument(i));
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private mapToDocument(i: any): IDocumentDto {
        return {
            id: i.Id,
            name: i.FileLeafRef,
            url: i.FileRef,
            modified: new Date(i.Modified),
            modifiedBy: i.Editor?.Title,
            fileSize: i.File?.Length ?? 0,
            fileType: i.FileLeafRef.includes('.')
                ? i.FileLeafRef.split('.').pop()
                : "unknown",

            // dynamic bag
            fields: { ...i }
        };
    }

    public async *getDocumentsPaged(
        listTitle: string,
        viewFieldNames: string[],
        pageSize: number = 100
    ): AsyncGenerator<IDocumentDto[], void, unknown> {

        const list = this.sp.web.lists.getByTitle(listTitle);

        const items = list.items
            .select(
                "Id",
                "FileLeafRef",
                "FileRef",
                "Modified",
                "Editor/Title",
                "File/Length",
                "File/Name",
                ...viewFieldNames
            )
            .expand("Editor", "File")
            .top(pageSize);

        for await (const page of items) {
            // ⭐ Convert raw SP items → IDocumentDto[]
            yield page.map(i => this.mapToDocument(i));
        }
    }

    // -------------------------------------------------------
    // LIST FIELDS: IFieldInfo[]
    // -------------------------------------------------------
    public async getListFields(listTitle: string, viewName?: string): Promise<IFieldInfoDto[]> {
        const list = this.sp.web.lists.getByTitle(listTitle);

        // 1. Get the view (default or named)
        let view;

        if (viewName) {
            view = list.views.getByTitle(viewName);
        } else {
            const defaultViewInfo = await list.defaultView.select("Id")();
            view = list.views.getById(defaultViewInfo.Id);
        }

        // 2. Get internal field names in the view (ordered)
        const viewFieldNames: string[] = (await view.fields()).Items;

        // 3. Fetch ALL fields (REST cannot filter by IN)
        const allFields = await list.fields
            .select("InternalName", "Title", "TypeAsString", "Hidden", "ReadOnlyField")();

        // 4. Filter in JS to preserve view order
        const filtered = viewFieldNames
            .map(name => allFields.find(f => f.InternalName === name))
            .filter((f): f is IFieldInfo => f !== undefined);

        // 5. Map to DTO
        return filtered.map(f => ({
            internalName: f.InternalName,
            title: f.Title,
            type: f.TypeAsString,
            hidden: f.Hidden,
            readOnly: f.ReadOnlyField
        }));
    }

    // -------------------------------------------------------
    // Stream pages: IDocumentDto[]
    // -------------------------------------------------------
    public async *getDocumentsStreamPaged(
        listTitle: string,
        viewXml: string
    ): AsyncGenerator<IDocumentDto[], void, unknown> {

        const list = this.sp.web.lists.getByTitle(listTitle);

        let position: string | undefined = undefined;

        while (true) {

            const result = await list.renderListDataAsStream({
                ViewXml: viewXml,
                Paging: position
            });

            // ⭐ Cast to any because PnPjs typing is incomplete
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const raw = result as any;

            const rows = raw?.Row ?? [];
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const dtos = rows.map((r: any) => this.mapStreamRowToDocument(r));

            yield dtos;

            const next = raw?.NextHref;
            position = next.startsWith("?") ? next.substring(1) : next;

            if (!position) break;
        }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private mapStreamRowToDocument(r: any): IDocumentDto {
        return {
            id: r.ID,
            name: r.FileLeafRef,
            url: r.FileRef,
            modified: new Date(r.Modified),
            modifiedBy: r.Editor?.[0]?.title ?? "",
            fileSize: r.SMTotalFileStreamSize ?? 0,
            fileType: r.File_x0020_Type ?? "unknown",

            // ⭐ dynamic bag: contains ALL fields (raw + formatted)
            fields: { ...r }
        };
    }
}
