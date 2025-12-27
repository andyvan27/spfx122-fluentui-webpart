export class CamlHelper {

    // -----------------------------
    // ORDER BY
    // -----------------------------
    public static buildOrderByXml(field: string, asc: boolean): string {
        return `
      <OrderBy>
        <FieldRef Name="${field}" Ascending="${asc ? "TRUE" : "FALSE"}" />
      </OrderBy>
    `;
    }

    // -----------------------------
    // WHERE (Contains)
    // -----------------------------
    public static buildWhereContainsXml(field: string, value: string): string {
        return `
      <Where>
        <Contains>
          <FieldRef Name="${field}" />
          <Value Type="Text">${value}</Value>
        </Contains>
      </Where>
    `;
    }

    // -----------------------------
    // FULL VIEW XML
    // -----------------------------
    public static buildViewXml(
        viewFieldNames: string[],
        whereXml?: string,
        orderByXml?: string,
        pageSize: number = 50
    ): string {

        const viewFieldsXml = viewFieldNames
            .map(n => `<FieldRef Name="${n}" />`)
            .join("");

        return `
            <View>
            <ViewFields>
                <FieldRef Name="ID" />
                <FieldRef Name="FileLeafRef" />
                <FieldRef Name="FileRef" />
                <FieldRef Name="Modified" />
                <FieldRef Name="Editor" />
                <FieldRef Name="SMTotalFileStreamSize" />
                <FieldRef Name="DocIcon" />
                <FieldRef Name="File_x0020_Type" />
                <FieldRef Name="FSObjType" />
                ${viewFieldsXml}
            </ViewFields>
            <Query>
                ${whereXml ?? ""}
                ${orderByXml ?? ""}
            </Query>
            <RowLimit Paged="TRUE">${pageSize}</RowLimit>
            </View>
        `;
    }
}
