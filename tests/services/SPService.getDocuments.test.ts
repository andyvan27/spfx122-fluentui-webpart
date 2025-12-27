import { SPService } from "../../src/services/SPService";
import { __finalCall } from "../mocks/pnp-sp.mock";

describe("SPService.getDocuments", () => {
  it("returns normalized document DTOs", async () => {
    const mockItems = [
      {
        Id: 10,
        FileLeafRef: "report.pdf",
        FileRef: "/sites/demo/Shared Documents/report.pdf",
        Modified: "2024-01-01T00:00:00Z",
        File_x0020_Size: 2048,
        File_x0020_Type: "pdf",
        Editor: { Title: "Andy" }
      }
    ];

    // Override the final "()"
    __finalCall.mockReturnValue(mockItems);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const service = new SPService({} as any);
    const result = await service.getDocuments("Documents");

    expect(result.length).toBe(1);
    expect(result[0].id).toBe(10);
    expect(result[0].name).toBe("report.pdf");
  });
});
