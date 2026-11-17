import { describe, expect, it } from "vitest";
import { fieldcapJobIdFromUrl, fieldcapJobNumberFromTitle, metadataEntities, stationFields } from "../../fieldcap-sync/src/index";

describe("fieldcapJobIdFromUrl", () => {
  it("reads the job from the route, then from an OData filter", () => {
    expect(fieldcapJobIdFromUrl("https://fieldcap.phxtech.com/Jobs/4242/Surveys")).toBe("4242");
    expect(fieldcapJobIdFromUrl("https://fieldcap.phxtech.com/jobs/77")).toBe("77");
    expect(fieldcapJobIdFromUrl("https://fieldcap-cdn.phxtech.com/odata/ToolAssemblies?$filter=ClientJobId%20eq%20913")).toBe("913");
    expect(fieldcapJobIdFromUrl("ClientJobId eq 55")).toBe("55");
  });

  it("is blank off a job page", () => {
    expect(fieldcapJobIdFromUrl("https://fieldcap.phxtech.com/Dashboard")).toBe("");
    expect(fieldcapJobIdFromUrl("https://fieldcap.phxtech.com/Jobs/abc")).toBe("");
    expect(fieldcapJobIdFromUrl("")).toBe("");
  });
});

describe("fieldcapJobNumberFromTitle", () => {
  it("reads the Job Number between the comma and the dash", () => {
    expect(fieldcapJobNumberFromTitle("Northwind Energy Ltd., 20803 - NWE HZ WILLOWBEND 04-22-68-08")).toBe("20803");
    expect(fieldcapJobNumberFromTitle("Acme,  1234 - WELL")).toBe("1234");
  });

  it("is blank when the title has no Job Number", () => {
    expect(fieldcapJobNumberFromTitle("FieldCap")).toBe("");
    expect(fieldcapJobNumberFromTitle("Acme, 35783")).toBe("");
    expect(fieldcapJobNumberFromTitle("Acme, 3578311 - WELL")).toBe("");
    expect(fieldcapJobNumberFromTitle("")).toBe("");
  });
});

describe("metadataEntities", () => {
  const xml = `<?xml version="1.0"?><edmx:Edmx><edmx:DataServices><Schema>
    <EntityType Name="SurveySheetEntry"><Key><PropertyRef Name="SurveySheetEntryId"/></Key>
      <Property Name="SurveySheetEntryId" Type="Edm.Int32" Nullable="false"/>
      <Property Name="MeasuredDepth" Type="Edm.Double"/>
      <Property Name="Inclination" Type="Edm.Double"/>
      <Property Name="Azimuth" Type="Edm.Double"/>
      <Property Name="TrueVerticalDepth" Type="Edm.Double"/>
      <NavigationProperty Name="SurveySheet" Type="X.SurveySheet"/>
    </EntityType>
    <EntityType Name="ActivityLog"><Property Name="StartDateTime" Type="Edm.DateTimeOffset"/><Property Name="Duration" Type="Edm.Double"/></EntityType>
  </Schema></edmx:DataServices></edmx:Edmx>`;

  it("lists each entity's fields in document order", () => {
    expect(metadataEntities(xml)).toEqual([
      { entity: "SurveySheetEntry", fields: ["SurveySheetEntryId", "MeasuredDepth", "Inclination", "Azimuth", "TrueVerticalDepth", "SurveySheet"] },
      { entity: "ActivityLog", fields: ["StartDateTime", "Duration"] },
    ]);
  });

  it("finds the station fields, or null when FieldCap has none", () => {
    expect(stationFields(metadataEntities(xml)[0]!.fields)).toEqual({ md: "MeasuredDepth", inc: "Inclination", azm: "Azimuth", tvd: "TrueVerticalDepth", bit: null });
    expect(stationFields(["MD", "Inc", "Azi", "TVD"])).toEqual({ md: "MD", inc: "Inc", azm: "Azi", tvd: "TVD", bit: null });
    expect(stationFields(["SurveyDateTime", "CourseLength", "SlidingDistance"])).toEqual({ md: null, inc: null, azm: null, tvd: null, bit: null });
  });

  it("finds the bit depth beside the survey depth", () => {
    expect(stationFields(["BitDepth", "SurveyDepth", "Inclination", "Azimuth"])).toMatchObject({ md: "SurveyDepth", bit: "BitDepth" });
    expect(stationFields(["Depth", "Bit_Depth", "Inc", "Azm"])).toMatchObject({ md: "Depth", bit: "Bit_Depth" });
  });
});
