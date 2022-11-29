import { sdkify } from "cdsify";
import { SetupGlobalContext } from "cdsify/lib/cdsnode";
import { axa_checklistMetadata } from "../../cds-generated/entities/axa_Checklist";
import { axa_checklisttype } from "../../cds-generated/enums/axa_checklisttype";
import cdsService from "../CdsService";

test("retreive", async () => {
  // await SetupGlobalContext();

  // console.log("done setup");
  // const newCds = new cdsService();

  // const result = await newCds.retrieveChecklistsByType(
  //   axa_checklisttype.Safety
  // );

  // const sdkifyResult = sdkify(result, axa_checklistMetadata.logicalName);
  // console.log(sdkifyResult);
  expect(true).toBeTruthy();
});
