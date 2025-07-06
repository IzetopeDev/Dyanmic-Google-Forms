// Disclaimer: This AppScript program requires account permissions to read and write the related Google Forms and Google Sheets.

// When working locally with clasp or ts-node, uncomment imports:
// import { sheets, SITSettings, tables } from './00SheetsVar';
// import { form } from './01FormVar';

/** Test function for checking date formatting */
function test(): void {
  const dateObj = new Date();
  const custDateObj = new CustDate();
  const custDateObjInherit = new CustDate(dateObj);

  console.log('dateObj :>> ', dateObj);
  console.log('custDateObj :>> ', custDateObj);
  console.log('custDateObjInherit :>> ', custDateObjInherit);
  console.log('custDateObj.toCustFormat() :>> ', custDateObj.toCustFormat());
}

/**
 * Sorts an array based on ascending/descending order.
 * @param order +1 for ascending, -1 for descending.
 * @param inputArr Array to sort (number or string).
 */
function sortBy<T extends string | number>(order: number, inputArr: T[]): T[] {
  return inputArr.sort((a, b) => {
    if (a < b) return -1 * order;
    if (a > b) return 1 * order;
    return 0;
  });
}

/** Triggered on form submission */
function onFormSubmit(): void {
  const latestResponse = form.responses[form.responses.length - 1].getItemResponses();
  const category = latestResponse[0]?.getResponse();
  const subcategory = latestResponse[1]?.getResponse();

  console.log('Form response:', latestResponse.map(r => r.getResponse()));

  switch (category) {
    case "Infra":
      console.log('case: Infra');
      switch (subcategory) {
        case "Defect":
          tables.infraDefects.updateSimTable(latestResponse, 3);
          tables.infraDefects.sortTableByAlpahbet();
          tables.infraDefects.writeToSheet();
          form.updateJobQnList();
          break;

        case "Job":
          tables.infraJobs.updateSimTable(latestResponse, 3);
          tables.infraJobs.sortTableByAlpahbet();
          tables.infraJobs.writeToSheet();
          form.updateJobQnList();
          break;

        default:
          console.warn(`No subcategory match: ${subcategory}`);
          break;
      }
      break;

    case "Office Purchases":
      console.log('case: Office Purchases');
      // Uncomment if officePurchases table is defined
      // tables.officePurchases.updateSimTable(latestResponse, 1);
      // tables.officePurchases.writeToSheet();
      // form.updateJobQnList(); // or your updateJobChoices
      break;

    default:
      console.error("No valid category found in response.");
      break;
  }
}
