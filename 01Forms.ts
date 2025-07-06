class FormClass {
  obj: GoogleAppsScript.Forms.Form;
  id: string;
  items: ItemClass;
  responses: GoogleAppsScript.Forms.FormResponse[];
  latestResponse: GoogleAppsScript.Forms.ItemResponse[];

  constructor(googleFormUrl: string) {
    console.info("building FormClass");
    this.obj = FormApp.openByUrl(googleFormUrl);
    this.id = this.obj.getId();
    this.items = new ItemClass(this.obj);
    this.responses = this.obj.getResponses();
    this.latestResponse = this.responses[this.responses.length - 1].getItemResponses();
  }

  getJobQnID(enableLog = false, response: GoogleAppsScript.Forms.ItemResponse[] = this.latestResponse, dataStartIndex = 2): number | null {
    console.info('FormClass.getJobQnID() called');

    const findQnFrom = (index: number, items: ItemClass): number | null => {
      if (enableLog) {
        console.log('findQnFrom() called', { index, items });
      }
      if (index == null || items == null) return null;

      for (let i = index; i < items.obj.length; i++) {
        const itemDetails = items.getDetails(enableLog, i, i)[0];
        if (itemDetails.type === 'MULTIPLE_CHOICE') return itemDetails.id;
      }
      return null;
    };

    const answers = response.map((item) => item.getResponse());
    const categories = this.items.getDetails(enableLog).filter((item) => item.type === "PAGE_BREAK");

    for (let category of categories) {
      const keywords = category.title.split(' ');
      const categoryKeywords = answers[dataStartIndex - 1];

      let matchedKeys = keywords.filter(k => k === categoryKeywords).length;

      if (matchedKeys > 0) {
        const jobQnID = findQnFrom(category.index, this.items);
        if (enableLog) console.log('jobQnID :>> ', jobQnID);
        return jobQnID;
      } else if (enableLog) {
        console.warn('No matched keywords in category:', category.title);
      }
    }
    return null;
  }

  updateJobQnList(enableLog = false, response: GoogleAppsScript.Forms.ItemResponse[] = this.latestResponse): void | null {
    console.info('FormClass.updateJobQnList() called');

    const jobQnID = this.getJobQnID(enableLog, response);
    if (jobQnID == null) {
      console.warn('No job Qn ID found!');
      return null;
    }

    const jobQnItem = this.obj.getItemById(jobQnID);
    const categoryKeyword = jobQnItem.getTitle().split(' ')[1];

    let selectedTable: UnlimitedTable;
    switch (categoryKeyword) {
      case 'Job':
        selectedTable = tables.infraJobs;
        break;
      case 'Defect':
        selectedTable = tables.infraDefects;
        break;
      default:
        console.warn('Unknown category keyword:', categoryKeyword);
        return null;
    }

    const newJobQnList = selectedTable.dataSets[0].filter((value) => value !== '');
    jobQnItem.asMultipleChoiceItem().setChoiceValues(newJobQnList);
  }
}

class ItemClass {
  obj: GoogleAppsScript.Forms.Item[];

  constructor(form: GoogleAppsScript.Forms.Form) {
    console.info('building ItemClass');
    this.obj = form.getItems();
  }

  getDetails(enableLog = false, startIndex = 0, endIndex = this.obj.length - 1): {
    title: string;
    index: number;
    id: number;
    type: string;
  }[] {
    if (enableLog) {
      console.info('form.getItemDetails() logging enabled', { startIndex, endIndex });
    }

    const itemDetails: {
      title: string;
      index: number;
      id: number;
      type: string;
    }[] = [];

    for (let i = startIndex; i <= endIndex; i++) {
      const item = this.obj[i];
      itemDetails.push({
        title: item.getTitle(),
        index: i,
        id: item.getId(),
        type: item.getType().toString()
      });

      if (enableLog) {
        console.log(`Item ${i}:`, item);
      }
    }

    if (enableLog) console.log('itemDetails:', itemDetails);
    return itemDetails;
  }

  getChoices(
    identifier: number | string | null = null,
    enableLog = false
  ): string[] | undefined {
    console.info('form.getItemChoices() called');

    let item: GoogleAppsScript.Forms.Item | undefined;

    if (identifier === null) {
      console.warn('Identifier is null. Please provide a valid identifier.');
      return;
    }

    if (typeof identifier === 'number') {
      if (identifier < 100) {
        item = this.obj[identifier]; // Index
      } else {
        item = this.obj.find(it => it.getId() === identifier); // ID
      }
    } else if (typeof identifier === 'string') {
      const details = this.getDetails();
      const match = details.find(d => d.title === identifier);
      if (match) item = this.obj[match.index];
    }

    if (!item) {
      console.warn('No item found for identifier:', identifier);
      return;
    }

    let itemChoices: string[] = [];

    switch (item.getType()) {
      case FormApp.ItemType.MULTIPLE_CHOICE:
        itemChoices = item.asMultipleChoiceItem().getChoices().map(choice => choice.getValue());
        break;
      case FormApp.ItemType.LIST:
        itemChoices = item.asListItem().getChoices().map(choice => choice.getValue());
        break;
      default:
        console.warn("Invalid item type for choice extraction");
        break;
    }

    if (enableLog) console.log('itemChoices:', itemChoices);
    return itemChoices;
  }
}

const form = new FormClass('')
