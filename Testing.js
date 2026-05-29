function resubmitLastLinkedFormResponse() {
  const form = getLinkedForm_();
  const responses = form.getResponses();
  if (!responses.length) {
    throw new Error("The linked form has no responses to resubmit.");
  }

  const latestResponse = responses[responses.length - 1];
  const duplicateResponse = form.createResponse();
  latestResponse.getItemResponses().forEach((itemResponse) => {
    const response = createDuplicateItemResponse_(itemResponse);
    if (response) {
      duplicateResponse.withItemResponse(response);
    }
  });

  duplicateResponse.submit();
  SpreadsheetApp.getActive().toast("Resubmitted the latest linked form response.");
}

function getLinkedForm_() {
  const formUrl = SpreadsheetApp.getActive().getFormUrl();
  if (!formUrl) {
    throw new Error("This spreadsheet does not have a linked Google Form.");
  }
  return FormApp.openByUrl(formUrl);
}

function createDuplicateItemResponse_(itemResponse) {
  const item = itemResponse.getItem();
  const response = itemResponse.getResponse();

  if (response === null || response === undefined || response === "") {
    return null;
  }

  switch (item.getType()) {
    case FormApp.ItemType.TEXT:
      return item.asTextItem().createResponse(String(response));
    case FormApp.ItemType.PARAGRAPH_TEXT:
      return item.asParagraphTextItem().createResponse(String(response));
    case FormApp.ItemType.MULTIPLE_CHOICE:
      return item.asMultipleChoiceItem().createResponse(String(response));
    case FormApp.ItemType.CHECKBOX:
      return item.asCheckboxItem().createResponse(asArray_(response));
    case FormApp.ItemType.LIST:
      return item.asListItem().createResponse(String(response));
    case FormApp.ItemType.DATE:
      return item.asDateItem().createResponse(asDateForForm_(response));
    case FormApp.ItemType.DATETIME:
      return item.asDateTimeItem().createResponse(asDateForForm_(response));
    case FormApp.ItemType.TIME:
      return createTimeResponse_(item.asTimeItem(), response);
    case FormApp.ItemType.DURATION:
      return createDurationResponse_(item.asDurationItem(), response);
    case FormApp.ItemType.SCALE:
      return item.asScaleItem().createResponse(Number(response));
    case FormApp.ItemType.GRID:
      return item.asGridItem().createResponse(asArray_(response));
    case FormApp.ItemType.CHECKBOX_GRID:
      return item.asCheckboxGridItem().createResponse(response);
    default:
      throw new Error(
        `Cannot resubmit response for unsupported item type ${item.getType()} on "${item.getTitle()}".`
      );
  }
}

function asArray_(value) {
  return Array.isArray(value) ? value : [value];
}

function asDateForForm_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return value;
  }
  const date = new Date(value);
  if (isNaN(date)) {
    throw new Error(`Unable to convert "${value}" to a date for form resubmission.`);
  }
  return date;
}

function asDurationParts_(value) {
  if (Array.isArray(value)) return value;
  const parts = String(value)
    .split(":")
    .map((part) => Number(part));
  if (parts.some((part) => !isFinite(part))) {
    throw new Error(`Unable to convert "${value}" to a duration for form resubmission.`);
  }
  return parts;
}

function createTimeResponse_(item, value) {
  const parts = asDurationParts_(value);
  return item.createResponse(parts[0] || 0, parts[1] || 0);
}

function createDurationResponse_(item, value) {
  const parts = asDurationParts_(value);
  return item.createResponse(parts[0] || 0, parts[1] || 0, parts[2] || 0);
}
