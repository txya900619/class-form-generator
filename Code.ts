function doGet() {
  return ContentService.createTextOutput("cc");
}

function doPost(e) {
  const param = e.parameter;
  const formTitle = param.formTitle;
  const semester = param.semester;
  const folderName = param.folderName;
  const courseStatement = param.courseStatement;
  const classInformation = param.classInformation;
  const signUpFormDescription = param.signUpFormDescription;
  const maxNumberOfStudent = param.maxNumberOfStudent;
  const numberOfWaitingList = param.numberOfWaitingList;
  const rootFolder = DriveApp.getFolderById(
    "1PnQObOfZwq3IYleIDkUS9wCGUXrnmhe6"
  );
  const semesterFolder = rootFolder.getFoldersByName(semester).next();
  const findFolder = semesterFolder.getFoldersByName(folderName);
  let currentFolder: GoogleAppsScript.Drive.Folder;
  const date = new Date(param.date);
  if (findFolder.hasNext()) {
    currentFolder = findFolder.next();
  } else {
    currentFolder = semesterFolder.createFolder(folderName);
  }
  const configSpreadsheetID = createConfigSpreadsheet(currentFolder);
  setConfigSpreadsheet(
    configSpreadsheetID,
    maxNumberOfStudent,
    numberOfWaitingList
  );
  const signUpInfo = createSignUpForm(
    folderName,
    currentFolder,
    signUpFormDescription,
    courseStatement,
    formTitle,
    classInformation
  );
  createFeedbackForm(folderName, currentFolder, formTitle);

  ScriptApp.newTrigger("sendPriorNotificationEmail")
    .timeBased()
    .at(date)
    .inTimezone("Asia/Taipei")
    .create();
}

function setProperty(
  signUpInfo: {
    signUpSpreadsheetID: string;
    priorNotificationEmailDocsID: string;
  },
  formTitle: string,
  semester: string
) {
  const properties = PropertiesService.getScriptProperties();
  const current = properties.getProperty("current");
  if (!current) {
    properties.setProperty(semester + formTitle, JSON.stringify(signUpInfo));
  }
  let tempCurrent: string = current;
  while (true) {
    const tempCurrentData = JSON.parse(properties.getProperty(tempCurrent));
    if (!tempCurrentData.next) {
      tempCurrentData.next = semester + formTitle;
      properties.setProperty(current, JSON.stringify(tempCurrentData));
      break;
    }
    tempCurrent = tempCurrentData.next;
  }
}

function createConfigSpreadsheet(
  currentFolder: GoogleAppsScript.Drive.Folder
): string {
  const spreadsheetID = createSpreadsheetInFolder("設定", currentFolder);
  return spreadsheetID;
}

function setConfigSpreadsheet(
  spreadSheetID: string,
  maxNumberOfStudent: number,
  numberOfWaitingList: number
) {
  const sheet = SpreadsheetApp.openById(spreadSheetID).getSheets()[0];
  sheet.appendRow(["正取人數上限", "備取人數"]);
  sheet.appendRow([maxNumberOfStudent, numberOfWaitingList]);
}
//return signUpSpreadsheetID and priorNotificationEmailDocsID
function createSignUpForm(
  folderName: string,
  currentFolder: GoogleAppsScript.Drive.Folder,
  description: string,
  courseStatement: string,
  title: string,
  classInformation: string
): { signUpSpreadsheetID: string; priorNotificationEmailDocsID: string } {
  const formID = createFormInFolder(folderName + " 報名表單", currentFolder);
  const spreadsheetID = createSpreadsheetInFolder(
    folderName + " 報名表單（回應）",
    currentFolder
  );
  setSignUpFormItem(formID, description, courseStatement, title, spreadsheetID);
  setSuccessEmail(currentFolder, title, classInformation);
  setWaitingListEmail(currentFolder, title);
  const priorNotificationEmailDocsID = setPriorNotificationEmail(
    currentFolder,
    title,
    classInformation
  );

  ScriptApp.newTrigger("SignUpFormOnSubmit")
    .forSpreadsheet(SpreadsheetApp.openById(spreadsheetID))
    .onFormSubmit()
    .create();

  return { signUpSpreadsheetID: spreadsheetID, priorNotificationEmailDocsID };
}

function createFeedbackForm(
  folderName: string,
  currentFolder: GoogleAppsScript.Drive.Folder,
  title: string
) {
  const formID = createFormInFolder(folderName + " 回饋表單", currentFolder);
  const spreadsheetID = createSpreadsheetInFolder(
    folderName + " 回饋表單（回覆）",
    currentFolder
  );
  setFeedbackItem(formID, title, spreadsheetID);
}

function createSpreadsheetInFolder(
  name: string,
  folder: GoogleAppsScript.Drive.Folder
): string {
  const tempSpreadsheetID = SpreadsheetApp.create(name).getId();
  const tempSpreadsheetFile = DriveApp.getFileById(tempSpreadsheetID);
  const spreadsheetID = tempSpreadsheetFile.makeCopy(name, folder).getId();
  tempSpreadsheetFile.setTrashed(true);
  return spreadsheetID;
}

function createFormInFolder(
  name: string,
  folder: GoogleAppsScript.Drive.Folder
): string {
  const tempFormID = FormApp.create(name).getId();
  const tempFormFile = DriveApp.getFileById(tempFormID);
  const formID = tempFormFile.makeCopy(name, folder).getId();
  tempFormFile.setTrashed(true);
  return formID;
}

function createDocsInFolder(
  name: string,
  folder: GoogleAppsScript.Drive.Folder
): string {
  const tempDocsID = DocumentApp.create(name).getId();
  const tempDocsFile = DriveApp.getFileById(tempDocsID);
  const docsID = tempDocsFile.makeCopy(name, folder).getId();
  tempDocsFile.setTrashed(true);
  return docsID;
}

function setSignUpFormItem(
  formID: string,
  description: string,
  courseStatement: string,
  title: string,
  spreadsheetID: string
): void {
  const form = FormApp.openById(formID);
  form.setTitle(title);
  form.setDescription(description);
  form.setCollectEmail(true);
  form.addTextItem().setTitle("班級").setRequired(true);
  form.addTextItem().setTitle("學號").setRequired(true);
  form.addTextItem().setTitle("姓名").setRequired(true);
  form.addSectionHeaderItem().setTitle("課程聲明").setHelpText(courseStatement);
  let choiceItem = form.addMultipleChoiceItem();
  choiceItem
    .setTitle("已看過課程聲明")
    .setChoices([choiceItem.createChoice("確認")])
    .setRequired(true);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetID);
}

function setFeedbackItem(formID: string, title: string, spreadsheetID: string) {
  const form = FormApp.openById(formID);
  form.setTitle("Feedback - " + title + " by NPC");
  form.setDescription(
    "感謝大家今天的蒞臨！\n你的每個意見都能夠讓 NPC 變得更好！"
  );
  form
    .addScaleItem()
    .setTitle("對於今天課程的滿意度")
    .setBounds(1, 5)
    .setRequired(true);
  form.addScaleItem().setTitle("課程難易度").setBounds(1, 5).setRequired(true);
  form
    .addScaleItem()
    .setTitle("講師說話速度")
    .setBounds(1, 5)
    .setLabels("太慢", "太快")
    .setRequired(true);
  let choiceItem = form.addMultipleChoiceItem();
  choiceItem
    .setTitle("未來希望 NPC 開放哪些課程？")
    .setChoices([
      choiceItem.createChoice("Python 應用"),
      choiceItem.createChoice("機器學習 (Machine Learning)"),
      choiceItem.createChoice("資訊安全 (CTF, 搶旗遊戲)"),
      choiceItem.createChoice("Unity (遊戲製作, 能製作 2D &3D 的遊戲)"),
      choiceItem.createChoice("網頁進階 (框架)"),
      choiceItem.createChoice(
        "Android App (如 TAT , 不是遊戲 App！不是遊戲 App！不是遊戲 App！)"
      ),
      choiceItem.createChoice(
        'Swift / iOS App (應用在 iOS / MacOS 等等的程式語言) (需要自備 Mac 筆電 or " 黑蘋果")'
      ),
    ])
    .showOtherOption(true)
    .setRequired(true);
  form.addParagraphTextItem().setTitle("有什麼其他的建議給我們嗎？");
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetID);
}

// return docsID
function addDocumentWithFolderAndNameAndHeaderAndFooter(
  currentFolder: GoogleAppsScript.Drive.Folder,
  name: string,
  header: string,
  footer: string
): string {
  const docsID = createDocsInFolder(name, currentFolder);
  const docs = DocumentApp.openById(docsID);
  docs.addHeader().setText(header);
  docs.addFooter().setText(footer);

  return docsID;
}

function setSuccessEmail(
  currentFolder: GoogleAppsScript.Drive.Folder,
  formTitle: string,
  classInformation: string
) {
  const subject = `【 報名成功通知 】${formTitle} by NPC 北科程式設計研究社【正取】`;
  const body =
    "Hi, \n\n" +
    `感謝您報名 NPC 北科程式設計研究社 ${formTitle}\n` +
    "在課程開始前一天，我們會再次寄信提醒您！\n\n" +
    "另外，由於資源寶貴，若臨時未能前來請您務必及早回信告知，讓備取學員得以遞補，謝謝您。\n\n" +
    classInformation +
    "\n若有任何疑問，歡迎隨時連絡我們。\n" +
    "期待在課程與您相見:)\n\n" +
    "Best regards,\n" +
    "NPC 北科程式設計研究社";

  addDocumentWithFolderAndNameAndHeaderAndFooter(
    currentFolder,
    "正取 email",
    subject,
    body
  );
  // successEmail.set(spreadSheetID, { subject, body });
}

function setWaitingListEmail(
  currentFolder: GoogleAppsScript.Drive.Folder,
  formTitle: string
) {
  const subject = `【 報名成功通知 】${formTitle} by NPC 北科程式設計研究社【備取】`;
  const body =
    "Hi, \n\n" +
    `感謝您報名 NPC 北科程式設計研究社 ${formTitle}\n` +
    "為了保證上課品質，我們人數已達到上限，如果有人放棄資格，我們會儘速通知您！\n\n" +
    "若有任何疑問，歡迎隨時連絡我們。\n" +
    "由衷感謝您:)\n\n" +
    "Best regards,\n" +
    "NPC 北科程式設計研究社";

  addDocumentWithFolderAndNameAndHeaderAndFooter(
    currentFolder,
    "備取 email",
    subject,
    body
  );
}

// return priorNotificationEmailDocsID
function setPriorNotificationEmail(
  currentFolder: GoogleAppsScript.Drive.Folder,
  formTitle: string,
  classInformation: string
): string {
  const subject = `【 ${formTitle} 】行前通知信 by NPC `;
  const body =
    "您好, \n\n" +
    `提醒您，【 ${formTitle} 】即將於明天晚上舉辦！\n` +
    "另外，由於資源寶貴，若臨時未能前來請您務必及早回信告知，讓備取學員得以遞補，謝謝您。\n\n" +
    classInformation +
    "\n若有任何疑問，歡迎隨時連絡我們。\n" +
    "期待在課程與您相見:)\n\n" +
    "Best regards,\n" +
    "NPC 北科程式設計研究社";

  return addDocumentWithFolderAndNameAndHeaderAndFooter(
    currentFolder,
    "行前 email",
    subject,
    body
  );
}

function getEmailByName(
  currentFolder: GoogleAppsScript.Drive.Folder,
  name: string
): Email {
  const docsID = currentFolder.getFilesByName(name).next().getId();
  const docs = DocumentApp.openById(docsID);

  return {
    subject: docs.getHeader().getText(),
    body: docs.getFooter().getText(),
  };
}

function getMaxNumberOfStudent(
  currentFolder: GoogleAppsScript.Drive.Folder
): number {
  const spreadsheetID = currentFolder.getFilesByName("設定").next().getId();
  const sheet = SpreadsheetApp.openById(spreadsheetID).getSheets()[0];
  const maxNumberOfStudent: number = Number(sheet.getRange(2, 1).getValue());
  return maxNumberOfStudent;
}

function getNumberOfWaitingList(
  currentFolder: GoogleAppsScript.Drive.Folder
): number {
  const spreadsheetID = currentFolder.getFilesByName("設定").next().getId();
  const sheet = SpreadsheetApp.openById(spreadsheetID).getSheets()[0];
  const numberOfWaitingList: number = Number(sheet.getRange(2, 2).getValue());
  return numberOfWaitingList;
}

function SignUpFormOnSubmit(e: GoogleAppsScript.Events.SheetsOnFormSubmit) {
  const spreadsheetID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const currentFolder = DriveApp.getFileById(spreadsheetID).getParents().next();
  const range = e.range;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const maxNumberOfStudent = getMaxNumberOfStudent(currentFolder);
  const numberOfWaitingList = getNumberOfWaitingList(currentFolder);
  const spreadsheetLastRow = range.getLastRow();
  Logger.log(maxNumberOfStudent);
  Logger.log(numberOfWaitingList);
  Logger.log(spreadsheetLastRow);
  if (spreadsheetLastRow - 1 <= maxNumberOfStudent) {
    sendSuccessEmail(currentFolder, range, sheet);
  } else {
    sendWaitingList(currentFolder, range, sheet);
  }
  if (spreadsheetLastRow - 1 === maxNumberOfStudent + numberOfWaitingList) {
    const activeForm = FormApp.openByUrl(
      SpreadsheetApp.getActiveSpreadsheet().getFormUrl()
    );
    activeForm.setCustomClosedFormMessage(
      "很抱歉，為保證上課品質，報名人數已滿，請等待之後的相關課程"
    );
    activeForm.setAcceptingResponses(false);
  }
}

function sendSuccessEmail(
  currentFolder: GoogleAppsScript.Drive.Folder,
  range: GoogleAppsScript.Spreadsheet.Range,
  sheet: GoogleAppsScript.Spreadsheet.Sheet
) {
  const emailAddress: string = sheet.getRange(range.getLastRow(), 2).getValue();

  const email: Email = getEmailByName(currentFolder, "正取 email");

  MailApp.sendEmail(emailAddress, email.subject, email.body);
}

function sendWaitingList(
  currentFolder: GoogleAppsScript.Drive.Folder,
  range: GoogleAppsScript.Spreadsheet.Range,
  sheet: GoogleAppsScript.Spreadsheet.Sheet
) {
  const emailAddress: string = sheet.getRange(range.getLastRow(), 2).getValue();

  const email = getEmailByName(currentFolder, "備取 email");

  MailApp.sendEmail(emailAddress, email.subject, email.body);
}

function sendPriorNotificationEmail() {
  const properties = PropertiesService.getScriptProperties();
  const current = properties.getProperty("current");
  const currentData: {
    signUpSpreadsheetID: string;
    priorNotificationEmailDocsID: string;
    next: string;
  } = JSON.parse(properties.getProperty(current));
  const sheet = SpreadsheetApp.openById(
    currentData.signUpSpreadsheetID
  ).getSheets()[0];
  const emailDocs = DocumentApp.openById(
    currentData.priorNotificationEmailDocsID
  );
  const subject = emailDocs.getHeader().getText();
  const emailBody = emailDocs.getFooter().getText();
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    const emailAddr = sheet.getSheetValues(i, 2, 1, 1);
    MailApp.sendEmail(emailAddr[0][0], subject, emailBody);
  }

  properties.setProperty("current", currentData.next);
  properties.deleteProperty(current);
}
interface Email {
  subject: string;
  body: string;
}
