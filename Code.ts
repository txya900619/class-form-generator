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
  const rootFolder = DriveApp.getFolderById(
    "1PnQObOfZwq3IYleIDkUS9wCGUXrnmhe6"
  );
  const semesterFolder = rootFolder.getFoldersByName(semester).next();
  const findFolder = semesterFolder.getFoldersByName(folderName);
  let currentFolder: GoogleAppsScript.Drive.Folder;
  if (findFolder.hasNext()) {
    currentFolder = findFolder.next();
  } else {
    currentFolder = semesterFolder.createFolder(folderName);
  }
  createSignUpForm(
    folderName,
    currentFolder,
    signUpFormDescription,
    courseStatement,
    formTitle
  );
  createFeedbackForm(folderName, currentFolder, formTitle);
}

function createSignUpForm(
  folderName: string,
  currentFolder: GoogleAppsScript.Drive.Folder,
  description: string,
  courseStatement: string,
  title: string
) {
  const formID = createFormInFolder(folderName + " 報名表單", currentFolder);
  const spreadsheetsID = createSpreadsheetsInFolder(
    folderName + " 報名表單（回應）",
    currentFolder
  );
  setSignUpFormItem(
    formID,
    description,
    courseStatement,
    title,
    spreadsheetsID
  );
}

function createFeedbackForm(
  folderName: string,
  currentFolder: GoogleAppsScript.Drive.Folder,
  title: string
) {
  const formID = createFormInFolder(folderName + " 回饋表單", currentFolder);
  const spreadSheetID = createSpreadsheetsInFolder(
    folderName + " 回饋表單（回覆）",
    currentFolder
  );
  setFeedbackItem(formID, title, spreadSheetID);
}

function createSpreadsheetsInFolder(
  name: string,
  folder: GoogleAppsScript.Drive.Folder
): string {
  const tempSpreadSheetID = SpreadsheetApp.create(name).getId();
  const tempSpreadSheetFile = DriveApp.getFileById(tempSpreadSheetID);
  const spreadSheetID = tempSpreadSheetFile.makeCopy(name, folder).getId();
  tempSpreadSheetFile.setTrashed(true);
  return spreadSheetID;
}

function createFormInFolder(
  name: string,
  folder: GoogleAppsScript.Drive.Folder
): string {
  const tempFormID = FormApp.create(name).getId();
  const tempFormFile = DriveApp.getFileById(tempFormID);
  const formID = tempFormFile.makeCopy(tempFormFile.getName(), folder).getId();
  tempFormFile.setTrashed(true);
  return formID;
}

function setSignUpFormItem(
  formID: string,
  description: string,
  courseStatement: string,
  title: string,
  spreadsheetsID: string
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

  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetsID);
}

function setFeedbackItem(formID: string, title: string, spreadSheetID: string) {
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
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadSheetID);
}

function getSuccessEmailBody(formTitle: string, emailBody: string): string {
  return (
    "Hi, \n\n" +
    `感謝您報名 NPC 北科程式設計研究社 ${formTitle}\n` +
    "在課程開始前一天，我們會再次寄信提醒您！\n\n" +
    "另外，由於資源寶貴，若臨時未能前來請您務必及早回信告知，讓備取學員得以遞補，謝謝您。\n\n" +
    emailBody +
    "\n若有任何疑問，歡迎隨時連絡我們。\n" +
    "期待在課程與您相見:)\n\n" +
    "Best regards,\n" +
    "NPC 北科程式設計研究社"
  );
}

function getSuccessEmailSubject(formTitle: string): string {
  return `【 報名成功通知 】${formTitle} by NPC 北科程式設計研究社【正取】`;
}

function getWaitingListEmailBody(formTitle: string): string {
  return (
    "Hi, \n\n" +
    `感謝您報名 NPC 北科程式設計研究社 ${formTitle}\n` +
    "為了保證上課品質，我們人數已達到上限，如果有人放棄資格，我們會儘速通知您！\n\n" +
    "若有任何疑問，歡迎隨時連絡我們。\n" +
    "由衷感謝您:)\n\n" +
    "Best regards,\n" +
    "NPC 北科程式設計研究社"
  );
}

function getWaitingListEmailSubject(formTitle: string): string {
  return `【 報名成功通知 】${formTitle} by NPC 北科程式設計研究社【備取】`;
}

function priorNotificationEmailBody(
  formTitle: string,
  classInformation: string
): string {
  return (
    "您好, \n\n" +
    `提醒您，【 ${formTitle} 】即將於明天晚上舉辦！\n` +
    "另外，由於資源寶貴，若臨時未能前來請您務必及早回信告知，讓備取學員得以遞補，謝謝您。\n\n" +
    classInformation +
    "\n若有任何疑問，歡迎隨時連絡我們。\n" +
    "期待在課程與您相見:)\n\n" +
    "Best regards,\n" +
    "NPC 北科程式設計研究社"
  );
}

function priorNotificationEmailSubject(formTitle: string): string {
  return `【 ${formTitle} 】行前通知信 by NPC `;
}
