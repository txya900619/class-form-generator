function doGet() {
  return ContentService.createTextOutput("cc");
}

function doPost(e) {
  const param = e.parameter;
  const title = param.title;
  const semester = param.semester;
  const folderName = param.folderName;
  const courseStatement = param.courseStatement;
  const description = param.description;
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
    description,
    courseStatement,
    title
  );
  createFeedbackForm(folderName, currentFolder, title);
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
