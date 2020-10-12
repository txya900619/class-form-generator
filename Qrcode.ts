const folderID = "0B0IVnu5YlOTNZzZsV21LLUd2VHc";

function doGet(e: GoogleAppsScript.Events.DoGet) {
  const password: string = e.parameter["password"];
  if (password !== "cchahacc") {
    throw "wrong authToken!";
  }
  const subFolders = DriveApp.getFolderById(folderID).getFolders();
  let semesters: semester[] = [];
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    const subFolderName = subFolder.getName();
    const subSubFolders = subFolder.getFolders();
    let activeCourses: course[] = [];
    while (subSubFolders.hasNext()) {
      const subSubFolder = subSubFolders.next();
      const subSubFolderName = subSubFolder.getName();
      const settingIter = subSubFolder.getFilesByName("設定");
      if (!settingIter.hasNext()) {
        continue;
      }
      const spreadsheetsID = settingIter.next().getId();
      if (
        SpreadsheetApp.openById(spreadsheetsID)
          .getSheets()[0]
          .getRange(2, 3)
          .getValue() != ""
      ) {
        continue;
      }
      activeCourses.push({
        name: subSubFolderName,
        spreadsheetsID: spreadsheetsID
      });
    }
    if (activeCourses.length > 0) {
      semesters.push({ name: subFolderName, activeCourses });
    }
  }
  Logger.log(semesters);
  return ContentService.createTextOutput(
    JSON.stringify({
      authToken:
        "1f85413799cefaf0f9e43d4f6f9bbec6e2c50aee20d21f7e081c95cef85af607",
      semesters
    })
  );
}

function doPost(e: GoogleAppsScript.Events.DoPost) {
  const data: {
    studentToken: string;
    spreadsheetsID: string;
    authToken: string;
    paid: boolean;
  } = JSON.parse(e.postData.contents);

  if (
    data.authToken !==
    "1f85413799cefaf0f9e43d4f6f9bbec6e2c50aee20d21f7e081c95cef85af607"
  ) {
    throw "wrong authToken";
  }

  const sheet = SpreadsheetApp.openById(data.spreadsheetsID).getSheets()[0];

  const row = sheet.createTextFinder(data.studentToken).findNext().getRow();

  if (data.paid) {
    sheet.getRange(row, 8).setValue("v");
  }

  const studentID = sheet.getRange(row, 4).getValue();

  return ContentService.createTextOutput(
    JSON.stringify({ isClubMember: isClubMember(studentID) })
  );
}

function isClubMember(studentID: string) {
  const folders = DriveApp.getFolderById(
    "1btwNYSROh4hd338CPMSks0kRrN4x40sW"
  ).getFolders();
  let result = false;
  while (folders.hasNext()) {
    const files = folders.next().getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() == MimeType.GOOGLE_SHEETS.toString()) {
        if (
          SpreadsheetApp.openById(file.getId())
            .getSheets()[0]
            .createTextFinder(studentID)
            .findAll().length != 0
        ) {
          result = true;
        }
      }
    }
  }
  return result;
}

class course {
  name: string;
  spreadsheetsID: string;
}

class semester {
  name: string;
  activeCourses: course[];
}
