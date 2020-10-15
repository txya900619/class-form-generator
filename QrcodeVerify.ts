const folderID = "0B0IVnu5YlOTNZzZsV21LLUd2VHc";

function doPost(e: GoogleAppsScript.Events.DoPost) {
  const data: {
    password: string;
  } = JSON.parse(e.postData.contents);
  const password: string = data.password;
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
      const sheets = SpreadsheetApp.openById(spreadsheetsID).getSheets()[0];
      if (sheets.getRange(2, 3).getValue() != "") {
        continue;
      }
      activeCourses.push({
        name: subSubFolderName,
        spreadsheetsID: sheets.getRange(2, 4).getValue()
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
