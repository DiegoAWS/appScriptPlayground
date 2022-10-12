//#region Common functions

//#region Enums 
const LOGS = "LOGS";


//#endregion


const cache = CacheService.getDocumentCache();

const writeCache = (key, newValue) => cache.put(key, JSON.stringify(newValue));

const readCache = (key) => JSON.parse(cache.get(key))

const log = (data) => {
  console.log(data)
  const oldLogs = readCache(LOGS) || [];
  const newLogs = [...oldLogs, data]
  writeCache(LOGS, newLogs)
};

const clearCache = () => {
  writeCache(LOGS, [])
}

const readAllCache = () => {
  const logs = readCache(LOGS);

  return { logs }
};

//#endregion

function start() {
  const html = HtmlService.createHtmlOutputFromFile('modal')
    .setWidth(800)
    .setHeight(500);
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Tool');
}


//#region Rename Engines

const renameElement = (element, searchRegex, renameValue) => {
  const name = element.getName();

  if (name.search(searchRegex) !== -1) {
    // Rename  Process 
    const newName = name.replace(searchRegex, renameValue);
    element.setName(newName);

    return {
      hasBeenRenamed: true,
      currentName: newName
    };
  }

  return {
    hasBeenRenamed: false,
    currentName: name
  };
};


const renameBodyDocument = (file, searchExpressionAsString, renameValue, search = []) => {

  const mimeType = file?.getMimeType()

  if (
    mimeType === "application/vnd.google-apps.document"
  ) {


    const idFile = file.getId();
    const fileName = file.getName();
    const body = DocumentApp.openById(idFile)?.getBody();
    const rangeElement = body.findText(searchExpressionAsString);

    if (rangeElement) {
      log(`Renaming inner Text on: ${fileName}`);
      body.replaceText(searchExpressionAsString, renameValue);
      return true;
    }
  }

  if (
    mimeType === "application/vnd.google-apps.spreadsheet"
  ) {
    const idFile = file.getId();

    var sheets = SpreadsheetApp.openById(idFile).getSheets();

    if (sheets.length === 0) return;


    sheets.forEach(sheet => {

      const textFinderSimple = sheet.createTextFinder(searchExpressionAsString);
      textFinderSimple.replaceAllWith(renameValue);

      search.forEach(look => {
        const textFinder = sheet.createTextFinder(look);
        textFinder.replaceAllWith(renameValue);
      })

    })
  }
  return false;
};

//#endregion

//#region Main RECURSIVE function

const treeOperation = (rootSource, action = () => { }, deepth = 0) => {

  action({ item: rootSource, deepth, isFolder: true });

  // Files loop
  const files = rootSource.getFiles();
  while (files.hasNext()) {
    const file = files.next();

    action({ item: file, deepth: deepth + 1, isFolder: false });
  }

  // Folders loop
  const folders = rootSource?.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();

    // recursive call
    treeOperation(folder, action, deepth + 1);
  }
};

//#endregion


//#region TBC Strategy
const tbcStrategy = (rootFolderUrl, renameValue) => {

  try {
    clearCache()

    const folderId = rootFolderUrl.toString().split("/").pop();

    const rootFolder = DriveApp.getFolderById(folderId);

    const action = ({ isFolder, item, deepth }) => {
      const regexTBC = /TBC|TBD/
      const searchExpressionAsString = regexTBC.toString().replace(/^\/|\/$/g, "");

      const renamed = renameElement(item, regexTBC, renameValue);
      const name = renamed?.currentName;

      const hasBeenUpdatedInside = (!isFolder) && renameBodyDocument(item, searchExpressionAsString, renameValue, ["TBC", "TBD"]);

      log(`${"|  ".repeat(deepth)}${isFolder ? '📁' : '|  📄'} ${renamed?.hasBeenRenamed ? "𝐑𝐄𝐍𝐀𝐌𝐄𝐃" : ""} ${hasBeenUpdatedInside ? "𝐔𝐏𝐃𝐀𝐓𝐄𝐃" : ""} ${name}`);

    }

    treeOperation(rootFolder, action)

    return true;

  } catch (e) {
    console.log(e)
    return false
  }
};
//#endregion

//#region Duplicate Folder
const duplicateFolder = (rootFolderUrl = "", newName = "") => {

  if (!rootFolderUrl) return;

  try {
    clearCache()

    const folderId = rootFolderUrl.toString().split("/").pop();

    const rootFolder = DriveApp.getFolderById(folderId);

    const parent = rootFolder.getParents().next()
    const rootName = rootFolder.getName()
    let copyName = newName || `COPY ${rootName}`;

    for (let i = 1; parent.getFoldersByName(copyName).hasNext(); i++) {
      copyName = `COPY (${i}) ${rootName}`;
    }

    if (!newName)
      log(`|📁 ✅ ${copyName}`)

    let targetFolder = parent.createFolder(copyName)
    let path = [targetFolder];



    const action = ({ isFolder, item, deepth }) => {
      if (isFolder) {

        if (item.getId() === folderId) return;

        while (deepth < path.length) {
          // going up in the folders tree
          path.pop();
        }



        const createdFolder = path[deepth - 1]?.createFolder(item?.getName())
        path.push(createdFolder)

      } else {
        const destinationFolder = path[path.length - 1]
        const name = item.getName();
        item.makeCopy(name, destinationFolder)
      }

      if (newName) return; // Avoiding feedback on custom function use

      log(`|--${"--".repeat(deepth)}${isFolder ? '📁' : '|  📄'} ✅ ${item}`)

    }

    treeOperation(rootFolder, action)

    return true;

  }
  catch (e) {
    console.log(e)
    return false
  }
};
//#endregion

//#region Empty Space
const emptySpace = (rootFolderUrl, codeName, clientName, users) => {


  try {
    clearCache()

    const folderId = rootFolderUrl.toString().split("/").pop();

    const rootFolder = DriveApp.getFolderById(folderId);

    const persons = [...new Set(users.filter(item => !!item))]

    const personsKey = /\[PersonName\]/

    const rename = [
      [/\[CustomerName\]|\[CLIENT\]/, clientName],
      [/\[Code\s*Name\]/, codeName],
    ]

    const action = ({ item, isFolder, deepth }) => {

      log(`|--${"--".repeat(deepth)}${isFolder ? '📁' : '|  📄'} ✅ ${item}`)

      rename.forEach(([search, replace]) => {
        renameElement(item, search, replace)
        if (!isFolder) {
          renameBodyDocument(item, search.toString().replace(/^\/|\/$/g, ""), replace)
        }
      })
    }

    const personsAction = ({ item, deepth, isFolder }) => {

      log(`|--${"--".repeat(deepth)}${isFolder ? '📁' : '|  📄'} ✅ ${item}`)

      const oldName = item.getName()
      if (oldName.search(personsKey) === -1) return;

      persons.forEach(person => {
        const newName = oldName.replace(personsKey, person)
        if (!isFolder) {
          renameBodyDocument(item, "\[Your Name\]", person)
          item.makeCopy(newName)
        } else {
          duplicateFolder(item, newName)
        }
      })
      item.setTrashed(true);
    }

    treeOperation(rootFolder, action) // expressions renames
    treeOperation(rootFolder, personsAction) // Person names renames

    return true
  }
  catch (e) {
    console.log(e)
    return false
  }

};
//#endregion

//#region Replace Name
const replaceName = (rootFolderUrl, searchName, renameValue) => {

  try {
    clearCache()

    const folderId = rootFolderUrl.toString().split("/").pop();

    const rootFolder = DriveApp.getFolderById(folderId);

    const action = ({ isFolder, item, deepth }) => {
      const renamed = renameElement(item, searchName, renameValue);
      const name = renamed?.currentName;

      const hasBeenUpdatedInside = (!isFolder) && renameBodyDocument(item, searchName, renameValue, [searchName]);

      log(`${"|  ".repeat(deepth)}${isFolder ? '📁' : '|  📄'} ${renamed?.hasBeenRenamed ? "𝐑𝐄𝐍𝐀𝐌𝐄𝐃" : ""} ${hasBeenUpdatedInside ? "𝐔𝐏𝐃𝐀𝐓𝐄𝐃" : ""} ${name}`);

    }

    treeOperation(rootFolder, action)

    return true;

  } catch (e) {
    console.log(e)
    return false
  }
};
//#endregion

