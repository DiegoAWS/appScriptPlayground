//#region Common functions

const getInfo = (entry) => {
  const id = entry.toString().split("/").pop();

  const element = DriveApp.getFolderById(id);

  const name = element.getName();
  const parent = element.getParents().next();

  return { id, element, name, parent };
}

const getNowString = () => {
  const now = new Date();
  return `${now.getFullYear()
    }${(now.getMonth() + 1).toString().padStart(2, 0)
    }${now.getDate().toString().padStart(2, 0)
    }_${now.getHours().toString().padStart(2, 0)
    }${now.getMinutes().toString().padStart(2, 0)
    }${now.getSeconds().toString().padStart(2, 0)} - `
}


//#region Enums 
const LOGS = "LOGS";
const TPATH = "TPATH"

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

const deleteCache = (key) => {
  cache.remove(key)
}

const readAllCache = () => {
  const logs = readCache(LOGS);

  return { logs }
};

//#endregion

function start() {
  clearCache()
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




//#endregion

//#region Main RECURSIVE function

const treeOperation = (rootSource, action = () => { }, depth = 0) => {

  action({ item: rootSource, depth, isFolder: true });

  // Files loop
  const files = rootSource.getFiles();
  while (files.hasNext()) {
    const file = files.next();

    action({ item: file, depth: depth + 1, isFolder: false });
  }

  // Folders loop
  const folders = rootSource?.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();

    // recursive call
    treeOperation(folder, action, depth + 1);
  }
};

//#endregion


//#region TBC Strategy
const tbcStrategy = (rootFolderUrl, renameValue) => {

  try {
    clearCache()

    const folderId = rootFolderUrl.toString().split("/").pop();

    const rootFolder = DriveApp.getFolderById(folderId);

    const action = ({ isFolder, item, depth }) => {
      const regexTBC = /TBC|TBD/
      const searchExpressionAsString = regexTBC.toString().replace(/^\/|\/$/g, "");

      const renamed = renameElement(item, regexTBC, renameValue);
      const name = renamed?.currentName;

      const hasBeenUpdatedInside = (!isFolder) && renameBodyDocument(item, searchExpressionAsString, renameValue, ["TBC", "TBD"]);

      log(`${"|  ".repeat(depth)}${isFolder ? '📁' : '|  📄'} ${renamed?.hasBeenRenamed ? "𝐑𝐄𝐍𝐀𝐌𝐄𝐃" : ""} ${hasBeenUpdatedInside ? "𝐔𝐏𝐃𝐀𝐓𝐄𝐃" : ""} ${name}`);

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
const duplicateFolder = (rootFolderUrl, newName = "", verbose = false) => {

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

    if (verbose || !newName) {
      log("")
      log("Starting COPY process")
      log("_".repeat(20))
      log(`|📁 ✅ ${copyName}`)
    }

    let targetFolder = parent.createFolder(copyName)
    let path = [targetFolder];



    const action = ({ isFolder, item, depth }) => {
      if (isFolder) {

        if (item.getId() === folderId) return;

        while (depth < path.length) {
          // going up in the folders tree
          path.pop();
        }



        const createdFolder = path[depth - 1]?.createFolder(item?.getName())
        path.push(createdFolder)

      } else {
        const destinationFolder = path[path.length - 1]
        const name = item.getName();
        item.makeCopy(name, destinationFolder)
      }


      if (verbose || !newName)
        log(`  ${"|  ".repeat(depth)}${isFolder ? '📁' : '|  📄'} ✅ ${item}`)

    }

    treeOperation(rootFolder, action)

    return targetFolder?.getId();

  }
  catch (e) {
    console.log(e)
    return false
  }
};
//#endregion

//#region Empty Space
const emptySpace = (rootFolderUrlEntry, codeName, clientName, users) => {
  try {
    clearCache()

    const rootFolderName = DriveApp.getFolderById(rootFolderUrlEntry.toString().split("/").pop()).getName();

    const newName = `${codeName}-${clientName}-${rootFolderName}`

    const rootFolderUrl = duplicateFolder(rootFolderUrlEntry, newName, true)

    if (!rootFolderUrl) return;


    log("")
    log("Starting RENAME process")
    log("_".repeat(20))


    const folderId = rootFolderUrl.toString().split("/").pop();

    const rootFolder = DriveApp.getFolderById(folderId);

    const persons = [...new Set(users.filter(item => !!item))]

    const personsKey = /\[PersonName\]/

    const rename = [
      [/\[CustomerName\]|\[CLIENT\]/, clientName],
      [/\[Code\s*Name\]/, codeName],
      [/\<Code\s*Name\>/, codeName],
      [/{Codename}/, codeName]
    ]

    const action = ({ item, isFolder, depth }) => {

      log(`  ${"|  ".repeat(depth)}${isFolder ? '📁' : '|  📄'} ✅ ${item}`)

      rename.forEach(([search, replace]) => {
        renameElement(item, search, replace)
        if (!isFolder) {
          renameBodyDocument(item, search.toString().replace(/^\/|\/$/g, ""), replace)
        }
      })
    }

    const personsAction = ({ item, depth, isFolder }) => {

      log(`  ${"|  ".repeat(depth)}${isFolder ? '📁' : '|  📄'} ✅ ${item}`)

      const oldName = item.getName()
      if (oldName.search(personsKey) === -1) return;

      persons.forEach(person => {
        const newName = oldName.replace(personsKey, person)
        if (!isFolder) {
          renameBodyDocument(item, "\[Your Name\]", person)
          renameBodyDocument(item, "<Your Name>", person)
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

    const action = ({ isFolder, item, depth }) => {
      const renamed = renameElement(item, searchName, renameValue);
      const name = renamed?.currentName;

      const hasBeenUpdatedInside = (!isFolder) && renameBodyDocument(item, searchName, renameValue, [searchName]);

      log(`${"|  ".repeat(depth)}${isFolder ? '📁' : '|  📄'} ${renamed?.hasBeenRenamed ? "𝐑𝐄𝐍𝐀𝐌𝐄𝐃" : ""} ${hasBeenUpdatedInside ? "𝐔𝐏𝐃𝐀𝐓𝐄𝐃" : ""} ${name}`);

    }

    treeOperation(rootFolder, action)

    return true;

  } catch (e) {
    console.log(e)
    return false
  }
};
//#endregion

















//#region make list of files

const getListFiles = (rootSource) => {

  if (!rootSource?.getId) {
    rootSource = DriveApp.getFolderById(rootSource.toString().split("/").pop())
  }

  const result = []

  const parentOfSource = rootSource.getParents().next()

  const recusiveCall = (item, parent, path = [], depth = 0) => {

    result.push({
      id: item.getId(),
      name: item.getName(),
      parent: parent.getId(),
      mimeType: "folder",
      path,
      isFolder: true,
      depth
    });

    path = path.concat(item.getName());
    depth++;

    const files = item?.getFiles();
    while (files.hasNext()) {
      const file = files.next();

      result.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        path,
        parent: parent.getId(),
        isFolder: false,
        depth,
      });
    }

    const folders = item?.getFolders();
    while (folders.hasNext()) {
      const folder = folders.next();
      recusiveCall(folder, item, path.concat(folder.getName()), depth)
    }
  }

  recusiveCall(rootSource, parentOfSource)

  const sortedResult = result
    .sort(({ path: aPath }, { path: bPath }) => aPath.join("/")?.localeCompare(bPath.join("/")))
    .map(({ path, ...rest }) => rest);

  return sortedResult;
};

//#endregion


//   const newName = 


//#region Rename Body
const renameBodyDocument = (file, replace = []) => {

  const mimeType = file?.getMimeType()
  const idFile = file.getId();
  if (
    mimeType === "application/vnd.google-apps.document"
  ) {

    const body = DocumentApp.openById(idFile)?.getBody();

    replace.forEach(([search, replace]) => {
      if (body.findText(search)) {
        body.replaceText(search, replace);
      }
    })
  }

  if (
    mimeType === "application/vnd.google-apps.spreadsheet"
  ) {
    const sheets = SpreadsheetApp.openById(idFile).getSheets();

    sheets.forEach(sheet => {
      search.forEach(look => {
        const textFinder = sheet.createTextFinder(look);
        textFinder.replaceAllWith(renameValue);
      })

    })
  }
  return false;
};
//#endregion


//#region Process Item 
const processItem = (element, index, cacheKey, entry, isCopy = true, users) => {

  const targetPaths = () => readCache(cacheKey)
  const setTargetPath = (data) => writeCache(cacheKey, data)
  let processElement = () => { }
  let getNewName = () => e?.name;

  const { id, isFolder, depth } = element;

  const classElement = getInfo(id).element


  if (depth === 0) {
    const { parent } = getInfo(id)
    setTargetPath([parent.getId()])
  }

  if (typeof entry === "string") {
    getNewName = ({ name, depth }) => depth === 0 ? entry : name
  }

  if (Array.isArray(entry)) {
    getNewName = ({ name }) => entry.reduce((acc, [pattern, replace]) => acc.replace(pattern, replace), name)

    processElement = (classElement) => {
      if (!classElement.getMimeType) return;

      renameBodyDocument(classElement, entry)

    }
  }


  const updatedName = getNewName(element);

  while (depth < targetPaths().length - 1) {
    // going up in the folders tree
    setTargetPath(targetPaths().slice(0, -1));
  }

  const parentInTarget = DriveApp.getFolderById(targetPaths()[depth])


  if (isFolder) {

    if (isCopy) {
      classElement = parentInTarget.createFolder(updatedName)
      processElement(classElement, element, index);

    } else {
      classElement.setName(updatedName)
      processElement(classElement, element, index)
    }
    setTargetPath([...targetPaths(), classElement.getId()]);

    return {
      ...element,
      name: updatedName,
      id: classElement.getId(),
    };

  }

  const file = DriveApp.getFileById(id)

  if (isCopy) {
    classElement = file.makeCopy(updatedName, parentInTarget)
    processElement(classElement, element, index);
  } else {
    classElement.setName(updatedName)
    processElement(classElement, element, index);

  }
  return {
    ...element,
    name: updatedName,
    id: classElement.getId(),
  };

}
//#endregion





//#region copy ALL folders under folderId
const copyFolder = (folderId, newName) => {

  const cacheKey = "COPY";
  const listOfFiles = getListFiles(folderId);

  const getNewName = ({ name, depth }) => depth === 0 ? (newName || getNowString() + name) : name

  const logLine = (newCreated, { depth, isFolder }) => console.log(`${" ".repeat(depth)}${isFolder ? "📁" : "📄"} ${newCreated.getName()}`)


  listOfFiles.forEach((element, item) => {

    processItem(element, item, cacheKey, getNewName, logLine)
  })

  deleteCache(cacheKey)

}
//#endregion



const index = (func, ...params) => {
  switch (func) {
    case "getListFiles":
      return getListFiles(...params)
    case "writeCache":
      return writeCache(...params)
    case "readCache":
      return readCache(...params)
    case "processItem":
      return processItem(...params)
    case "deleteCache":
      return deleteCache(...params)
    case "getNowString":
      return getNowString(...params)
    case "getInfo":
      return getInfo(...params)

  }
}





