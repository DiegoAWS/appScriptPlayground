//#region Common functions

const getInfo = (entry) => {
  const id = entry.toString().split("/").pop();

  let element = DriveApp.getFolderById(id);

  if (!element.getMimeType) {
    element = DriveApp.getFileById(id)
  }

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

const deleteCache = (key) => {
  cache.remove(key)
}

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

  return sortedResult;
};

//#endregion


//#region Rename Body

const renameBodyDocument = (file, replace = []) => {
  const clearRegex = (reg) => {
    const regString = reg.toString()
    return regString.substring(regString.indexOf("/") + 1, regString.lastIndexOf("/"))
  }

  const mimeType = file?.getMimeType()
  const idFile = file.getId();
  let changed = false;
  if (
    mimeType === "application/vnd.google-apps.document"
  ) {
    const body = DocumentApp.openById(idFile)?.getBody();

    replace.forEach(([searchRegex, replace]) => {
      const search = clearRegex(searchRegex);
      if (body.findText(search)) {
        body.replaceText(search, replace);
        changed = true;
      }
    })
  }

  if (
    mimeType === "application/vnd.google-apps.spreadsheet"
  ) {
    const sheets = SpreadsheetApp.openById(idFile).getSheets();

    sheets.forEach(sheet => {
      replace.forEach(([searchRegex, replace]) => {
        const search = clearRegex(searchRegex);
        const textFinder = sheet.createTextFinder(search);
        if (textFinder) {
          textFinder.replaceAllWith(replace);
          changed = true;
        }
      })

    })
  }

  if (
    mimeType === "application/vnd.google-apps.presentation"
  ) {
    const slides = SlidesApp.openById(idFile)?.getSlides();

    slides.forEach(slide => {
      replace.forEach(([searchRegex, replace]) => {
        const numberReplaced = slide.replaceAllText(searchRegex, replace)
        console.log({ numberReplaced })
        if (numberReplaced) {
          changed = true;
        }

      })
    })
  }

  return changed;
};
//#endregion


//#region Process Item 

const processItem = (element, index, cacheKey, entry, isCopy, users) => {
  const targetPaths = () => readCache(cacheKey)
  const setTargetPath = (data) => writeCache(cacheKey, data)
  let processElement = () => { }
  let getNewName = (e) => e?.name;
  const { id, isFolder, depth } = element;

  let classElement = getInfo(id).element

  if (depth === 0) {
    const { parent } = getInfo(id)
    setTargetPath([parent.getId()])
  }

  if (Array.isArray(entry)) {
    getNewName = ({ name }) => entry.reduce((acc, [pattern, replace]) => acc.replace(pattern, replace), name)

    processElement = (e) => {
      if (!e.getMimeType) return;

      return renameBodyDocument(e, entry)

    }
  }

  while (depth < targetPaths().length - 1) {
    // going up in the folders tree
    setTargetPath(targetPaths().slice(0, -1));
  }

  const parentInTarget = DriveApp.getFolderById(targetPaths()[depth])

  const updatedName = ((isCopy && depth === 0) ? getNowString() : "") + getNewName(element);

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

  const logLine = (newCreated, { depth, isFolder }) => Logger.log(`${" ".repeat(depth)}${isFolder ? "📁" : "📄"} ${newCreated.getName()}`)


  listOfFiles.forEach((element, item) => {

    processItem(element, item, cacheKey, getNewName, logLine)
  })

  deleteCache(cacheKey)

}
//#endregion


//#region Main entrypoints



// Open Modal
function start() {
  const html = HtmlService.createHtmlOutputFromFile('modal')
    .setWidth(800)
    .setHeight(500);
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Tool');
}

const index = (func, ...params) => {
  try {
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
  } catch (e) {
    console.log("ERROR")
    console.log(e)
  }
}

//#endregion 


const test = () => {
  const url = "https://drive.google.com/drive/folders/15wnmtVFo4ZdSb7cXMnmLeMepmBs8elL3" // big production like folder
  // const url = "https://drive.google.com/drive/folders/1f3nM4pLXwYOJutbeL0QJOa1Duc25eMyu" // test folder little 26 elements
  const list = getListFiles(url);

  const clientName = "Client 113";
  const codeName = "Pink";

  const entry = [
    [/[\(\[](?:CustomerName|CLIENT)[\)\]]/, clientName],
    [/[\(\[\{<](?:Code\s?Name|CODENAME)[\)\]\}>]/i, codeName],
    [/Project <Name>/, "Project " + codeName]
  ];

  const users = [
    "Diego Escobar",
    "Hazem Alborous"
  ];

  const total = list.length
  list.forEach((element, index) => {
    const { name, isFolder, depth } = processItem(element, index, "test", entry, true, users)
    console.log(`${index + 1}/${total} ${" ".repeat(depth)}${isFolder ? "📁" : "📄"} ${name}`)
  });
}

const test2 = () => {
  const id = "16KUzNx03ayKMmjhyvFDP9XGSvBypTb7DX20QfoW48OU";

  const file = DriveApp.getFileById(id)

  const mimeType = file?.getMimeType()
  const idFile = file.getId();
  const replace = [[/Project <Name>/, "Project " + "Pink"]]
  if (
    mimeType === "application/vnd.google-apps.presentation"
  ) {
    const slides = SlidesApp.openById(idFile)?.getSlides();
    console.log(slides.length)
    slides.forEach(slide => {
      replace.forEach(([searchRegex, replace]) => {
        const a = slide?.replaceAllText && slide.replaceAllText(searchRegex, replace)

        console.log({ a })
      })
    })
  }



}






