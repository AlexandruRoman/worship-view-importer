var JSZip = require("jszip");
var convert = require("xml-js");
const fs = require("fs");
const path = require("path");
const fse = require("fs-extra");

const getAllFiles = function (dirPath, arrayOfFiles) {
  const files = fs.readdirSync(dirPath);

  arrayOfFiles = arrayOfFiles || [];

  files.forEach(function (file) {
    if (fs.statSync(dirPath + "/" + file).isDirectory()) {
      arrayOfFiles = getAllFiles(dirPath + "/" + file, arrayOfFiles);
    } else {
      arrayOfFiles.push(path.join(dirPath, "/", file));
    }
  });

  return arrayOfFiles;
};

const getFileToStrophes = async (fileName) => {
  const querySelectorAll = (obj, tag) => {
    if (obj.name === tag) return [obj];
    if (obj.elements) {
      let r = [];
      obj.elements.forEach((element) => {
        const result = querySelectorAll(element, tag);
        r.push(...result);
      });
      return r;
    }
    return [];
  };

  const extractFromTextBody = (textBodyObj) => {
    let text = "";
    const pObjs = textBodyObj?.elements?.filter((e) => e.name === "a:p");
    pObjs &&
      pObjs.forEach((pObj) => {
        const tokenObjs = pObj.elements;
        if (tokenObjs && tokenObjs.length > 0) {
          tokenObjs.forEach((tokenObj) => {
            if (tokenObj.name === "a:r") {
              const [textObj] = querySelectorAll(tokenObj, "a:t");
              if (textObj) {
                if (textObj.elements && textObj.elements[0].text)
                  text += textObj.elements[0].text;
                else text += " ";
              }
            } else if (tokenObj.name === "a:br") text += "\n";
          });
        }
        text += "\n";
      });

    return text;
  };

  const getSanitizedSlides = (text) => {
    const getRepeatedSlide = (t) => {
      let textToBeProcessed = t;
      let factor = 1;
      const result =
        /^(\/\/|\/:|\|:|%)(?<text>[^//]+)(\/\/|:\/|:\||%)\s*(?<factor>\d\s*[xX]|[xX]\s*\d)?$/.exec(
          t
        );
      if (result && result.groups && result.groups.text) {
        textToBeProcessed = result.groups.text;
        if (result.groups.factor) {
          const factorResult = result.groups.factor.match(
            /(?<factor>\d)\s*[xX]|[xX]\s*(?<factor2>\d)/
          );
          if (factorResult && factorResult.groups) {
            const x =
              parseInt(factorResult.groups.factor) ||
              parseInt(factorResult.groups.factor2);
            if (x) factor = x;
          }
        } else factor = 2;
      }

      return { textToBeProcessed, factor };
    };

    const getExpandedText = (t) => {
      let expandedText = "";
      let chunks = [
        ...t.matchAll(
          /((\/\/|\/:)([^/]+)(\/\/|:\/)(\s*\d\s*[xX]|\s*[xX]\s*\d)?)|[^/]+/g
        ),
      ];
      chunks.forEach((chunk) => {
        let { factor, textToBeProcessed } = getRepeatedSlide(chunk[0]);
        expandedText += Array(factor).fill(textToBeProcessed).join("\n");
      });
      return expandedText;
    };

    let newText = text
      .replaceAll(/(!|\?|,|\.|\s|\n)+$/g, "")
      .replaceAll(/^(!|\?|,|\.|\s|\n)+/g, "");
    let { factor, textToBeProcessed } = getRepeatedSlide(newText);

    textToBeProcessed = getExpandedText(textToBeProcessed)
      .replaceAll(/(!|\?|,|;|:|\.|\s|\n)+\n/g, "\n")
      .replaceAll(/\n\s+/g, "\n")
      .replaceAll(/(!|\?|,|;|:|\.|\s|\n)+$/g, "")
      .replaceAll(/^(!|\?|,|;|:|\.|\s|\n)+/g, "")
      .replaceAll(/\s+\s/g, " ");
    if (textToBeProcessed === "") return [];
    return Array(factor).fill(textToBeProcessed);
  };

  const data = fs.readFileSync(fileName);
  const zip = await JSZip.loadAsync(data);
  let strophes = [];
  const files = Object.entries(zip.folder("ppt/slides/").files)
    .filter(([path]) => path.startsWith("ppt/slides/slide"))
    .sort((a, b) => a[0].localeCompare(b[0]));
  for (const [, file] of files) {
    const fileContent = await file.async("text");
    const json = convert.xml2js(fileContent, { compact: false });
    const tags = querySelectorAll(json, "p:txBody");
    const tagsText = tags.map(extractFromTextBody);
    const [slideText] = tagsText
      .filter((text) => !text.includes("bisericaemanuelploiesti"))
      .filter((text) => !text.includes("Aspose.Slides"))
      .sort((a, b) => b.length - a.length);
    strophes.push(...getSanitizedSlides(slideText));
  }
  return strophes;
};

const writeFile = (fileName, data) => {
  fse.outputFileSync(
    fileName.replace("fancy", "fancy2").replace(".pptx", ""),
    data
  );
  console.log("Saved!", fileName.split("\\").at(-1));
};

const getSong = (strophes) => {
  const getMappedStrophes = (strophes) => {
    let uniqueStrophes = {};
    strophes.forEach((cur) => {
      if (!uniqueStrophes[cur]) uniqueStrophes[cur] = 1;
      else uniqueStrophes[cur]++;
    });
    const sortedStrophes = Object.entries(uniqueStrophes).sort(
      (a, b) => b[1] - a[1]
    );
    let mappedStrophes = {};
    let count = 0;
    sortedStrophes.forEach(([text, occurence]) => {
      if (occurence > 1 && !mappedStrophes["r"]) mappedStrophes["r"] = text;
      else {
        count++;
        mappedStrophes["v" + count] = text;
      }
    });
    return Object.entries(mappedStrophes);
  };

  const getArrangement = (strophes, mappedStrophes) => {
    let arrangement = [];
    strophes.forEach((s) => {
      const [key] = mappedStrophes.find(([, text]) => s === text);
      arrangement.push(key);
    });
    return arrangement;
  };

  const mappedStrophes = getMappedStrophes(strophes);
  const arrangement = getArrangement(strophes, mappedStrophes);
  let songText = "";
  mappedStrophes.forEach(([key, text]) => {
    songText += `${key}\n`;
    const lines = text.split("\n");
    lines.forEach((line, index) => {
      if (index % 2 === 0 || lines.length - 1 === index)
        songText += `${line}\n`;
      else songText += `${line}\n\n`;
    });
    songText += "---\n";
  });
  songText += arrangement.join(" ");
  return songText;
};

const files = getAllFiles("/Projects/church/fancy");
files.forEach(async (fileName) => {
  if (fileName.endsWith(".pptx") && !fileName.includes("~")) {
    try {
      const strophes = await getFileToStrophes(fileName);
      const data = getSong(strophes);
      writeFile(fileName, data);
    } catch (error) {
      console.log("-------------- ERROR -------------");
    }
  }
});
