import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";
import xlsxCalc from "https://esm.sh/xlsx-calc";
import * as formulajs from "https://esm.sh/@formulajs/formulajs@4.5.6";

const DEFAULT_SPREADSHEET_URL = "./assets/Spreadsheet.Default.xlsx";
const DEFAULT_FILENAME = "Copy of Stardew Step Manipulation V1.6.xlsx";

const saveInput = document.getElementById("save-input");
const dropzone = document.getElementById("dropzone");
const fileCard = document.getElementById("file-card");
const fileName = document.getElementById("file-name");
const clearButton = document.getElementById("clear-button");
const analyzeButton = document.getElementById("analyze-button");
const downloadButton = document.getElementById("download-button");
const resultText = document.getElementById("result-text");
const statusPill = document.getElementById("status-pill");
const inputGrid = document.getElementById("input-grid");
const warningsList = document.getElementById("warnings-list");
const resultsBody = document.getElementById("results-body");
const npcBody = document.getElementById("npc-body");
const summaryText = document.getElementById("summary-text");

let selectedFile = null;
let parsedInputs = null;
let importedFunctions = false;

const MACHINE_MAP = new Map([
  ["Furnace", "Furnaces"],
  ["Crystalarium", "Crystalariums"],
  ["Preserves Jar", "Jars"],
  ["Keg", "Kegs"],
  ["Tapper", "Tappers"],
  ["Heavy Tapper", "Tappers"],
  ["Bee House", "Bee Houses"],
  ["Oil Maker", "Oil Makers"],
]);

saveInput.addEventListener("change", () => {
  updateSelectedFile(saveInput.files?.[0] ?? null);
});

clearButton.addEventListener("click", () => {
  saveInput.value = "";
  updateSelectedFile(null);
  clearOutputs();
});

analyzeButton.addEventListener("click", async () => {
  if (!selectedFile) {
    setStatus("No file", "Choose a Stardew save file first.");
    return;
  }

  try {
    setStatus("Parsing", "Reading save file...");
    const xmlText = await selectedFile.text();
    parsedInputs = parseSaveFile(xmlText, selectedFile.name);
    renderInputs(parsedInputs);

    setStatus("Calculating", "Loading workbook and recalculating predictions...");
    const result = await calculatePredictions(parsedInputs);
    renderResults(result.rows);

    const first = result.rows[0];
    summaryText.textContent = first
      ? `Parsed ${parsedInputs.stepsTaken} steps. First row: +${first.additionalSteps} steps, ${first.weather}, ${first.dish}.`
      : "No prediction rows were produced.";
    setStatus("Done", `Parsed ${parsedInputs.stepsTaken} steps and generated ${result.rows.length} result rows.`);
  } catch (error) {
    console.error(error);
    setStatus("Error", error?.message || "Prediction failed.");
    summaryText.textContent = "Calculation failed. Check the warning text for details.";
  }
});

downloadButton.addEventListener("click", async () => {
  await downloadDefaultSpreadsheet(DEFAULT_FILENAME);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-dragging");
  });
});

dropzone.addEventListener("drop", (event) => {
  const file = event.dataTransfer?.files?.[0] ?? null;
  if (!file) {
    return;
  }

  updateSelectedFile(file);
});

function updateSelectedFile(file) {
  selectedFile = file;
  parsedInputs = null;

  if (!file) {
    fileCard.hidden = true;
    fileName.textContent = "";
    setStatus("Ready", "Choose a save file to begin.");
    return;
  }

  fileCard.hidden = false;
  fileName.textContent = file.name;
  setStatus("File loaded", "Save file selected. Click Analyze Save.");
}

function clearOutputs() {
  inputGrid.innerHTML = "";
  warningsList.innerHTML = "<li>No save loaded yet.</li>";
  resultsBody.innerHTML = "<tr><td colspan=\"8\" class=\"empty-cell\">No results yet.</td></tr>";
  npcBody.innerHTML = "<tr><td colspan=\"2\" class=\"empty-cell\">No save loaded yet.</td></tr>";
  summaryText.textContent = "Upload a save file to generate predictions.";
}

function setStatus(label, message) {
  statusPill.textContent = label;
  resultText.textContent = message;
}

function parseSaveFile(xmlText, fileLabel) {
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  if (doc.querySelector("parsererror")) {
    throw new Error("The selected file is not valid XML.");
  }

  const root = doc.documentElement;
  const player = findFirst(root, "player") || root;

  const inputs = {
    saveFilePath: fileLabel,
    season: readTextAny(root, "currentSeason", "spring"),
    currentDay: readIntAny(root, "dayOfMonth"),
    totalDaysPlayed: readIntAny(player, "daysPlayed"),
    stepsTaken: readIntAny(player, "stepsTaken"),
    gameSeed: readIntAny(root, "uniqueIDForThisGame"),
    debrisTomorrowFlag: detectDebrisTomorrow(root),
    npcs: [],
    machines: {
      Furnaces: 0,
      Crystalariums: 0,
      Jars: 0,
      Kegs: 0,
      Tappers: 0,
      "Bee Houses": 0,
      "Oil Makers": 0,
    },
    specialRequests: {
      biomeBalance: false,
      cropOrder: false,
      robinsResourceRush: false,
      islandIngredients: false,
    },
    isSeasonTransitionNight: false,
    relevantWeedCount: 0,
    warnings: [],
  };

  inputs.isSeasonTransitionNight = inputs.currentDay === 28;

  readNpcFriendships(root, inputs);
  readMachines(root, inputs);
  readSpecialRequests(root, inputs);
  inputs.relevantWeedCount = countRelevantWeeds(root);

  if (!inputs.stepsTaken) {
    inputs.warnings.push("`stepsTaken` was missing or zero in the save, so the result may need manual correction.");
  }
  if (inputs.isSeasonTransitionNight && inputs.relevantWeedCount === 0) {
    inputs.warnings.push("This looks like a season transition night, but no relevant weeds were detected. Verify that count manually if needed.");
  }
  inputs.warnings.push("Machine counts are inferred from current machine timers in the save. If the save timing is wrong for your run, adjust expectations accordingly.");

  return inputs;
}

function readNpcFriendships(root, inputs) {
  const friendshipRoot = findFirst(root, "friendshipData");
  if (!friendshipRoot) {
    inputs.warnings.push("Friendship data was not found.");
    return;
  }

  const items = childrenByLocalName(friendshipRoot, "item");
  for (const item of items) {
    const key = firstChildByLocalName(item, "key");
    const value = firstChildByLocalName(item, "value");
    const name = key?.textContent?.trim();
    if (!name) {
      continue;
    }

    const points = readInt(value, "Points");
    inputs.npcs.push({
      name,
      hearts: Math.max(0, Math.min(14, Math.floor(points / 250))),
    });
  }
}

function readMachines(root, inputs) {
  const currentTime = readIntAny(root, "timeOfDay");
  const minutesRemainingToday = minutesUntilNextSixAm(currentTime);
  const objects = [...root.getElementsByTagName("*")].filter((node) => node.localName === "Object");

  for (const item of objects) {
    const name = readText(item, "name");
    if (!MACHINE_MAP.has(name)) {
      continue;
    }

    if (!readBool(item, "bigCraftable")) {
      continue;
    }

    if (hasAncestorNamed(item, ["items", "inventory", "chest", "fridge"])) {
      continue;
    }

    const minutesUntilReady = readInt(item, "minutesUntilReady");
    const readyForHarvest = readBool(item, "readyForHarvest");
    const heldObject = firstChildByLocalName(item, "heldObject");
    const hasHeldObject = heldObject && [...heldObject.children].length > 0;

    if (readyForHarvest || minutesUntilReady <= 0 || !hasHeldObject) {
      continue;
    }

    if (minutesUntilReady > minutesRemainingToday) {
      inputs.machines[MACHINE_MAP.get(name)] += 1;
    }
  }
}

function readSpecialRequests(root, inputs) {
  const strings = [...root.getElementsByTagName("*")]
    .filter((node) => node.localName === "string")
    .map((node) => (node.textContent || "").trim())
    .filter(Boolean);

  const hasRequest = (name) => strings.some((value) => value.toLowerCase().includes(name.toLowerCase()));

  inputs.specialRequests.biomeBalance = hasRequest("Biome Balance");
  inputs.specialRequests.cropOrder = hasRequest("Crop Order");
  inputs.specialRequests.robinsResourceRush = hasRequest("Robin's Resource Rush");
  inputs.specialRequests.islandIngredients = hasRequest("Island Ingredients");

  if (!Object.values(inputs.specialRequests).some(Boolean)) {
    inputs.warnings.push("No matching Sunday-to-Monday special requests were detected.");
  }
}

function countRelevantWeeds(root) {
  const objects = [...root.getElementsByTagName("*")].filter((node) => node.localName === "Object");
  let count = 0;
  for (const item of objects) {
    const name = readText(item, "name");
    const category = readInt(item, "category");
    if (name === "Weeds" || name === "Green Rain Weeds" || category === -999) {
      count += 1;
    }
  }
  return count;
}

function detectDebrisTomorrow(root) {
  for (const candidate of ["isDebrisWeather", "debrisWeatherForTomorrow"]) {
    const value = readTextAny(root, candidate, "");
    if (value === "true") {
      return "Y";
    }
    if (value === "false") {
      return "N";
    }
  }

  const weather = readTextAny(root, "weatherForTomorrow", "");
  if (!weather) {
    return "N";
  }
  if (weather.toLowerCase().includes("debris") || weather.toLowerCase().includes("wind")) {
    return "Y";
  }
  return Number(weather) === 2 ? "Y" : "N";
}

async function calculatePredictions(inputs) {
  if (!importedFunctions) {
    xlsxCalc.import_functions(formulajs);
    importedFunctions = true;
  }

  const workbookBytes = await fetch(DEFAULT_SPREADSHEET_URL, { cache: "no-store" }).then((response) => {
    if (!response.ok) {
      throw new Error("Could not load the bundled spreadsheet template.");
    }
    return response.arrayBuffer();
  });

  const workbook = XLSX.read(workbookBytes, {
    type: "array",
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
    sheetStubs: true,
  });

  const sheet = workbook.Sheets["Stardew Step Manipulation"];
  if (!sheet) {
    throw new Error("The main worksheet was not found in the bundled template.");
  }

  writeCell(sheet, "C22", inputs.season.toLowerCase(), "s");
  writeCell(sheet, "C23", inputs.currentDay, "n");
  writeCell(sheet, "C24", inputs.totalDaysPlayed, "n");
  writeCell(sheet, "C25", inputs.stepsTaken, "n");
  writeCell(sheet, "C26", inputs.gameSeed, "n");
  writeCell(sheet, "C27", inputs.debrisTomorrowFlag.toUpperCase(), "s");
  writeCell(sheet, "C28", inputs.npcs.length, "n");

  writeCell(sheet, "C32", inputs.machines.Furnaces, "n");
  writeCell(sheet, "C33", inputs.machines.Crystalariums, "n");
  writeCell(sheet, "C34", inputs.machines.Jars, "n");
  writeCell(sheet, "C35", inputs.machines.Kegs, "n");
  writeCell(sheet, "C36", inputs.machines.Tappers, "n");
  writeCell(sheet, "C37", inputs.machines["Bee Houses"], "n");
  writeCell(sheet, "C38", inputs.machines["Oil Makers"], "n");

  writeCell(sheet, "C42", inputs.specialRequests.biomeBalance ? 1 : 0, "b");
  writeCell(sheet, "C43", inputs.specialRequests.cropOrder ? 1 : 0, "b");
  writeCell(sheet, "C44", inputs.specialRequests.robinsResourceRush ? 1 : 0, "b");
  writeCell(sheet, "C45", inputs.specialRequests.islandIngredients ? 1 : 0, "b");

  writeCell(sheet, "C49", inputs.isSeasonTransitionNight ? 1 : 0, "b");
  writeCell(sheet, "C50", inputs.relevantWeedCount, "n");

  for (let row = 19; row <= 80; row += 1) {
    writeCell(sheet, `F${row}`, "", "s");
    writeCell(sheet, `G${row}`, "", "n");
  }

  inputs.npcs.forEach((npc, index) => {
    const row = 19 + index;
    writeCell(sheet, `F${row}`, npc.name, "s");
    writeCell(sheet, `G${row}`, npc.hearts, "n");
  });

  try {
    xlsxCalc(workbook, { continue_after_error: true, log_error: true });
  } catch (error) {
    throw new Error("The browser spreadsheet engine failed to calculate this workbook.");
  }

  const rows = [];
  for (let row = 21; row <= 220; row += 1) {
    const additionalSteps = readSheetNumber(sheet, `I${row}`);
    const totalSteps = readSheetNumber(sheet, `J${row}`);
    const dish = readSheetString(sheet, `K${row}`);

    if (additionalSteps === 0 && totalSteps === 0 && !dish && row > 21) {
      break;
    }

    rows.push({
      additionalSteps,
      totalSteps: totalSteps || (inputs.stepsTaken + additionalSteps),
      dish,
      npcGiftInMail: readSheetString(sheet, `L${row}`),
      giftSucceeds: readSheetString(sheet, `M${row}`),
      dailyLuck: String(readSheetValue(sheet, `N${row}`) ?? ""),
      weather: readSheetString(sheet, `O${row}`),
      gingerIslandWeather: readSheetString(sheet, `P${row}`),
    });
  }

  return { rows };
}

function renderInputs(inputs) {
  inputGrid.innerHTML = "";
  const items = [
    ["Season", inputs.season],
    ["Current Day", inputs.currentDay],
    ["Total Days Played", inputs.totalDaysPlayed],
    ["Steps Taken", inputs.stepsTaken],
    ["Game Seed", inputs.gameSeed],
    ["Debris Tomorrow", inputs.debrisTomorrowFlag],
    ["NPC Count", inputs.npcs.length],
    ["Furnaces", inputs.machines.Furnaces],
    ["Crystalariums", inputs.machines.Crystalariums],
    ["Jars", inputs.machines.Jars],
    ["Kegs", inputs.machines.Kegs],
    ["Tappers", inputs.machines.Tappers],
    ["Bee Houses", inputs.machines["Bee Houses"]],
    ["Oil Makers", inputs.machines["Oil Makers"]],
    ["Biome Balance", inputs.specialRequests.biomeBalance ? "Yes" : "No"],
    ["Crop Order", inputs.specialRequests.cropOrder ? "Yes" : "No"],
    ["Robin's Rush", inputs.specialRequests.robinsResourceRush ? "Yes" : "No"],
    ["Island Ingredients", inputs.specialRequests.islandIngredients ? "Yes" : "No"],
    ["Season Transition", inputs.isSeasonTransitionNight ? "Yes" : "No"],
    ["Relevant Weeds", inputs.relevantWeedCount],
  ];

  for (const [label, value] of items) {
    const wrapper = document.createElement("dl");
    wrapper.className = "kv-item";
    wrapper.innerHTML = `<dt>${label}</dt><dd>${escapeHtml(String(value))}</dd>`;
    inputGrid.append(wrapper);
  }

  warningsList.innerHTML = "";
  for (const warning of inputs.warnings) {
    const li = document.createElement("li");
    li.textContent = warning;
    warningsList.append(li);
  }

  npcBody.innerHTML = "";
  if (!inputs.npcs.length) {
    npcBody.innerHTML = "<tr><td colspan=\"2\" class=\"empty-cell\">No NPC entries were found.</td></tr>";
  } else {
    for (const npc of inputs.npcs) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${escapeHtml(npc.name)}</td><td>${npc.hearts}</td>`;
      npcBody.append(tr);
    }
  }
}

function renderResults(rows) {
  resultsBody.innerHTML = "";
  if (!rows.length) {
    resultsBody.innerHTML = "<tr><td colspan=\"8\" class=\"empty-cell\">No results were produced.</td></tr>";
    return;
  }

  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.additionalSteps}</td>
      <td>${row.totalSteps}</td>
      <td>${escapeHtml(row.dish)}</td>
      <td>${escapeHtml(row.npcGiftInMail)}</td>
      <td>${escapeHtml(row.giftSucceeds)}</td>
      <td>${escapeHtml(row.dailyLuck)}</td>
      <td>${escapeHtml(row.weather)}</td>
      <td>${escapeHtml(row.gingerIslandWeather)}</td>
    `;
    resultsBody.append(tr);
  }
}

async function downloadDefaultSpreadsheet(filename) {
  const response = await fetch(DEFAULT_SPREADSHEET_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error("Failed to fetch the default spreadsheet.");
  }

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function writeCell(sheet, address, value, type) {
  const cell = sheet[address] || {};
  cell.t = type;
  cell.v = value;
  delete cell.w;
  sheet[address] = cell;
}

function readSheetValue(sheet, address) {
  return sheet[address]?.v ?? "";
}

function readSheetNumber(sheet, address) {
  const value = Number(readSheetValue(sheet, address));
  return Number.isFinite(value) ? value : 0;
}

function readSheetString(sheet, address) {
  const value = readSheetValue(sheet, address);
  return value == null ? "" : String(value).trim();
}

function firstChildByLocalName(node, localName) {
  return [...node.children].find((child) => child.localName === localName) || null;
}

function childrenByLocalName(node, localName) {
  return [...node.children].filter((child) => child.localName === localName);
}

function findFirst(node, localName) {
  return [...node.getElementsByTagName("*")].find((child) => child.localName === localName) || null;
}

function readText(node, childName, fallback = "") {
  return firstChildByLocalName(node, childName)?.textContent?.trim() || fallback;
}

function readTextAny(node, localName, fallback = "") {
  return findFirst(node, localName)?.textContent?.trim() || fallback;
}

function readInt(node, childName) {
  return Number.parseInt(readText(node, childName, "0"), 10) || 0;
}

function readIntAny(node, localName) {
  return Number.parseInt(readTextAny(node, localName, "0"), 10) || 0;
}

function readBool(node, childName) {
  return readText(node, childName, "false") === "true";
}

function hasAncestorNamed(node, names) {
  const set = new Set(names);
  let current = node.parentElement;
  while (current) {
    if (set.has(current.localName)) {
      return true;
    }
    current = current.parentElement;
  }
  return false;
}

function minutesUntilNextSixAm(timeOfDay) {
  if (!timeOfDay || timeOfDay <= 0) {
    return 24 * 60;
  }
  const hours = Math.floor(timeOfDay / 100);
  const minutes = timeOfDay % 100;
  const currentMinutes = (hours * 60) + minutes;
  const nextSixAm = (24 * 60) + (6 * 60);
  return Math.max(0, nextSixAm - currentMinutes);
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}
