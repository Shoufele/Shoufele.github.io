const DEFAULT_SPREADSHEET_URL = "./assets/Spreadsheet.Default.xlsx";
const DEFAULT_FILENAME = "Copy of Stardew Step Manipulation V1.6.xlsx";

const input = document.getElementById("spreadsheet-input");
const dropzone = document.getElementById("dropzone");
const fileCard = document.getElementById("file-card");
const fileName = document.getElementById("file-name");
const clearButton = document.getElementById("clear-button");
const restoreButton = document.getElementById("restore-button");
const defaultButton = document.getElementById("default-button");
const resultText = document.getElementById("result-text");
const statusPill = document.getElementById("status-pill");

let selectedFile = null;

input.addEventListener("change", () => {
  updateSelectedFile(input.files?.[0] ?? null);
});

clearButton.addEventListener("click", () => {
  input.value = "";
  updateSelectedFile(null);
});

restoreButton.addEventListener("click", async () => {
  await downloadDefaultSpreadsheet(selectedFile?.name || DEFAULT_FILENAME);
});

defaultButton.addEventListener("click", async () => {
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

  input.files = event.dataTransfer.files;
  updateSelectedFile(file);
});

function updateSelectedFile(file) {
  selectedFile = file;

  if (!file) {
    fileCard.hidden = true;
    fileName.textContent = "";
    statusPill.textContent = "Ready";
    resultText.textContent = "The default workbook is served from docs/assets/Spreadsheet.Default.xlsx.";
    return;
  }

  fileCard.hidden = false;
  fileName.textContent = file.name;
  statusPill.textContent = "File loaded";
  resultText.textContent = "Restore will download a fresh default workbook using the same filename.";
}

async function downloadDefaultSpreadsheet(filename) {
  try {
    statusPill.textContent = "Downloading";
    resultText.textContent = "Fetching the pristine spreadsheet...";

    const response = await fetch(DEFAULT_SPREADSHEET_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
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

    statusPill.textContent = "Done";
    resultText.textContent = `Downloaded restored spreadsheet as "${filename}".`;
  } catch (error) {
    console.error(error);
    statusPill.textContent = "Error";
    resultText.textContent = "Failed to fetch the bundled spreadsheet. Check that docs/assets/Spreadsheet.Default.xlsx is published on GitHub Pages.";
  }
}
