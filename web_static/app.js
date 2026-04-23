const form = document.querySelector("#uploadForm");
const fileInput = document.querySelector("#fileInput");
const fileLabel = document.querySelector("#fileLabel");
const dropZone = document.querySelector("#dropZone");
const statusText = document.querySelector("#status");
const submitButton = document.querySelector("#submitButton");

function setStatus(message, type = "") {
  statusText.textContent = message;
  statusText.className = `status ${type}`.trim();
}

function selectedFile() {
  return fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
}

fileInput.addEventListener("change", () => {
  const file = selectedFile();
  fileLabel.textContent = file ? file.name : "Choose Excel file";
  setStatus("");
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragging");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");

  if (!event.dataTransfer.files.length) return;
  fileInput.files = event.dataTransfer.files;
  fileLabel.textContent = event.dataTransfer.files[0].name;
  setStatus("");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const file = selectedFile();
  if (!file) {
    setStatus("Please choose an Excel file.", "error");
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  submitButton.disabled = true;
  setStatus("Calculating workbook...");

  try {
    const response = await fetch("/calculate", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      let message = "Could not calculate the workbook.";
      try {
        const error = await response.json();
        message = error.detail || message;
      } catch {
        message = await response.text();
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/i);
    const downloadName = match ? match[1] : "calculated_output.xlsx";

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = downloadName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    setStatus("Output generated and downloaded.", "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    submitButton.disabled = false;
  }
});
