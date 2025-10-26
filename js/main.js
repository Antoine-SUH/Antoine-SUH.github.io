// === Variables globales ===
let currentDoc = null;
let currentArrayBuffer = null;
let zip = null;

// === Sélecteurs ===
const fileInput = document.getElementById("fileInput");
const loadButton = document.getElementById("loadButton");
const detectButton = document.getElementById("detectButton"); // nouveau bouton
const messageBox = document.getElementById("messageBox");
const fieldList = document.getElementById("fieldList");

// === Fonction : afficher un message utilisateur ===
function showMessage(msg, type = "info") {
  messageBox.innerText = msg;
  messageBox.style.display = "block";
  messageBox.style.color = type === "error" ? "red" : "#333";
}

// === Fonction : masquer le message ===
function clearMessage() {
  messageBox.innerText = "";
  messageBox.style.display = "none";
}

// === Fonction : charger un fichier Word ===
loadButton.addEventListener("click", async () => {
  const file = fileInput.files[0];
  if (!file) {
    showMessage("Veuillez sélectionner un fichier Word (.docx) d'abord.", "error");
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    currentArrayBuffer = arrayBuffer;
    zip = new PizZip(arrayBuffer);
    showMessage(`✅ Document "${file.name}" chargé avec succès.`);
  } catch (err) {
    console.error(err);
    showMessage("Erreur lors du chargement du document.", "error");
  }
});

// === Fonction : détecter les balises {{...}} ===
detectButton.addEventListener("click", () => {
  if (!zip) {
    showMessage("Veuillez d'abord charger un document.", "error");
    return;
  }

  try {
    const doc = new window.docxtemplater().loadZip(zip);
    const text = doc.getFullText();

    const matches = [...text.matchAll(/{{(.*?)}}/g)];
    if (matches.length === 0) {
      showMessage("Aucun champ {{...}} détecté dans le document.");
      return;
    }

    clearMessage();
    fieldList.innerHTML = "";
    matches.forEach((match, i) => {
      const fieldName = match[1].trim();
      const li = document.createElement("li");
      li.textContent = `${i + 1}. ${fieldName}`;
      fieldList.appendChild(li);
    });

    showMessage(`✅ ${matches.length} champ(s) détecté(s).`);
  } catch (err) {
    console.error(err);
    showMessage("Erreur lors de la détection des champs.", "error");
  }
});
