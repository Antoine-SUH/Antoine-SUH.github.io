let zip, docXml, originalZip, fields = {};

const uploadInput = document.getElementById("upload");
const loadButton = document.getElementById("load-doc");
const detectButton = document.getElementById("detect-tags");
const exportFilledButton = document.getElementById("export-filled");
const exportTemplateButton = document.getElementById("export-template");
const fieldsContainer = document.getElementById("fields");
const status = document.getElementById("status");

// Charger le document
loadButton.addEventListener("click", async () => {
  const file = uploadInput.files[0];
  if (!file) return alert("Sélectionne un fichier .docx d'abord !");

  status.textContent = "Chargement du document...";
  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      originalZip = new PizZip(new Uint8Array(event.target.result));

      // Lecture du XML principal du document Word
      docXml = originalZip.files["word/document.xml"].asText();

      // On nettoie le XML pour regrouper les morceaux de texte
      docXml = docXml.replace(/<\/w:t><w:t[^>]*>/g, "");

      status.textContent = "✅ Document chargé avec succès !";
      document.getElementById("detect-section").classList.remove("hidden");
    } catch (error) {
      console.error("Erreur :", error);
      status.textContent = "❌ Erreur lors du chargement du document.";
    }
  };

  reader.readAsArrayBuffer(file);
});

// Détecter les balises {{...}}
detectButton.addEventListener("click", () => {
  if (!docXml) return alert("Charge d'abord un document !");
  
  const regex = /\{\{(.*?)\}\}/g;
  const matches = [...docXml.matchAll(regex)];
  fieldsContainer.innerHTML = "";

  if (matches.length === 0) {
    fieldsContainer.innerHTML = "<p>Aucune balise {{...}} détectée.</p>";
    return;
  }

  matches.forEach(match => {
    const key = match[1].trim();
    if (!fields[key]) fields[key] = "";

    const div = document.createElement("div");
    div.classList.add("field");
    div.innerHTML = `
      <label>${key}</label>
      <input type="text" id="field-${key}" placeholder="Valeur pour ${key}" value="${fields[key]}"/>
    `;
    fieldsContainer.appendChild(div);
  });

  fieldsContainer.innerHTML += "<p>🟢 Balises détectées et prêtes à être remplies.</p>";
});

// Exporter le document rempli
exportFilledButton.addEventListener("click", () => {
  if (!originalZip || !docXml) return alert("Charge un document d'abord !");
  
  let modifiedXml = docXml;
  Object.keys(fields).forEach(key => {
    const val = document.getElementById(`field-${key}`)?.value || "";
    const regex = new RegExp(`\\{\\{\\s*${key}\\s*\\}\\}`, "g");
    modifiedXml = modifiedXml.replace(regex, val);
  });

  const newZip = new PizZip(originalZip);
  newZip.file("word/document.xml", modifiedXml);

  const out = newZip.generate({ type: "blob" });
  saveAs(out, "document_modifié.docx");
});

// Exporter le modèle (avec balises intactes)
exportTemplateButton.addEventListener("click", () => {
  if (!originalZip) return alert("Charge un document d'abord !");
  const out = originalZip.generate({ type: "blob" });
  saveAs(out, "document_template.docx");
});
