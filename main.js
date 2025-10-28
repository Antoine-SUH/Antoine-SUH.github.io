let zip, doc, data = {};

const uploadInput = document.getElementById("upload");
const loadButton = document.getElementById("load-doc");
const detectButton = document.getElementById("detect-tags");
const exportFilledButton = document.getElementById("export-filled");
const exportTemplateButton = document.getElementById("export-template");
const fieldsContainer = document.getElementById("fields");
const status = document.getElementById("status");

loadButton.addEventListener("click", async () => {
  const file = uploadInput.files[0];
  if (!file) return alert("Sélectionne un fichier .docx d'abord !");
  
  status.textContent = "Chargement du document...";
  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      zip = new PizZip(new Uint8Array(event.target.result));
      doc = new window.docxtemplater().loadZip(zip);
      status.textContent = "✅ Document chargé avec succès !";
      document.getElementById("detect-section").classList.remove("hidden");
    } catch (error) {
      console.error("Erreur :", error);
      status.textContent = "❌ Erreur lors du chargement du document.";
    }
  };

  reader.readAsArrayBuffer(file);
});

detectButton.addEventListener("click", () => {
  if (!doc) return alert("Charge d'abord un document !");
  try {
    const tags = doc.getFullText().match(/\{\{(.*?)\}\}/g) || [];
    const uniqueTags = [...new Set(tags.map(t => t.replace(/[{}]/g, '').trim()))];

    fieldsContainer.innerHTML = "";
    if (uniqueTags.length === 0) {
      fieldsContainer.innerHTML = "<p>Aucune balise {{...}} détectée.</p>";
      return;
    }

    uniqueTags.forEach(key => {
      if (!data[key]) data[key] = "";
      const div = document.createElement("div");
      div.classList.add("field");
      div.innerHTML = `
        <label>${key}</label>
        <input type="text" id="field-${key}" placeholder="Valeur pour ${key}" value="${data[key]}" />
      `;
      fieldsContainer.appendChild(div);
    });

    fieldsContainer.innerHTML += "<p>🟢 Balises détectées et prêtes à être remplies.</p>";
  } catch (error) {
    console.error("Erreur détection :", error);
    alert("Erreur pendant la détection des balises.");
  }
});

exportFilledButton.addEventListener("click", () => {
  if (!doc) return alert("Charge un document d'abord !");
  try {
    Object.keys(data).forEach(key => {
      data[key] = document.getElementById(`field-${key}`)?.value || "";
    });

    const newDoc = new window.docxtemplater(zip);
    newDoc.render(data);
    const out = newDoc.getZip().generate({ type: "blob" });
    saveAs(out, "document_modifié.docx");
  } catch (error) {
    console.error("Erreur génération :", error);
    alert("Erreur pendant la génération du document.");
  }
});

exportTemplateButton.addEventListener("click", () => {
  if (!zip) return alert("Charge un document d'abord !");
  const out = zip.generate({ type: "blob" });
  saveAs(out, "document_template.docx");
});
