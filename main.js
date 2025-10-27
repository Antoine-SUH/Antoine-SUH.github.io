let zip, doc, content, fields = {};

document.getElementById("load-doc").addEventListener("click", async () => {
  const fileInput = document.getElementById("upload");
  const file = fileInput.files[0];
  if (!file) return alert("Sélectionne un fichier .docx d'abord !");
  
  const status = document.getElementById("status");
  status.textContent = "Chargement du document...";
  
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      zip = new PizZip(event.target.result);
      content = zip.files["word/document.xml"].asText();
      status.textContent = "✅ Document chargé avec succès !";
      document.getElementById("detect-section").classList.remove("hidden");
    } catch (error) {
      console.error(error);
      status.textContent = "❌ Erreur lors du chargement du document.";
    }
  };
  reader.readAsBinaryString(file);
});

document.getElementById("detect-tags").addEventListener("click", () => {
  if (!content) return alert("Charge d'abord un document !");
  
  const regex = /\{\{(.*?)\}\}/g;
  const matches = [...content.matchAll(regex)];
  
  const container = document.getElementById("fields");
  container.innerHTML = "";

  if (matches.length === 0) {
    container.innerHTML = "<p>Aucune balise {{...}} détectée.</p>";
    return;
  }

  matches.forEach(match => {
    const key = match[1].trim();
    if (!fields[key]) fields[key] = "";

    const div = document.createElement("div");
    div.classList.add("field");
    div.innerHTML = `
      <label>${key}</label>
      <input type="text" id="field-${key}" placeholder="Valeur pour ${key}" />
    `;
    container.appendChild(div);
  });

  container.innerHTML += "<p>🟢 Balises détectées et prêtes à être remplies.</p>";
});

document.getElementById("export-filled").addEventListener("click", () => {
  if (!zip || !content) return alert("Charge un document d'abord !");
  
  let modified = content;
  Object.keys(fields).forEach(key => {
    const val = document.getElementById(`field-${key}`)?.value || "";
    const regex = new RegExp(`\\{\\{\\s*${key}\\s*\\}\\}`, "g");
    modified = modified.replace(regex, val);
  });

  zip.file("word/document.xml", modified);
  const out = zip.generate({ type: "blob" });
  saveAs(out, "document_modifié.docx");
});

document.getElementById("export-template").addEventListener("click", () => {
  if (!zip) return alert("Charge un document d'abord !");
  const out = zip.generate({ type: "blob" });
  saveAs(out, "document_template.docx");
});
