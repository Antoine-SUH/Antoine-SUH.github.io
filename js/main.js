let zip, doc, templateName = "", fieldValues = {};

function getBaseName(filename) {
  return filename.substring(0, filename.lastIndexOf(".")) || filename;
}

document.getElementById("loadBtn").addEventListener("click", () => {
  const fileInput = document.getElementById("upload");
  if (!fileInput.files.length) return alert("Sélectionnez un fichier .docx");
  const file = fileInput.files[0];
  templateName = getBaseName(file.name);

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      zip = new PizZip(e.target.result);
      doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

      // Extraire les balises {{...}}
      const tags = doc.getFullText().match(/{{(.*?)}}/g);
      if (!tags) return alert("Aucun champ {{...}} détecté dans le document.");

      const uniqueTags = [...new Set(tags.map(t => t.replace(/[{}]/g, '').trim()))];
      displayFields(uniqueTags);
    } catch (error) {
      alert("Erreur de chargement : " + error);
    }
  };
  reader.readAsBinaryString(file);
});

function displayFields(fields) {
  const div = document.getElementById("fields");
  div.innerHTML = "<h3>Champs détectés :</h3>";
  fields.forEach(f => {
    const fieldDiv = document.createElement("div");
    fieldDiv.classList.add("field");
    fieldDiv.innerHTML = `
      <label><b>${f}</b></label><br>
      <input type="text" id="field_${f}" placeholder="Valeur pour ${f}">
    `;
    div.appendChild(fieldDiv);
  });

  document.getElementById("exportFilled").disabled = false;
  document.getElementById("exportTemplate").disabled = false;
  document.getElementById("status").innerText = "Champs prêts à être remplis.";
}

function collectValues() {
  const inputs = document.querySelectorAll("[id^='field_']");
  fieldValues = {};
  inputs.forEach(input => {
    const key = input.id.replace("field_", "");
    fieldValues[key] = input.value;
  });
}

function exportDoc(mode) {
  collectValues();
  const zipCopy = new PizZip(zip.generate({ type: "arraybuffer" }));
  const docCopy = new window.docxtemplater(zipCopy, { paragraphLoop: true, linebreaks: true });
  if (mode === "filled") docCopy.render(fieldValues);

  const out = docCopy.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
  const fileName = mode === "filled"
    ? `${templateName}_modifié.docx`
    : `${templateName}_template.docx`;
  saveAs(out, fileName);
}

document.getElementById("exportFilled").addEventListener("click", () => exportDoc("filled"));
document.getElementById("exportTemplate").addEventListener("click", () => exportDoc("template"));
