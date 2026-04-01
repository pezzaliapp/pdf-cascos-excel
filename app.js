
const { jsPDF } = window.jspdf;

let excelData = [];
let pdfImages = [];

document.getElementById("excelFile").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  excelData = XLSX.utils.sheet_to_json(sheet);
  document.getElementById("status").innerText = "Excel caricato";
});

async function extractPDFImages(file){
  const data = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({data}).promise;
  const images = [];

  for(let i=1;i<=pdf.numPages;i++){
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({scale:1});
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");

    canvas.width = viewport.width;
    canvas.height = viewport.height;

    await page.render({canvasContext:ctx, viewport}).promise;
    images.push(canvas.toDataURL("image/jpeg"));
  }
  return images;
}

document.getElementById("generateBtn").addEventListener("click", async () => {
  document.getElementById("status").innerText = "Processing...";

  const pdf1 = document.getElementById("pdf1").files[0];
  const pdf2 = document.getElementById("pdf2").files[0];

  const images1 = await extractPDFImages(pdf1);
  const images2 = await extractPDFImages(pdf2);

  pdfImages = [...images1, ...images2];

  const doc = new jsPDF();

  let y = 20;

  excelData.forEach((row, i) => {

    if(y > 250){
      doc.addPage();
      y = 20;
    }

    doc.setFontSize(14);
    doc.text(`${row.Codice} - ${row.Nome}`, 10, y);

    y += 6;
    doc.setFontSize(10);
    doc.text(`Prezzo: ${row.Prezzo}`, 10, y);

    y += 6;
    doc.text(`${row.Descrizione || ""}`, 10, y);

    y += 6;

    const img = pdfImages[i % pdfImages.length];

    if(img){
      doc.addImage(img, "JPEG", 10, y, 80, 50);
      y += 55;
    }

    y += 10;
  });

  doc.save("catalogo_professionale.pdf");
  document.getElementById("status").innerText = "PDF creato!";
});
