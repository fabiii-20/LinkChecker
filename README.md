https://www.microsoft.com/en-us/nonprofits?sr20=msc#starlight-children%E2%80%99s-foundation
The site for testing

 //Download Excel
document.getElementById('downloadExcelButton').addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  wb.Props = {
    Title: "Link Report",
    Subject: "Link Checker",
    Author: "Your Name",
    CreatedDate: new Date()
  };

  // Convert all links to worksheet
  const allLinksData = [["URL", "Status"]].concat(
    Array.from(document.querySelectorAll('#allLinksTable tr')).map(row => {
      return Array.from(row.cells).map(cell => cell.textContent);
    })
  );
  const allLinksSheet = XLSX.utils.aoa_to_sheet(allLinksData);
  XLSX.utils.book_append_sheet(wb, allLinksSheet, "All Links");

  // Convert broken links to worksheet
  const brokenLinksData = [["URL", "Status"]].concat(
    Array.from(document.querySelectorAll('#brokenLinksTable tr')).map(row => {
      return Array.from(row.cells).map(cell => cell.textContent);
    })
  );
  const brokenLinksSheet = XLSX.utils.aoa_to_sheet(brokenLinksData);
  XLSX.utils.book_append_sheet(wb, brokenLinksSheet, "Broken Links");

  // Convert local language links to worksheet
  const localLanguageLinksData = [["URL", "Language String"]].concat(
    Array.from(document.querySelectorAll('#localLanguageLinksTable tr')).map(row => {
      return Array.from(row.cells).map(cell => cell.textContent);
    })
  );
  const localLanguageLinksSheet = XLSX.utils.aoa_to_sheet(localLanguageLinksData);
  XLSX.utils.book_append_sheet(wb, localLanguageLinksSheet, "Local Language Links");

  XLSX.writeFile(wb, 'links_report.xlsx');
});

//Download PDF
document.getElementById('downloadPdfButton').addEventListener('click', () => {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  // Adding "All Links" table to the PDF
  doc.text('All Links', 10, 10);
  doc.autoTable({
    head: [['URL', 'Status']],
    body: Array.from(document.querySelectorAll('#allLinksTable tr')).map(row => {
      return Array.from(row.cells).map(cell => cell.textContent);
    }),
    startY: 20
  });

  // Adding "Broken Links" table to the PDF
  doc.text('Broken Links', 10, doc.lastAutoTable.finalY + 10);
  doc.autoTable({
    head: [['URL', 'Status']],
    body: Array.from(document.querySelectorAll('#brokenLinksTable tr')).map(row => {
      return Array.from(row.cells).map(cell => cell.textContent);
    }),
    startY: doc.lastAutoTable.finalY + 20
  });

  // Adding "Local Language Links" table to the PDF
  doc.text('Local Language Links', 10, doc.lastAutoTable.finalY + 10);
  doc.autoTable({
    head: [['URL', 'Language String']],
    body: Array.from(document.querySelectorAll('#localLanguageLinksTable tr')).map(row => {
      return Array.from(row.cells).map(cell => cell.textContent);
    }),
    startY: doc.lastAutoTable.finalY + 20
  });

  doc.save('links_report.pdf');
});
