document.getElementById('doneButton').addEventListener('click', async () => {
    const tab = (await chrome.tabs.query({ active: true, currentWindow: true }))[0];
    if (tab.url.startsWith('chrome://')) {
      alert('This extension cannot be run on chrome:// URLs.');
      return;
    }
  
    const checkAllLinks = document.getElementById('checkAllLinks').checked;
    const checkBrokenLinks = document.getElementById('checkBrokenLinks').checked;
    const checkLocalLanguageLinks = document.getElementById('checkLocalLanguageLinks').checked;
    const checkAllDetails = document.getElementById('checkAllDetails').checked;
  
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: checkLinks,
        args: [checkAllLinks, checkBrokenLinks, checkLocalLanguageLinks, checkAllDetails]
      });
  
      // Store results in localStorage for later use by the preview button
      localStorage.setItem('linkResults', JSON.stringify(results[0].result));
      alert('Completed');
    } catch (error) {
      console.error(error);
      alert('An error occurred while checking links.');
    }
  });
  
  document.getElementById('previewButton').addEventListener('click', () => {
    const results = JSON.parse(localStorage.getItem('linkResults'));
    if (results) {
      document.getElementById('result').style.display = 'block';
      displayAllLinks(results.allLinks);
      displayBrokenLinks(results.brokenLinks);
      displayLocalLanguageLinks(results.localLanguageLinks);
    } else {
      alert('No data to preview. Please click "Done" first.');
    }
  });
  
  document.getElementById('downloadPdfButton').addEventListener('click', () => {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
  
    doc.text('All Links', 10, 10);
    doc.autoTable({ html: '#allLinksTable', startY: 20 });
    doc.addPage();
    doc.text('Broken Links', 10, 10);
    doc.autoTable({ html: '#brokenLinksTable', startY: 20 });
    doc.addPage();
    doc.text('Local Language Links', 10, 10);
    doc.autoTable({ html: '#localLanguageLinksTable', startY: 20 });
  
    doc.save('links_report.pdf');
  });
  
  document.getElementById('downloadExcelButton').addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
  
    const allLinksTable = document.getElementById('allLinksTable');
    const brokenLinksTable = document.getElementById('brokenLinksTable');
    const localLanguageLinksTable = document.getElementById('localLanguageLinksTable');
  
    const allLinksSheet = XLSX.utils.table_to_sheet(allLinksTable);
    const brokenLinksSheet = XLSX.utils.table_to_sheet(brokenLinksTable);
    const localLanguageLinksSheet = XLSX.utils.table_to_sheet(localLanguageLinksTable);
  
    XLSX.utils.book_append_sheet(wb, allLinksSheet, 'All Links');
    XLSX.utils.book_append_sheet(wb, brokenLinksSheet, 'Broken Links');
    XLSX.utils.book_append_sheet(wb, localLanguageLinksSheet, 'Local Language Links');
  
    XLSX.writeFile(wb, 'links_report.xlsx');
  });
  
  function displayAllLinks(links) {
    let html = '<table><tr><th>Link</th><th>Status</th></tr>';
    links.forEach(link => {
      html += `<tr><td>${link.url}</td><td>${link.status}</td></tr>`;
    });
    html += '</table>';
    document.getElementById('allLinksTable').innerHTML = html;
  }
  
  function displayBrokenLinks(links) {
    let html = '<table><tr><th>Link</th><th>Status</th></tr>';
    links.forEach(link => {
      html += `<tr><td>${highlightPercent20(link.url)}</td><td>${link.status}</td></tr>`;
    });
    html += '</table>';
    document.getElementById('brokenLinksTable').innerHTML = html;
  }
  
  function displayLocalLanguageLinks(links) {
    let html = '<table><tr><th>Link</th><th>Language string</th></tr>';
    links.forEach(link => {
      html += `<tr><td>${link.url}</td><td>${getLocalLanguageString(link.url)}</td></tr>`;
    });
    html += '</table>';
    document.getElementById('localLanguageLinksTable').innerHTML = html;
  }

  function getLocalLanguageString(url) {
    const localLanguageList = [
        'en-us', 'en-au', 'en-ca', 'en-gb', 'en-hk', 'en-ie', 'en-in', 'en-my', 'en-nz', 'en-ph', 'en-sg', 'en-za', 'es-es',
        'es-mx', 'fr-be', 'fr-ca', 'fr-fr', 'it-it', 'ko-kr', 'pt-br', 'de-de', 'ar-sa', 'da-dk', 'fi-fi', 'ja-jp', 'nb-no', 
        'nl-be', 'nl-nl', 'zh-cn'
    ];
    for (const language of localLanguageList) {
        if (url.includes(language)) {
            return language;
        }
    }
    return 'unknown';
}

function highlightPercent20(url) {
  return url.replace(/%20/g, '<span style="color: red;">%20</span>');
}
  
  async function checkLinks(checkAllLinks, checkBrokenLinks, checkLocalLanguageLinks, checkAllDetails) {
    const allLinks = [];
    const brokenLinks = [];
    const localLanguageLinks = [];
    const localLanguageList = [
    'en-us','en-au','en-ca','en-gb','en-hk','en-ie','en-in','en-my','en-nz','en-ph','en-sg','en-za','es-es',
    'es-mx','fr-be','fr-ca','fr-fr','it-it','ko-kr','pt-br','de-de','ar-sa','da-dk','fi-fi','ja-jp','ja-jp',
    'nb-no','nl-be','nl-nl','zh-ch'];
  
    const links = Array.from(document.querySelectorAll('a')).map(link => link.href);
    
    for (const url of links) {
     
        const response = await fetch(url);
        const status = response.status;
        if (checkAllLinks || checkAllDetails) allLinks.push({ url, status });
        if ((checkBrokenLinks || checkAllDetails) && (status === 400 || status === 404 || status === 410 || url.includes('%20'))) brokenLinks.push({ url, status });
        if ((checkLocalLanguageLinks || checkAllDetails) && localLanguageList.some(language => url.includes(language))) {
          localLanguageLinks.push({ url });
        }
      } 
  
    return { allLinks, brokenLinks, localLanguageLinks };
  }
  