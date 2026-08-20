/* Finance Forward — "Export to Word" button.
   Wraps the page's .article content in a minimal Word-compatible HTML
   shell (the standard application/msword MHTML trick) and downloads it
   as a .doc file. No server, no build step — everything happens in the
   browser. Works on any page that has one .article element and an
   .export-btn button with data-export-name for the filename.
*/
(function () {
  function absolutize(url) {
    try { return new URL(url, document.baseURI).href; }
    catch (e) { return url; }
  }

  function buildDoc(article, title) {
    var clone = article.cloneNode(true);

    // Drop interactive/non-content chrome that shouldn't end up in the doc.
    clone.querySelectorAll('.export-bar').forEach(function (n) { n.remove(); });

    // Word needs absolute image URLs (it fetches them itself).
    clone.querySelectorAll('img').forEach(function (img) {
      img.setAttribute('src', absolutize(img.getAttribute('src')));
      img.style.maxWidth = '100%';
    });

    var bodyHtml = clone.innerHTML;

    return '' +
      '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:w="urn:schemas-microsoft-com:office:word" ' +
      'xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta charset="utf-8">' +
      '<title>' + title + '</title>' +
      '<!--[if gte mso 9]><xml><w:WordDocument>' +
      '<w:View>Print</w:View><w:Zoom>100</w:Zoom>' +
      '</w:WordDocument></xml><![endif]-->' +
      '<style>' +
      'body{font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.5;color:#22303c;margin:1in}' +
      'h1{font-size:22pt;color:#1b3a5b;margin-bottom:.3em}' +
      'h4{font-size:10pt;text-transform:uppercase;letter-spacing:.05em;color:#6b7a86}' +
      'p{margin:0 0 12pt}' +
      'figure{margin:16pt 0}' +
      'figcaption{font-style:italic;font-size:9pt;color:#22303c;text-align:center}' +
      '.src{display:block;font-style:normal;color:#6b7a86;font-size:8pt}' +
      '.nudge{background:#fdf1ec;padding:10pt 14pt;border-left:3pt solid #c0492f;margin:14pt 0}' +
      '.subjectbar{background:#f3faf8;padding:10pt 14pt;margin:0 0 14pt}' +
      '.sources ul,.steps{margin:0 0 12pt}' +
      '.alt-links,.pager,.lang-toggle,.crumb{display:none}' +
      '</style></head><body>' +
      bodyHtml +
      '</body></html>';
  }

  document.addEventListener('click', function (e) {
    var btn = e.target.closest('.export-btn');
    if (!btn) return;
    var article = document.querySelector('.article');
    if (!article) return;
    var title = document.title.replace(/\s*—\s*Finance Forward\s*$/, '');
    var filename = (btn.getAttribute('data-export-name') || 'finance-forward-issue') + '.doc';
    var html = buildDoc(article, title);
    var blob = new Blob(['﻿', html], { type: 'application/msword' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  });
})();
