// storyboard_print.html → storyboard.pdf (신랑 측 검토용). 실행: node print_pdf.js
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + __dirname + '/storyboard_print.html', { waitUntil: 'load' });
  await p.pdf({ path: __dirname + '/storyboard.pdf', preferCSSPageSize: true, printBackground: true });
  await b.close();
})();
