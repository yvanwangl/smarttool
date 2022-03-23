const XLSX = require('xlsx');
const fs = require('fs');
const fetch = require('node-fetch');
const workbook = XLSX.readFile('./galf.xlsx');

workbook.SheetNames.forEach((sheetName) => {
  const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  const promiseList = json.map(({ name, value }) => {
    let projectId = value.replace('\n', '');
    return fetch(`http://gulfstream.dataapp.dev.sankuai.com/api/getVersionsByProject?project=${projectId}`, { headers: { 'Authorization': 'Basic YWRtaW46cGFzcw==' }})
    .then(res=> res.json()).then(({ data }) => ({ name, value: projectId, version: data[0] || '0.0.0'}));
  });
  Promise.all(promiseList).then(list => {
    fs.writeFileSync('./galf.json', JSON.stringify({ data: list, status: 0, message: 'success' }, null, 2));
  })
});