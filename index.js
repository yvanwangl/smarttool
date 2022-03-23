const XLSX = require('xlsx');
const fs = require('fs');
const workbook = XLSX.readFile('./all.xlsx');
const types = [
  '旋转变压器',
  '自整角',
  '电感移相器',
  '测速',
  '直流',
  '交流',
  '异步',
  '同步',
  '伺服',
  '力矩',
  '步进',
  '永磁',
  '机组',
  '电机扩大机',
  '特种电机及其它',
];
const map = { '无类别': [] };
const nameDup = {};
types.forEach(key => map[key] = []);
workbook.SheetNames.forEach((sheetName) => {
  const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  // console.log(json);
  json.forEach((item, index) => {
    let mainKey = '';
    let values = []
    Object.keys(item).forEach(key => {
      if (key === '商品') {
        let index = -1;
        if (!nameDup[item['商品']]) {
          nameDup[item['商品']] = true;
          types.forEach((type, idx) => {
            if (item[key].indexOf(type) !== -1) {
              map[type].push({ '类型': type, ...item });
              index = idx;
            }
          });
          if (index === -1) {
            map['无类别'].push({ '类型': '无类别', ...item });
          };
        }
      }
    });
  });
});
const wb = XLSX.utils.book_new();
const allItems = Object.keys(map).reduce((acc, cur) => {
  return [...acc, ...map[cur]];
}, []);
console.log(allItems.length);
const ws = XLSX.utils.json_to_sheet(allItems);
XLSX.utils.book_append_sheet(wb, ws, '全部商品');
XLSX.writeFile(wb, 'out.xlsx');

// console.log(Object.keys(map).length)
const result = {};

fs.writeFileSync('./result.json', JSON.stringify(result, null, 2));
// console.log(result);



