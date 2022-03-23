const XLSX = require('xlsx');
const fs = require('fs');
const fetch = require('node-fetch');
const { prependOnceListener } = require('process');
const workbook = XLSX.readFile('./province.xlsx');

const cookie = `pgv_pvid=2889114604; pac_uid=0_b38ba7f49fe17; qb_guid=8130664ca882475d839fbde90e719866; Q-H5-GUID=8130664ca882475d839fbde90e719866; pgv_pvi=7470141440; RK=WGYtu91ydq; ptcz=d97b281ff600246681c23de35e3798dcf52d66e2eff5a888764b48e791ffbb99; wxuin=02832398161501; ua_id=X9OtEuXf38SQx9DxAAAAAPhriX0OEufjrkZ7kWlsxAc=; openid2ticket_o-Tn60H1u2NbpsiKpqnXUuFv4xLM=; mm_lang=zh_CN; openid2ticket_oS5xs5S6y_B6P7jLegms5917v_N4=; ptui_loginuin=1012305328@qq.com; openid2ticket_oRRr90PSR5S76dTQeen9y-Bnz8DQ=; uin=o1012305328; skey=@9o0IumglJ; pgv_info=ssid=s7111398016; uuid=86ff6fd5830c6bad8bee499e843f7088; rand_info=CAESIGP9qlG2VssfK4lnYH8xmAOHUMAesBE36fzi5NLMAgAg; slave_bizuin=3849100363; data_bizuin=3849100363; bizuin=3849100363; data_ticket=4OuGBdENnmjs/58wI0iu0EjG09ZN4/iqvFskPmAdw6j0k0EWfQFwAvr5zaSrxzhV; slave_sid=X3V4SklfaWxEUVBRRDRkXzgwRVVfeDE0eExheHJ3X0F4emkyOVhnT21TNGJWcVNmNVNPWWQ0ZEJpWmtoZkRCNWZBcVJJQWQycG10M2w0N2RndGVsRXB0WnVMWE93Uk9VSmFFWWNoRzJSUmZ1S3BjWDF6Ym1QTXBqUnBPeTdFSmw0eXNoeldTSEdRUnJEZFdN; slave_user=gh_d11984fc0fea; xid=afe95c73192838eaa32e40d9c1a6f200; sig=h01ff7e3b7059b6c6febad6f454dac540e8ba6f83debe20acbd79da4751fc98453baf33ea77dc975474; rewardsn=; wxtokenkey=777`;

// https://mp.weixin.qq.com/wxamp/cgi/route?path=%2Fwxopen%2Fuserportrait%3Faction%3Dget_city_distribution%26f%3Djson%26index%3D1%26time_scope%3D1%26province%3D%25E5%25B9%25BF%25E4%25B8%259C%25E7%259C%2581%26token%3D773665774%26lang%3Dzh_CN&token=773665774&lang=zh_CN&random=0.7644531216939072
// https://mp.weixin.qq.com/wxamp/cgi/route?path%3D%2Fwxopen%2Fuserportrait%3Faction%3Dget_city_distribution%26f%3Djson%26index%3D1%26time_scope%3D1%26province%3D%E5%B1%B1%E4%B8%9C%E7%9C%81%26token%3D773665774%26lang%3Dzh_CN%26token%3D773665774%26lang%3Dzh_CN%26random%3D0.6050599811352699

const fetchData = (provinceName) => {
  return fetch(`https://mp.weixin.qq.com/wxamp/cgi/route?path=%2Fwxopen%2Fuserportrait%3Faction%3Dget_city_distribution%26f%3Djson%26index%3D1%26time_scope%3D1%26province%3D${encodeURIComponent(provinceName)}%26token%3D773665774%26lang%3Dzh_CN&token=773665774&lang=zh_CN&random=${Math.random()}`, {
    headers: {
      'cookie': cookie,
      'authority': 'mp.weixin.qq.com',
      'content-type': 'application/json; charset=utf-8'
    },
    method: 'GET',
  })
    .then(res => {
      return res.json();
    }).then((data) => {
      return JSON.parse(data.data_info).result_list[0].line_list.map(item => {
        let cityName;
        let count;
        const data_list = item.data_list.forEach(data => {
          if (data.key === 'key') {
            cityName = data.value;
          }
          if (data.key === 'value') {
            count = data.value;
          }
        });
        return { '城市': cityName, '用户数': count };
      });
    });
}

workbook.SheetNames.forEach((sheetName) => {
  const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  // console.log(json);
  const promiseList = json.map(({ name, value }) => {
    return fetchData(name);
  });
  const wb = XLSX.utils.book_new();
  Promise.all(promiseList).then(list => {
    console.log(list);
    const allItems = list.reduce((acc, cur) => {
      acc = acc.concat(cur);
      return acc;
    }, []).sort((itemA, itemB) => itemB['用户数'] - itemA['用户数']);
    const ws = XLSX.utils.json_to_sheet(allItems);
    XLSX.utils.book_append_sheet(wb, ws, '全部城市');
    XLSX.writeFile(wb, './allCity.xlsx');
  });
});