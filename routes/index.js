var express = require('express');
const path = require("path");
var router = express.Router();
const request = require('request')
const {response} = require("express");
const axios = require("axios");
const ExcelJS = require('exceljs');


//页面跳转
// router.get('/', function (req, res) {
//   var path = require('path');
//   path= path.resolve(__dirname,'../public/table.html')
//   console.log(path)
//   res.sendFile(path);
// })

// integrateAndSort();

OutPut()

// 执行输出文件
async function OutPut(){
  const partialResultArray = await integrateAndSort();
  const outputPath = path.join('C://Users/MOYEE/Desktop', 'output.xlsx');
  await writeToExcel(partialResultArray,outputPath)
}

// 导出文件
async function writeToExcel(dataArray, outputPath) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  // 添加表头
  worksheet.addRow(['物品名称', '最低卖单价', '最高收单价', '期望收益率']);

  // 添加数据
  dataArray.forEach(item => {
    const { '物品名称': name, '最低卖单价': minSell, '最高收单价': maxBuy, '期望收益率': moneyRate } = item;
    worksheet.addRow([name, minSell, maxBuy, moneyRate]);
  });

  // 将数据写入文件
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Excel file created successfully at ${outputPath}`);
}










//最终处理用函数
async function integrateAndSort() {
  const dataArray = await readXLS(); // 调用readXLS获取数据数组
  const resultArray = [];
  let i =1
  let ii = dataArray.length
  for (const item of dataArray) {
    const { type, name } = item;
    const profitMarginResult = await getProfitMargin(type, name);
    console.log('正在处理第'+i+'条数据，总计'+ ii+'条')
    i=i+1;
    // 将type、name和MoneyRate整合到新数组中
    resultArray.push({
      '物品ID':type,
      '物品名称':name,
      '最低卖单价':profitMarginResult.MinSell,
      '最高收单价':profitMarginResult.MaxBuy,
      MoneyRate: profitMarginResult.MoneyRate,
    });
  }

  console.log('所有数据导入完毕，正在计算。。。')
  // 对MoneyRate字段进行降序排序
  resultArray.sort((a, b) => b.MoneyRate - a.MoneyRate);

  const partialResultArray = resultArray.map(item => ({
    '物品名称': item['物品名称'],
    '最低卖单价': item['最低卖单价'],
    '最高收单价': item['最高收单价'],
    '期望收益率': item.MoneyRate+' %'
  }));

  console.log(partialResultArray);

  // 输出排序后的数组
  // console.log(resultArray);
  return partialResultArray;
}

//读取xls
async function readXLS() {
  const workbook = new ExcelJS.Workbook();
  //实例对象
  const dataArray = [];
  //创建数组
  try {
    await workbook.xlsx.readFile('C:\\Users\\MOYEE\\Desktop\\ID.xlsx');
    //读取文件路径
    // 假设你的数据在第一个工作表中（worksheets数组的第一个元素）
    const worksheet = workbook.worksheets[0];

    // 迭代每一行并获取数据
    worksheet.eachRow((row, rowNumber) => {
      const rowData = {
        type: row.getCell(1).value, // 假设type字段在第一列
        name: row.getCell(2).value, // 假设name字段在第二列
      };

      dataArray.push(rowData);
    });

    // 输出整合后的数据数组
    // console.log(dataArray);
    return dataArray;
  } catch (error) {
    console.error('Error reading XLS file:', error.message);
    return null;
  }
}

//获取物品数据
async function getData(types){
  try{
    const Url = 'https://market.fuzzwork.co.uk/aggregates/?region=60003760&types='+types
    //拼接链接
    const res = await axios.get(Url)
    //访问api
    const resData = res.data
    //获得JSON
    return resData
    //返回数据
  }catch (error){
  }
}

//获得物品利润率
async function getProfitMargin(type,name){
  const Data = await getData(type)
  //获得初始数据
  const MinSell = Data[type].sell.min
  // 最低卖价
  const MaxBuy = Data[type].buy.max
  //最高收价
  const TureMoney = MinSell/1.069-MaxBuy
  //期望实际收益
  const MoneyRate = ((TureMoney/MinSell)*100).toFixed(3)
  //实际收益率
  return{type, MoneyRate, name,MaxBuy,MinSell}
  //返回结果
}
module.exports = router;
