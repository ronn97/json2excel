import { ERPExcelWorkbook } from './excel.js';

const workbook = new ERPExcelWorkbook();
workbook.excelExcelWorkbookName = 'Excel文件名';

const sheet = workbook.addSheet(); //添加一页Excel

const json_data = [
    {
        "name": "花金妹",
        "address": "1177弄",
        "照片地址": "https://pubuserqiniu.paperol.cn/158627462_20_q19_ZFVLhcH0UeXOQQSMVo9dw.png?attname=20_19_splicing.png",
        "照片长度": 1,
        "药名": " ",
        "联系方式": "13817930730",
        "门牌": 1,
        "户室": 204
    }
];

const json_columns = Object.keys(json_data[0]);

// const row = sheet.addRow(); //add 行

setRowData = (cell, cellIndex, dataX, dataY) => {
    // { text: '' }
    // { text: '', image: dataY.pictures }
    if (dataY[dataX].indexOf('.png') > -1 ||
        dataY[dataX].indexOf('.jpg') > -1 ||
        dataY[dataX].indexOf('.jpeg') > -1 ||
        dataY[dataX].indexOf('.gif') > -1
    ) {
        cell.value = { text: '', image: dataY[dataX] };
    } else {
        cell.value = { text: dataY[dataX] };
    }
};

for (let y = 0; y < json_data.length; ++y) {
    let dataY = json_data[y];
    let row = sheet.addRow();
    console.log(row, 'row');
    for (let x = 0; x < json_columns.length; ++x) {
        let dataX = json_columns[x];
        let cell = row.addCell();
        this.setRowData(cell, cellIndex, dataX, dataY)
    }
}

// for (let x = 0; x < json_data.length; ++x) {
//     const dataX = json_data[x];
//     const cell = row.addCell(); //add 单元格
//     this.setRowData(cell, x, json_data, dataY)
// }
