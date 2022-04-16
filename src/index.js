import { ERPExcelWorkbook } from './assets/js/excel.js';


export class excelInfo {
    constructor(data) {
        this.json_columns = data?.json_columns;
        this.json_data = data?.json_data;

        this.workbook = new ERPExcelWorkbook();
        this.workbook.excelExcelWorkbookName = data?.fileName;;
    };

    initDom = () => {
        if (!this.json_columns) {
            this.json_columns = Object.keys(this.json_data[0]);
        }

        // this.workbook.excelExcelWorkbookName = this.fileName;
        const sheet = this.workbook.addSheet(); //添加一页Excel
        this.addTitle(sheet);
        this.addTableData(sheet);
    };

    excelExport = (callback) => {
        this.workbook.excelExport(isDown => {
            console.log(isDown)
            callback(isDown);
        })
    }

    // const row = sheet.addRow(); //add 行
    setRowData = (cell, cellIndex, dataX, dataY) => {
        // { text: '' }
        // { text: '', image: dataY.pictures }
        if (
            dataY[dataX] &&
            typeof dataY[dataX] === 'string' &&
            (
                dataY[dataX].indexOf('.png') > -1 ||
                dataY[dataX].indexOf('.jpg') > -1 ||
                dataY[dataX].indexOf('.jpeg') > -1 ||
                dataY[dataX].indexOf('.gif') > -1
            )
        ) {
            cell.value = { text: '', image: dataY[dataX] };
        } else {
            cell.value = { text: dataY[dataX] };
        }
    }

    addTitle = (sheet) => {
        const row = sheet.addRow();
        for (let x = 0; x < this.json_columns.length; ++x) {
            const dataX = this.json_columns[x];
            const cell = row.addCell();
            cell.value = { text: dataX };
        }
    };

    addTableData = (sheet) => {
        for (let y = 0; y < this.json_data.length; ++y) {
            let dataY = this.json_data[y];
            let row = sheet.addRow();
            console.log(row, 'row');
            for (let x = 0; x < this.json_columns.length; ++x) {
                let dataX = this.json_columns[x];
                let cell = row.addCell();
                this.setRowData(cell, x, dataX, dataY)
            }
        }
    };
}

// json_data = [
// {
//     "name": "花金妹",
//         "address": "1177弄",
//             "照片地址": "https://img-blog.csdnimg.cn/20200204233620255.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxMzU5MDUx,size_16,color_FFFFFF,t_70",
//                 "照片长度": 1,
//                     "药名": " ",
//                         "联系方式": "13817930730",
//                             "门牌": 1,
//                                 "户室": 204
// },
        //     {
        //         "name": "花金妹",
        //         "address": "1177弄",
        //         // "照片地址": "https://pubuserqiniu.paperol.cn/158627462_20_q19_ZFVLhcH0UeXOQQSMVo9dw.png?attname=20_19_splicing.png",
        //         "照片地址": "https://img-blog.csdnimg.cn/20200204233620255.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxMzU5MDUx,size_16,color_FFFFFF,t_70",
        //         "照片长度": 1,
        //         "药名": " ",
        //         "联系方式": "13817930730",
        //         "门牌": 1,
        //         "户室": 204
        //     },
        //     {
        //         "name": "花金妹",
        //         "address": "1177弄",
        //         // "照片地址": "https://pubuserqiniu.paperol.cn/158627462_20_q19_ZFVLhcH0UeXOQQSMVo9dw.png?attname=20_19_splicing.png",
        //         "照片地址": "https://img-blog.csdnimg.cn/20200204233620255.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxMzU5MDUx,size_16,color_FFFFFF,t_70",
        //         "照片长度": 1,
        //         "药名": " ",
        //         "联系方式": "13817930730",
        //         "门牌": 1,
        //         "户室": 204
        //     },
        // ];

        // json_columns = Object.keys(json_data[0]);
