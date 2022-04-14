import Util from "./util.js";
const Excel = require('exceljs');

export class ERPExcelCell {

    row = null;
    key = null;

    block = null;

    excelCell = null;
    beMerged = false;


    // excel cell value
    /*
        text,
        hyperlink
        tooltip,
        image // URL
        imageBase64 // base64
     */
    value = null;

    // excel font
    font = { color: { 'argb': 'FFFFFFFF' } };

    // excel fill
    fill = {
        type: 'pattern',
        pattern: 'darkTrellis',
        fgColor: { argb: "FF1890FF" },
        bgColor: { argb: 'FF1890FF' }
    };

    excelBorder = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    // excel text alignment
    alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true
    };

    // 默认td样式
    ui_defaultTdStyle = {
        border: '1px solid #ddd',
        verticalAlign: 'middle'
    };

    // td下文字样式
    ui_tdTextStyle = {
        height: '50px',
        minWidth: '100px',
        // display: 'flex',
        // flexDirection: 'column',
        // justifyContent: 'center',
        textAlign: 'center',
        maxWidth:'120px',
        padding:'0 5px',
        lineHeight: '50px',
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap'
    };

    // td图片样式
    ui_tdImgStyle = {
        width: '50px',
        height: '50px',
        display: 'block',
        margin: '0 auto',
    };

    // td URL样式
    ui_tdUrlStyle = {
        maxWidth: '100px',
        display: 'block',
        // height: '50px',
        // display: "-webkit-box",
        // webkitBoxOrient: 'vertical',
        // webkitLineClamp: '3',
        // overflow: 'hidden',
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap'
    };

    constructor(props) {
        this.row = props.row;
        this.value = props.value;
        this.key = props.key;
    }

    get excelRow() {
        return this.row.excelRow;
    }

    get excelWorksheet() {
        return this.row.sheet.excelWorksheet;
    }

    get excelWorkbook() {
        return this.row.sheet.workbook.excelWorkbook;
    }

    get type() {
        if (this.isImage()) return 'img';
        if (this.isURL()) return 'url';
        return 'text';
    }

    get sheet() {
        return this.row.sheet;
    }

    get left() {
        return this.sheet.getCell(this.x - 1, this.y);
    }

    get right() {
        return this.sheet.getCell(this.x + 1, this.y);
    }

    get up() {
        return this.sheet.getCell(this.x, this.y - 1);
    }

    get down() {
        return this.sheet.getCell(this.x, this.y + 1);
    }

    get x() {
        return this.row.cells.indexOf(this);
    };

    get y() {
        return this.row.y;
    };

    isNeighbor = (cell) => {
        if (this === cell.left || this === cell.right || this === cell.up || this === cell.down) {
            return true;
        }
        return false;
    };

    toJSON = () => {
        return {
            x: this.x,
            y: this.y,
            value: this.value,
            key: this.key
        }
    };

    needShowMeAndBelow = () => {
        let ret = { yes: true, rowSpan: 1 };

        const up = this.up;
        if (!up) {
            ret.yes = true;
        }

        if (this.key && up && (up.key === this.key)) {
            ret.yes = false;
            return ret;
        }

        let cell = this.down;
        while (cell) {
            if (this.key && (cell.key === this.key)) {
                ret.rowSpan++;
            } else {
                break;
            }
            cell = cell.down;
        }

        return ret;
    };

    needShowMeAndRight = () => {
        let ret = { yes: true, colSpan: 1 };

        const left = this.left;
        if (!left) {
            ret.yes = true;
        }

        if (this.key && left && (left.key === this.key)) {
            ret.yes = false;
            return ret;
        }

        let cell = this.right;
        while (cell) {
            if (this.key && (cell.key === this.key)) {
                ret.colSpan++;
            } else {
                break;
            }
            cell = cell.right;
        }

        return ret;
    };

    mergeOneCell = (cell) => {
        if (cell && this.isNeighbor(cell)) {
            if (!this.key) this.key = Util.createUUID();
            cell.key = this.key;
            if (!this.block) this.block = [this];
            if (cell.block) {
                Util.addArrayItems(this.block, cell.block);
            } else {
                Util.addArrayItem(this.block, cell);
            }
            cell.block = this.block;

            const key = this.block[0].key;
            for (let i = 1; i < this.block.length; ++i) {
                const c = this.block[i];
                c.key = key;
            }
        }
    };

    mergeCells = (cells) => {
        for (let i = 0; i < cells.length; ++i) {
            const cell = cells[i];
            cell.key = this.key;
        }
    };

    getContentOfText = () => {
        if (this.value) return this.value.text;
        return '';
    };

    getContentOfImage = () => {
        return this.value.image;
    };

    isImage = () => {
        return this.value && this.value.image !== undefined;
    };

    isURL = () => {
        return this.value && this.value.hyperlink !== undefined;
    };

    getImage = (callback) => {
        if (this.value.image) {
            Util.webGetImage(this.value.image, base64 => {
                this.value.imageBase64 = base64;
                callback(this);
            });
        } else {
            callback(this);
        }
    };

    getContent = () => {
        if (this.isImage()) return this.getContentOfImage();
        return this.getContentOfText();
    };

    getFont = (isExcel) => {
        if (isExcel) {
            return this.font
        } else {
            return null;
        }
    };

    getBorder = (isExcel) => {
        if (isExcel) {
            return this.excelBorder;
        } else {
            return null;
        }
    };

    getFill = (isExcel) => {
        if (isExcel) {
            return this.fill;
        } else {
            return null;
        }
    };

    getAlignment = (isExcel) => {
        if (isExcel) {
            return this.alignment;
        } else {
            return null;
        }
    };

    getIsColumn = () => {
        if (this.y === 0) {
            return true;
        } else {
            return false;
        }
    };

    findLeftMostWithSameKey = () => {
        if (!this.key) return this;

        let cell = this;
        while (cell) {
            if (cell.left && cell.left.key === this.key) {
                cell = cell.left;
            } else {
                break;
            }
        }
        return cell;
    };

    findBlock = () => {

        //....

        // return {
        //     leftTop: xxx
        //     rightDown: yyy
        // }
    };

    excelAddImage = () => {
        if (!this.value || !this.value.imageBase64) return;

        this.value.excelImageId = this.excelWorkbook.addImage({
            base64: this.value.imageBase64,
            extension: 'jpg'
        });
    };

    excelSetCellStyle = () => {
        let x = this.x + 1;
        let y = this.y + 1;
        let excelFill = this.getFill(true);
        let excelBorder = this.getBorder(true);
        let excelAlignment = this.getAlignment(true);
        let excelFont = this.getFont(true);
        const setCell = this.excelWorksheet.getCell(y, x);
        const dobCol = this.excelWorksheet.getColumn(x);
        setCell.border = excelBorder;
        setCell.alignment = excelAlignment;
        if (this.isImage()) {
            dobCol.width = 10;
        } else {
            dobCol.width = 20;
        }
        if (this.getIsColumn()) {
            setCell.fill = excelFill;
            setCell.font = excelFont;
        }
    };

    excelSetImage = () => {
        if (
            (this.value && this.value.excelImageId) ||
            (this.value && (this.value.excelImageId === 0))
        ) {
            this.excelWorksheet.addImage(this.value.excelImageId,
                {
                    tl: { col: this.x + 0.25, row: this.y + 0.2 },
                    ext: { width: 32, height: 32 }
                }
            );
        } else {
            return;
        }
    };

    excelMergeCells = () => {
        if (this.beMerged) return;

        let newBelow = this.needShowMeAndBelow();
        let newRight = this.needShowMeAndRight();
        if (newBelow && newBelow.rowSpan > 1 || newRight && newRight.colSpan > 1) {

            let x = this.x + 1;
            let y = this.y + 1;
            let rowSpan = newBelow.rowSpan === 1 ? 0 : newBelow.rowSpan > 1 ? newBelow.rowSpan - 1 : newBelow.rowSpan;
            let colSpan = newRight.colSpan === 1 ? 0 : newRight.colSpan > 1 ? newRight.colSpan - 1 : newRight.colSpan;
            console.log(y, x,
                y + rowSpan,
                x + colSpan)
            this.excelWorksheet.mergeCells(
                y, x,
                y + rowSpan,
                x + colSpan
            );
            this.beMerged = true;

            for (let i = 0; i < this.block.length; ++i) {
                const c = this.block[i];
                c.beMerged = true;
            }
        }
    };

}

export class ERPExcelRow {

    sheet = null;
    cells = [];

    excelRow = null;

    // tr下标题样式
    ui_rowTitleStyle = {
        backgroundColor: '#1890FF',
        color: '#fff',
        textAlign: 'center'
    };

    constructor(props) {
        this.sheet = props.sheet;
        if(props.cells) {
            for(let i = 0; i < props.cells.length; i++) {
                props.cells[i].row = this;
                let cellObj = Object.assign({}, props.cells[i], { row: this });
                let cell = new ERPExcelCell(cellObj);
                this.cells.push(cell);
            }
        }
    }

    get y() {
        return this.sheet.rows.indexOf(this);
    };

    toJSON = () => {
        let cells = [];
        for (let i = 0; i < this.cells.length; ++i) {
            const cell = this.cells[i];
            const json = cell.toJSON();
            cells.push(json);
        }
        return {
            cells: cells
        }
    };

    getImages = (callback) => {
        function getImageOfOneCell(cell, params, callback) {
            cell.getImage(callback);
        }

        Util.handleMultiObjects(this.cells, getImageOfOneCell, null, callback);
    };

    addCell = () => {
        let cell = new ERPExcelCell({ row: this });
        this.cells.push(cell);
        return cell;
    };

    getCell = (x) => {
        return this.cells[x];
    };

    excelAddImage = () => {
        for (let i = 0; i < this.cells.length; ++i) {
            const cell = this.cells[i];
            cell.excelAddImage();
        }
    };

    excelSetImage = () => {
        for (let i = 0; i < this.cells.length; ++i) {
            const cell = this.cells[i];
            cell.excelSetImage();
        }
    };

    excelMergeCells = () => {
        for (let i = 0; i < this.cells.length; ++i) {
            const cell = this.cells[i];
            cell.excelMergeCells();
        }
    };

    excelSetCellStyle = () => {
        for (let i = 0; i < this.cells.length; ++i) {
            const cell = this.cells[i];
            cell.excelSetCellStyle();
        }
    };

    excelSetRowStyle = () => {
        this.excelRow.height = 30;
    };

    excelAddItem = () => {
        let texts = [];
        for (let i = 0; i < this.cells.length; ++i) {
            const cell = this.cells[i];
            texts.push(cell.getContentOfText());
        }

        this.excelRow = this.sheet.excelWorksheet.addRow(texts);

        return this.excelRow;
    };

}

export class ERPExcelSheet {

    workbook = null;
    name = '';
    rows = [];

    excelWorksheet = null;

    constructor(props) {
        this.workbook = props.workbook;
        this.name = props.name;
        this.rows = [];
        if(props.rows) {
            for(let i = 0; i < props.rows.length; i++) {
                let rowObj = Object.assign({}, props.rows[i], { sheet: this });
                let row = new ERPExcelRow(rowObj);
                this.rows.push(row);
            }
        }
    }

    toJSON = () => {
        let rows = [];
        for (let i = 0; i < this.rows.length; ++i) {
            const row = this.rows[i];
            const json = row.toJSON();
            rows.push(json);
        }
        return {
            rows: rows
        }
    };

    getImages = (callback) => {
        function getImageOfOneRow(row, params, callback) {
            row.getImages(callback);
        }

        Util.handleMultiObjects(this.rows, getImageOfOneRow, null, callback);
    };

    addRow = () => {
        let row = new ERPExcelRow({ sheet: this });
        this.rows.push(row);
        return row;
    };

    getRow = (y) => {
        return this.rows[y];
    };

    getCell = (x, y) => {
        if (x < 0 || y < 0 || y >= this.rows.length) {
            return null;
        }
        return this.rows[y].getCell(x);
    };

    excelAddImage = () => {
        for (let i = 0; i < this.rows.length; ++i) {
            const row = this.rows[i];
            row.excelAddImage();
        }
    };

    excelSetImage = () => {
        for (let i = 0; i < this.rows.length; ++i) {
            const row = this.rows[i];
            row.excelSetImage();
        }
    };

    excelMergeCells = () => {
        for (let i = 0; i < this.rows.length; ++i) {
            const row = this.rows[i];
            row.excelMergeCells();
        }
    };

    excelAddItem = () => {
        this.excelWorksheet = this.workbook.excelWorkbook.addWorksheet(this.name);
        for (let i = 0; i < this.rows.length; ++i) {
            const row = this.rows[i];
            row.excelAddItem();
        }
        return this.excelWorksheet;
    };

    excelSetCellStyle = () => {
        for (let i = 0; i < this.rows.length; ++i) {
            const row = this.rows[i];
            row.excelSetCellStyle();
        }
    };

    excelSetRowStyle = () => {
        for (let i = 0; i < this.rows.length; ++i) {
            const row = this.rows[i];
            row.excelSetRowStyle();
        }
    };

}

export class ERPExcelWorkbook {

    sheets = [];
    excelWorkbook = null;
    excelExcelWorkbookName = '标题';

    constructor(props) {
        if(props && props.sheets) {
            for(let i = 0; i < props.sheets.length; i++) {
                let sheetObj = Object.assign({}, props.sheets[i], { workbook: this});
                let sheet = new ERPExcelSheet(sheetObj);
                this.sheets.push(sheet);
            }
        }
    }

    toJSON = () => {
        let sheets = [];
        for (let i = 0; i < this.sheets.length; ++i) {
            const sheet = this.sheets[i];
            const json = sheet.toJSON();
            sheets.push(json);
        }
        return {
            sheets: sheets
        }
    };

    getImages = (callback) => {
        function getImageOfOneSheet(sheet, params, callback) {
            sheet.getImages(callback);
        }

        Util.handleMultiObjects(this.sheets, getImageOfOneSheet, null, callback);
    };

    addSheet = (name) => {
        let sheet = new ERPExcelSheet({ workbook: this, name: name });
        this.sheets.push(sheet);
        return sheet;
    };

    excelAddImage = () => {
        for (let i = 0; i < this.sheets.length; ++i) {
            const sheet = this.sheets[i];
            sheet.excelAddImage();
        }
    };

    excelSetImage = () => {
        for (let i = 0; i < this.sheets.length; ++i) {
            const sheet = this.sheets[i];
            sheet.excelSetImage();
        }
    };

    excelMergeCells = () => {
        for (let i = 0; i < this.sheets.length; ++i) {
            const sheet = this.sheets[i];
            sheet.excelMergeCells();
        }
    };

    excelAddItem = () => {
        for (let i = 0; i < this.sheets.length; ++i) {
            const sheet = this.sheets[i];
            sheet.excelAddItem();
        }
    };

    excelSetCellStyle = () => {
        for (let i = 0; i < this.sheets.length; ++i) {
            const sheet = this.sheets[i];
            sheet.excelSetCellStyle();
        }
    };

    excelSetRowStyle = () => {
        for (let i = 0; i < this.sheets.length; ++i) {
            const sheet = this.sheets[i];
            sheet.excelSetRowStyle();
        }
    };

    excelExport = (callback) => {
        this.getImages(data => {
            debugger
            this.excelWorkbook = new Excel.Workbook();
            this.excelAddItem();
            this.excelAddImage();

            this.excelSetCellStyle();
            this.excelSetRowStyle();
            // this.excelSetCellStyle();

            this.excelSetImage();
            this.excelMergeCells();

            this.excelWorkbook.xlsx.writeBuffer()
                .then(buffer => {
                    // this.progressGaga(false)
                    let base64 = buffer.toString('base64')
                    // 使用atob方法解码base64
                    let raw = window.atob(base64);
                    // 创建一个存储解码后数据的数组
                    let uInt8Array = new Uint8Array(raw.length);
                    // blob只能接收二进制编码，需要讲base64转为二进制再塞进去
                    for (let i = 0; i < raw.length; ++i) {
                        uInt8Array[i] = raw.charCodeAt(i);
                    }
                    // 这里给了一个返回值，在别的方法掉用传入base64编码就可以得到转化后的blob
                    const link = document.createElement('a');
                    const blob = new Blob([uInt8Array], { type: 'application/vnd.ms-excel' });
                    link.style.display = 'none';
                    link.href = URL.createObjectURL(blob);
                    //设置下载的Excel表名
                    link.setAttribute('download', this.excelExcelWorkbookName + '.xlsx');
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    if(callback){
                        callback(true);
                    }
                }).catch(error => {
                    // this.progressGaga(false)
                    throw error;
                });
        });
    };
}


