import ExcelJS from 'exceljs';

export interface ITitleItem {
    content: string;
    bold: boolean;
    width: number;
    fontSize?: number;
}

export class ExcelManager {
    static maxWidth = 30;

    constructor() { }

    public static async load(
        column: Array<any>,
        data: Array<any>,
        titleList: Array<ITitleItem>
    ) {
        const workbook: any = new ExcelJS.Workbook();
        const worksheet: any = workbook.addWorksheet('Sheet');
        const header: Array<any> = [];
        const convertColumn: any = [];
        const subHeader: Array<string> = [];
        const mergeCell: Array<any> = [];
        const visibleHeader: Array<string> = [];

        column.forEach((item) => {
            if (item.children) {
                mergeCell.push({ start: subHeader.length + 1, stop: subHeader.length });
                visibleHeader.push(worksheet.getCell(titleList.length + 1, subHeader.length + 1).address);
                item.children.forEach((element: any) => {
                    convertColumn.push({ key: element.key });
                    subHeader.push(element.header);
                    mergeCell[mergeCell.length - 1].stop += 1;
                });
            } else {
                visibleHeader.push(worksheet.getCell(titleList.length + 1, subHeader.length + 1).address);
                convertColumn.push({ key: item.key });
                subHeader.push('');
            }
            header.push(item.header);
        });

        const mergerColumn: Array<any> = [];
        const subVisible: Array<string> = [];
        subHeader.forEach((_it, index) => {
            if (subHeader[index] === '') {
                mergerColumn.push({
                    start: worksheet.getCell(titleList.length + 1, index + 1).address,
                    stop: worksheet.getCell(titleList.length + 2, index + 1).address,
                });
            } else {
                subVisible.push(worksheet.getCell(titleList.length + 2, index + 1).address);
            }
        });

        if (titleList.length > 0) {
            titleList.forEach((item: ITitleItem, index: number) => {
                const first: string = worksheet.getCell(index + 1, 1).address;
                const second: string = worksheet.getCell(index + 1, Math.min(item.width, subHeader.length)).address;
                worksheet.mergeCells(`${first}:${second}`);
                const locate: any = worksheet.getCell(first);
                locate.value = item.content;
                locate.font = { bold: item.bold, size: item.fontSize ? item.fontSize : 12 };
                locate.alignment = {
                    horizontal: 'center',
                    vertical: 'middle',
                };
            });
        }

        mergeCell.forEach((item) => {
            worksheet.mergeCells(
                `${worksheet.getCell(titleList.length + 1, item.start).address}:${worksheet.getCell(titleList.length + 1, item.stop).address
                }`
            );
        });

        mergerColumn.forEach((item) => {
            worksheet.mergeCells(`${item.start}:${item.stop}`);
        });

        visibleHeader.forEach((item, index) => {
            const located: any = worksheet.getCell(item);
            located.value = header[index];
            located.alignment = { vertical: 'middle', horizontal: 'center' };
            located.font = { bold: true };
        });

        const newSub = subHeader.filter((it) => it !== '');
        subVisible.forEach((item, index) => {
            const located: any = worksheet.getCell(item);
            located.value = newSub[index];
            located.alignment = { vertical: 'middle', horizontal: 'center' };
            located.font = { bold: true };
        });

        worksheet.columns = convertColumn;
        data.forEach((item) => worksheet.addRow(item));

        worksheet.columns.forEach((col: any) => {
            let maxLength = 0;
            col.eachCell({ includeEmpty: true }, (cell: any) => {
                const columnLength = cell.value ? cell.value.toString().length : 10;
                if (columnLength > maxLength) {
                    maxLength = columnLength;
                }
            });
            col.width = Math.min(maxLength + 2, this.maxWidth);
        });

        const buffer = await workbook.xlsx.writeBuffer();
        return buffer;
    };

    public static downloadBuffer(buffer: ArrayBuffer, fileName: string) {
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        const urlObj = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.download = fileName;
        a.href = urlObj;
        a.click();
        a.remove();
        window.URL.revokeObjectURL(urlObj);
    }
}