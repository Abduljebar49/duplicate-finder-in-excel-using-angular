import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import * as fileSaver from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'analyze-excel-in-angular';
  excelData: any[] = [];
  convertedData: any[] = [];
  excelHeader = []
  fileName: string = "";
  p: number = 1;
  isFileUploading: boolean = false;

  getValueFromObject(object: any, value: string) {
    var index = "";
    if (value) {
      try {
        index = value.toLocaleLowerCase();
      } catch (e) {
        return 0
      }
    }
    else
      return 0;
    return object[index];
  }

  getFileName(file: any) {
    var name = file.name.toString().split('.');
    this.fileName = name[0];
  }

  onFileChange($event: any) {
    this.excelData = [];
    this.getFileName($event.target.files[0]);
    this.excelHeader = [];
    this.excelData = [];
    const target: DataTransfer = <DataTransfer>$event.target;
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { raw: false, type: 'binary' });
      const wsname: string = wb.SheetNames[0];

      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
      this.excelData = XLSX.utils.sheet_to_json(ws, { header: 1 });
      this.excelHeader = this.excelData[0];
      this.excelData = this.excelData.slice(1);
      console.log(this.excelHeader);
      console.log(this.excelData);
    };
    reader.readAsBinaryString(target.files[0]);
  }

  checkForDuplication() {
    const data = this.findDuplicates(this.excelData);
    console.log("data : ", data);
  }

  uploadStockStatus() {
    this.isFileUploading = true;
  }

  changeGivenDataToJSON(data: any[]) {
    const result = data.reduce((acc, cur) => {
      const values = Object.values(cur);
      acc.push(this.excelHeader.reduce((obj: any, header, i) => {
        obj[header] = values[i];
        return obj;
      }, {}));
      return acc;
    }, []);
    return result;// this.stockStatusJson = result;
  }

  exportToExacel() {
    const result = this.changeGivenDataToJSON(this.convertedData)
    this.exportAsExcelFile(result, 'converted data');
  }

  public exportAsExcelFile(json: any[], excelFileName: string): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ['data'],
    };
    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });
    this.saveAsExcelFile(excelBuffer, excelFileName);
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], { type: EXCEL_TYPE });
    fileSaver.saveAs(
      data,
      fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION
    );
  }

  findDuplicates(dataArray: any) {
    const uniqueArray: any[] = [];
    const duplicatesArray = [];
    for (let i = 0; i < dataArray.length; i++) {
      const obj = dataArray[i];
      const isExist = dataArray.filter((ele: any) => JSON.stringify(ele) === JSON.stringify(obj)).length;
      console.log("isExist  : ", isExist);
      if (isExist == 1) {
        uniqueArray.push(obj);
      } else {
        duplicatesArray.push(obj);
      }
    }

    const data = {
      unique: uniqueArray,
      duplicate: duplicatesArray
    }
    return data;
  }

}
