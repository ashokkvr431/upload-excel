import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';


@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './app.component.html',
})
export class AppComponent {
  excelData: any[][] = [];
  displayedRows: any[][] = [];

  onFileChange(event: any): void {
    const target: DataTransfer = <DataTransfer>(event.target);

    const reader: FileReader = new FileReader();

    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      this.excelData = XLSX.utils.sheet_to_json(ws, { header: 1 });

      this.displayedRows = [];
      let i = 0;
      const interval = setInterval(() => {
        if (i < this.excelData.length) {
          this.displayedRows.push(this.excelData[i]);
          i++;
        } else {
          clearInterval(interval);
        }
      }, 0); // Adjust speed here
    };

    reader.readAsBinaryString(target.files[0]);
  }

  exportToExcel(): void {
    if (!this.displayedRows || this.displayedRows.length === 0) return;

    const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.displayedRows);
    const workbook: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });

    const data: Blob = new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
    });

    FileSaver.saveAs(data, 'exported-data.xlsx');
  }

  exportToPDF(): void {
  if (!this.displayedRows || this.displayedRows.length === 0) return;

  const doc = new jsPDF();

  doc.setFontSize(16);
  doc.text('Exported Excel Data', 14, 15);

  const headers = this.displayedRows[0];
  const dataRows = this.displayedRows.slice(1);


  autoTable(doc, {
    startY: 25,
    head: [headers],
    body: dataRows,
    theme: 'grid',             
    styles: {
      fontSize: 10,
      cellPadding: 4,
      halign: 'center'
    },
    headStyles: {
      fillColor: [41, 128, 185],
      textColor: 255,
      fontStyle: 'bold'
    },
    alternateRowStyles: {
      fillColor: [245, 245, 245]
    },
    margin: { top: 20 }
  });

  doc.save('exported-data.pdf');
}

}
