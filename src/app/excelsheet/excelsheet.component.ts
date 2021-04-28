import { Component, OnInit } from '@angular/core';
import { LocalDataSource } from 'ng2-smart-table'
import Excel from 'exceljs';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-excelsheet',
  templateUrl: './excelsheet.component.html',
  styleUrls: ['./excelsheet.component.css']
})
export class ExcelsheetComponent implements OnInit {

  dataSource: [][];

  /**
   *
   * "": "Item"
Bill: "Code"
External: "Link"
Net: "Rate"
Net_1: "Amount"
Price: "Code"
_1: "Unit"
_2: "Trade"
__EMPTY: "Level"
__EMPTY_1: "Bill description"
__EMPTY_2: "Bill quantity"
__rowNum__: 1
   */
  settings = {
    columns: {
      External: {
        title: 'Le',
      },
      __EMPTY: {
        title: 'Level',
      },
      Price: {
          title: 'Preço'
      },
      Bill: {
        title: 'Bill'
      },

      __EMPTY_1: {
        title: 'Tipo',
      },
      Resource: {
        title: 'Recurso',
      },
      "": {
        title: 'ffdf'
      },
      _1: {
        title: 'Uni'
      }
    },
  };

  data = []

  /**data = [
    {
      id: 1,
      name: 'Leanne Graham',
      username: 'Bret',
      email: 'Sincere@april.biz',
    },
    {
      id: 2,
      name: 'Ervin Howell',
      username: 'Antonette',
      email: 'Shanna@melissa.tv',
    },
    {
      id: 3,
      name: 'Clementine Bauch',
      username: 'Samantha',
      email: 'Nathan@yesenia.net',
    },
    {
      id: 4,
      name: 'Patricia Lebsack',
      username: 'Karianne',
      email: 'Julianne.OConner@kory.org',
    },
    {
      id: 5,
      name: 'Chelsey Dietrich',
      username: 'Kamren',
      email: 'Lucio_Hettinger@annie.ca',
    },
    {
      id: 6,
      name: 'Mrs. Dennis Schulist',
      username: 'Leopoldo_Corkery',
      email: 'Karley_Dach@jasper.info',
    },
    {
      id: 7,
      name: 'Kurtis Weissnat',
      username: 'Elwyn.Skiles',
      email: 'Telly.Hoeger@billy.biz',
    },
    {
      id: 8,
      name: 'Nicholas Runolfsdottir V',
      username: 'Maxime_Nienow',
      email: 'Sherwood@rosamond.me',
    },
    {
      id: 9,
      name: 'Glenna Reichert',
      username: 'Delphine',
      email: 'Chaim_McDermott@dana.io',
    },
    {
      id: 10,
      name: 'Clementina DuBuque',
      username: 'Moriah.Stanton',
      email: 'Rey.Padberg@karina.biz',
    },
    {
      id: 11,
      name: 'Nicholas DuBuque',
      username: 'Nicholas.Stanton',
      email: 'Rey.Padberg@rosamond.biz',
    },
  ];
 */
  constructor() { }

  ngOnInit(): void {
  }

  onFileUpdate(event){
    if(event.target.files.length !== 1) throw new Error('Não podes carregar multiplos ficheiros!')
    const file = event.target.files[0];
    const wb = new Excel.Workbook();
    const reader = new FileReader();

    reader.readAsArrayBuffer(file);

    reader.onload = () => {
      const buffer = reader.result;
      wb.xlsx.load(buffer).then(workbook => {
        const worksheetName = workbook.worksheets[0].name;
        const worksheet = workbook.getWorksheet(worksheetName);

        const cellA1 = worksheet.getCell('A1').value
        const cellA2 = worksheet.getCell('A2').value
        const columA = worksheet.getColumn('A');
        columA.header =`${cellA1} / ${cellA2}`;

        const cellB1 = worksheet.getCell('B1').value
        const cellB2 = worksheet.getCell('B2').value
        const columB = worksheet.getColumn('B');
        columB.header =`${cellB1} / ${cellB2}`;

        const cellC1 = worksheet.getCell('C1').value
        const cellC2 = worksheet.getCell('C2').value
        const columC = worksheet.getColumn('C');
        columC.header =`${cellC1} / ${cellC2}`;

        const cellD1 = worksheet.getCell('D1').value
        const cellD2 = worksheet.getCell('D2').value
        const columD = worksheet.getColumn('D');
        columD.header =`${cellD1} / ${cellD2}`;

        const cellE1 = worksheet.getCell('E1').value
        const cellE2 = worksheet.getCell('E2').value
        const columE = worksheet.getColumn('E');
        columE.header =`${cellE1} / ${cellE2}`;

        const cellF1 = worksheet.getCell('F1').value
        const cellF2 = worksheet.getCell('F2').value
        const columF = worksheet.getColumn('F');
        columF.header =`${cellF1} / ${cellF2}`;

        const cellG1 = worksheet.getCell('G1').value
        const cellG2 = worksheet.getCell('G2').value
        const columG = worksheet.getColumn('G');
        columG.header =`${cellG1} / ${cellG2}`;

        const cellH1 = worksheet.getCell('H1').value
        const cellH2 = worksheet.getCell('H2').value
        const columH = worksheet.getColumn('H');
        columH.header =`${cellH1} / ${cellH2}`;

        const cellI1 = worksheet.getCell('I1').value
        const cellI2 = worksheet.getCell('I2').value
        const columI = worksheet.getColumn('I');
        columI.header =`${cellI1} / ${cellI2}`;

        const cellJ1 = worksheet.getCell('J1').value
        const cellJ2 = worksheet.getCell('J2').value
        const columJ = worksheet.getColumn('J');
        columJ.header =`${cellJ1} / ${cellJ2}`;

        const cellK1 = worksheet.getCell('K1').value
        const cellK2 = worksheet.getCell('K2').value
        const columK = worksheet.getColumn('K');
        columK.header =`${cellK1} / ${cellK2}`;




        //const cellA = worksheet.getColumn('A');
       // cellA.header = 'External / Link'


       // console.log(cellA);
        /**
         * worksheet.eachRow((row, rowNumber) => {
          console.log('Row ' + rowNumber + row.values)
        })
         */
        //console.log(worksheet.columnCount , ' Valor da celular')
        //console.log(worksheet.rowCount, ' numeros de colunas')
      })
    }
  }

  onFileChange(evt: any) {
    const target : DataTransfer =  <DataTransfer>(evt.target);


    if (target.files.length !== 1) throw new Error('Cannot use multiple files');

    const reader: FileReader = new FileReader();

    reader.onload = (e: any) => {
      const bstr: string = e.target.result;

      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      const wsname : string = wb.SheetNames[0];


      wb.SheetNames.forEach((name) => {
        console.log(`Nome aqui ${name}`)
      })

      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      //console.log(ws);

      this.data = (XLSX.utils.sheet_to_json(ws, { header: 4 }));

      console.log(wb)

     // let x = this.dataSource.slice(1);
      //console.log(x);

    };

   reader.readAsBinaryString(target.files[0]);

  }

}
