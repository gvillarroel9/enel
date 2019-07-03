import { Component, OnInit } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import * as XLSX from 'xlsx';

class Cliente {
  constructor(
    public id: String,
    public montoFactura: String,
    public direccion: String,
    public fechaVencimiento: String)
    {}
}

@Component({
  selector: 'app-plantilla',
  templateUrl: './plantilla.component.html',
  styleUrls: ['./plantilla.component.scss']
})
export class PlantillaComponent implements OnInit {
  data: any;
  fileName: string;
  complete = true;
  url = 'https://metric-marks-80160.herokuapp.com/api/factura/';
  clienteActual: Cliente = new Cliente('', '' , '' , '');
  cantidadClientes: number = 0;

  constructor( private http: HttpClient) {
   }

  ngOnInit() {
    this.fileName = 'report.xlsx';
  }

  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = (evt.target) as DataTransfer;
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});
      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = XLSX.utils.sheet_to_json(ws, {header: 1});
      //console.log(this.data[1][3]);
      for( let i = 1; i <this.data.length; i++){    
        this.data[i][4] = null;
        this.data[i][5] = null;
      }
    };
    reader.readAsBinaryString(target.files[0]);
  }

  consultaMasiva() {
    this.complete = false;
    console.log(this.data)
    console.log(this.data[3].length);
    for( let i = 1; i <this.data.length; i++){
      if(this.data[i][3]&&!this.data[i][4]){
        this.consultar(this.data[i][3],i);
        this.cantidadClientes++;
      }
    }
    if(!this.cantidadClientes){
      this.complete = true;
    }

  }

  consultar(idClient: string, indice:number) {
    this.http.get(this.url + idClient).subscribe(
      (res) => {
        console.log(idClient);
        if(res.supplyBalance == '0'){
          this.data[indice][4] = 'SIN DEUDA';
        }else {
          this.data[indice][4] = res.supplyBalance;
        }        
        this.data[indice][5] = res.expiryDate;
        this.data[indice][0] = res.address;
        if(res.beResultCode=="1"){
          this.data[indice][4] = "Error en Consulta";
          console.log(res);
        }
        if(indice==this.cantidadClientes){
          this.complete = true;
          this.cantidadClientes=0;
        }      
      },
      (err) => {
        console.log(err);
      }
    );
  }


  export(): void {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);
    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }

}
