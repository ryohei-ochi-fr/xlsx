import { Component } from '@angular/core';
import { read, utils, writeFileXLSX } from 'xlsx-js-style';

interface President { Name: string; Index: number };

@Component({
  selector: 'app-root',
  template: `
<div class="content" role="main"><table>
  <thead><th>Name</th><th>Index</th></thead>
  <tbody>
    <tr *ngFor="let row of rows">
      <td>{{row.Name}}</td>
      <td>{{row.Index}}</td>
    </tr>
  </tbody><tfoot>
    <button (click)="onSave()">Export XLSX</button>
  </tfoot>
</table></div>
`
})
export class AppComponent {

  rows: President[] = [{ Name: "SheetJS", Index: 0 }];

  ngOnInit(): void {
    (async () => {
      /* Download from https://sheetjs.com/pres.numbers */
      // const f = await fetch("https://sheetjs.com/pres.numbers");
      const f = await fetch('assets/sample_001.xlsx');
      const ab = await f.arrayBuffer();

      /* parse workbook */
      const wb = read(ab);

      /* update data */
      this.rows = utils.sheet_to_json<President>(wb.Sheets[wb.SheetNames[0]]);

      console.log(
        utils.sheet_to_json<string>(wb.Sheets[wb.SheetNames[0]])
      );

    })();
  }
  /* get state data and export to XLSX */
  onSave(): void {
    const ws = utils.json_to_sheet(this.rows);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "Data");
    writeFileXLSX(wb, "SheetJSAngularAoO.xlsx");
  }
}
