// Import stylesheets
import './style.css';

const generateExcelBtn = document.querySelector('#generateExcelBtn');

generateExcelBtn.addEventListener('click', (event) => {
  import('exceljs').then((Excel) => {
    console.log(Excel);
    const data = [
      {
        dcNbr: '7471',
        countryCode: 'MX',
        docName: '21165679_1451066723_RVprintpod070130665.pdf',
        date: '2022-04-18T12:15:10.067+0000',
        appointmentNbr: 21165679,
        poNbr: '1451066723',
        scheduledDate: '2022-04-25',
        vendorNbr: 757822020,
        vendorName: 'PROCTER AND GAMB S DE RL DE CV',
        receiverNbr: 962851,
        itemNbr: 250210,
        itemDesc: 'PANTENE SH BRILLO',
        orderVnpk: 14, //
        orderWhpk: 14, //
        receivedVnpk: 14,
        receivedWhpk: 14,
        overageVnpk: 0,
        overageWhpk: 0,
        shortVnpk: 0,
        shortWhpk: 0,
        damagedVnpk: 0,
        damagedWhpk: 14,
        damagedReason: '',
        rejectedVnpk: 0,
        rejectedWhpk: 0,
        rejectedReason: '',
      },
    ];
    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('My Sheet');

    const header = [
      'dcNbr',
      'countryCode',
      'docName',
      'date',
      'appointmentNbr',
      'poNbr',
      'scheduledDate',
      'vendorNbr',
      'vendorName',
      'receiverNbr',
      'itemNbr',
      'itemDesc',
      'Order QTY',
      '',
      'Received Qty',
      '',
      'Overage Qty',
      '',
      'Shortage',
      '',
      'Damage',
      '',
      'damagedReason',
      'Reject',
      '',
    ];
    sheet.addRow(header);
    const subHeader = [];
    subHeader[13] = 'Vnpk';
    subHeader[14] = 'Whpk';

    subHeader[15] = 'Vnpk';
    subHeader[16] = 'Whpk';

    subHeader[17] = 'Vnpk';
    subHeader[18] = 'Whpk';

    subHeader[19] = 'Vnpk';
    subHeader[20] = 'Whpk';

    subHeader[21] = 'Vnpk';
    subHeader[22] = 'Whpk';

    subHeader[24] = 'Vnpk';
    subHeader[25] = 'Whpk';

    sheet.addRow(subHeader);
    sheet.mergeCells('A1:A2');

    sheet.mergeCells('M1:N1');
    sheet.mergeCells('O1:P1');
    sheet.mergeCells('Q1:R1');
    sheet.mergeCells('S1:T1');
    sheet.mergeCells('U1:V1');
    sheet.mergeCells('X1:Y1');
    const x = data
      .map((d) => Object.values(d))
      .forEach((d, index) => {
        sheet.addRow(d);
      });

    //sheet.addRow(['2020', 'March', 'Abc', 'xyz', 'Y', '']);
    //sheet.addRow(['2020', 'March', 'Abc', 'xyz', '', '', 'Y']);

    import('file-saver').then((fs) => {
      console.log(fs);
      workbook.xlsx.writeBuffer().then((data) => {
        let blob = new Blob([data], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        fs.saveAs(blob, 'Data.xlsx');
      });
    });
  });
});
