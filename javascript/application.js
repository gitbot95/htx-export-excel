function RunExcelJSExport() {
  var dataField = document.getElementById('data-field');
  var workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');
  var jsonData = JSON.parse(dataField.value);
  for (var key in jsonData) {
    worksheet.addRow([
      parseInt(key) + 1,
      jsonData[key]['name'],
      jsonData[key]['email'],
      jsonData[key]['address'],
      jsonData[key]['description'].replace(/(<([^>]+)>)/gi, '').replace(/\&nbsp;/g, ''),
      [jsonData[key]['backup_name'], jsonData[key]['backup_phone'], jsonData[key]['backup_description']].join(' - '),
      jsonData[key]['backup_address'],
      jsonData[key]['product_name'],
      [jsonData[key]['category'], jsonData[key]['category_type']].join(' / '),
      jsonData[key]['avg'],
      jsonData[key]['price_departure'],
      jsonData[key]['price_destination'],
    ]);
  }
  debugger;
  var stamp = Date.now();
  workbook.xlsx.writeBuffer().then(function (buffer) {
    saveAs(new Blob([buffer], { type: 'application/octet-stream' }), `htx_${stamp}.xlsx`);
  });
}
