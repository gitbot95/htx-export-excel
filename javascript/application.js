$(function () {
  var pageTitle = window.location.href.match(/#(\w*)/) ? window.location.href.match(/#(\w*)/)[1] : 'coopSale';
  var downloadFields = {
    coopSale: [
      ['name'],
      ['email', 'phone'],
      ['address'],
      ['description'],
      ['backup_name', 'backup_phone', 'backup_description'],
      ['backup_address'],
      ['product_name'],
      ['category', 'category_type'],
      ['avg'],
      ['price_departure'],
      ['price_destination']
    ],
    coopBuy: [['name'], ['email', 'phone'], ['description'], ['address'], ['product_name'], ['category', 'category_type']],
    lmhtxSale: [['nguoiBanHangTenDonVi', 'nguoiBanHang'], ['nguoiBanHangDienThoai'], ['tenHang'], ['giaBanHTX'], ['giaBanGD']],
    lmhtxBuy: [['nguoiMuaHang', 'nguoiMuaHangTenDonVi'], ['nguoiMuaHangDienThoai'], ['tenDonVi'], ['tenHang']]
  };

  $('.dropdown-menu .dropdown-item').on('click', function () {
    pageTitle = this.href.match(/#(\w*)/)[1];
    swicthPage(pageTitle);
  });

  $('#downloadFile').on('click', function () {
    RunExcelJSExport(downloadFields[pageTitle]);
  });
});

function swicthPage(page) {
  var titleHead = {
    coopSale: 'CoopLink Cần bán',
    coopBuy: 'CoopLink Cần mua',
    lmhtxSale: 'Liên minh Cần bán',
    lmhtxBuy: 'Liên minh Cần mua'
  };
  $('.guide-zone').empty().append(`<h1 class="d-flex justify-content-center">${titleHead[page]}</h1>`);
}

function RunExcelJSExport(fields) {
  var workbook = new ExcelJS.Workbook();
  var worksheet = workbook.addWorksheet('Sheet 1');
  var jsonData = JSON.parse($('#data-field').val());
  for (var key in jsonData) {
    worksheet.addRow(fillData(jsonData, key, fields));
  }
  var stamp = Date.now();
  workbook.xlsx.writeBuffer().then(function (buffer) {
    saveAs(new Blob([buffer], { type: 'application/octet-stream' }), `htx_${stamp}.xlsx`);
  });
}

function fillData(jsonData, key, fields) {
  arr = [parseInt(key) + 1];
  for (var str of fields) {
    arr.push(
      str
        .map((text) =>
          jsonData[key][text]
            ?.toString()
            .replace(/(<([^>]+)>)/gi, '')
            .replace(/\&nbsp;/g, '')
        )
        .join(' \n')
    );
  }
  return arr;
}
