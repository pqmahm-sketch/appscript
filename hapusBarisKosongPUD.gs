/**
 * Menghapus baris-baris kosong di sheet "PUD" pada spreadsheet Database FU PUD.
 * Baris dianggap kosong jika semua sel di baris tersebut kosong atau hanya berisi spasi.
 */
function hapusBarisKosongPUD() {
  var spreadsheetId = '1bC7tJPVoCz6x_lqq1kHbFEO15FRoyjdV6iQ9zMzAFrA';
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName('PUD');

  if (!sheet) {
    Logger.log('Sheet "PUD" tidak ditemukan.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    Logger.log('Tidak ada data selain header.');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    var isEmpty = data[i].every(function(cell) {
      return cell === '' || (typeof cell === 'string' && cell.trim() === '');
    });
    if (isEmpty) {
      rowsToDelete.push(i + 2);
    }
  }

  if (rowsToDelete.length === 0) {
    Logger.log('Tidak ada baris kosong yang ditemukan.');
    return;
  }

  Logger.log('Ditemukan ' + rowsToDelete.length + ' baris kosong. Menghapus...');

  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  Logger.log('Selesai. ' + rowsToDelete.length + ' baris kosong berhasil dihapus.');
}
