import 'dart:io';

import 'package:excel/excel.dart';
import 'package:path/path.dart';

void main(List<String> args) async {
  final stopwatch = Stopwatch()..start();

  final excel = Excel.createExcel();
  final sh = excel['Sheet1'];
  for (var i = 0; i < 8; i++) {
    sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).value = TextCellValue('Column $i');
    //sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).cellStyle =CellStyle(bold: true);
  }
  for (var row = 1; row < 1000; row++) {
    for (var column = 0; column < 1000; column++) {
      sh.cell(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: column)).value = TextCellValue(
        '$row$column value',
      );
    }
  }
  print('Generating executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  final fileBytes = await excel.encode();

  print('Encoding executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  if (fileBytes != null) {
    File(join('/Users/kawal/Desktop/r2.xlsx'))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
  print('Downloaded executed in ${stopwatch.elapsed}');
}
