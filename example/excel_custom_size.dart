import 'dart:convert';
import 'dart:io';
import 'dart:math';

import 'package:excel/excel.dart';

void main(List<String> args) async {
  final excel = Excel.createExcel();
  final sheet = excel[excel.getDefaultSheet()!];

  for (var row = 0; row < 100; row++) {
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: row)).value = TextCellValue(getRandString());

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: row)).value = TextCellValue(getRandString());

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: row)).value = TextCellValue(getRandString());

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: row)).value = TextCellValue(getRandString());

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: row)).value = TextCellValue(getRandString());

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 50, rowIndex: row)).value = TextCellValue(getRandString());
  }

  sheet.setDefaultColumnWidth();
  sheet.setDefaultRowHeight();

  sheet.setColumnAutoFit(0);
  sheet.setColumnAutoFit(1);
  sheet.setColumnAutoFit(2);

  sheet.setColumnWidth(0, 10);
  sheet.setColumnWidth(1, 10);
  sheet.setColumnWidth(50, 10);

  sheet.setRowHeight(1, 100);

  sheet.merge(
    CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 10),
  );

  // Create the example excel file in the current directory
  final outputFile = 'excel_custom.xlsx';

  final fileBytes = await excel.save();
  if (fileBytes != null) {
    File(outputFile)
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}

String getRandString() {
  final random = Random.secure();
  final len = random.nextInt(20);
  final values = List<int>.generate(len, (i) => random.nextInt(255));
  return base64UrlEncode(values);
}
