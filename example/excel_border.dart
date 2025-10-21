import 'dart:io';

import 'package:excel/excel.dart';

void main(List<String> args) async {
  final excel = Excel.createExcel();
  final sheet = excel[excel.getDefaultSheet()!];

  sheet.merge(
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
    CellIndex.indexByColumnRow(columnIndex: 10, rowIndex: 5),
  );

  sheet.merge(
    CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 10),
    CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 10),
  );

  final border = Border(borderColorHex: '#FF000000'.excelColor, borderStyle: BorderStyle.Thin);

  sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1), TextCellValue('Merged cell border'));

  sheet.setMergedCellStyle(
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
    CellStyle(
      fontSize: 25,
      topBorder: border,
      bottomBorder: border,
      leftBorder: border,
      rightBorder: border,
      diagonalBorder: border,
      diagonalBorderDown: true,
    ),
  );

  sheet.setMergedCellStyle(
    CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 10),
    CellStyle(
      topBorder: border,
      bottomBorder: border,
      leftBorder: border,
      rightBorder: border,
      diagonalBorder: border,
      diagonalBorderDown: true,
      diagonalBorderUp: true,
    ),
  );

  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
    TextCellValue('Normal border'),
    cellStyle: CellStyle(
      fontSize: 25,
      topBorder: border,
      bottomBorder: border,
      leftBorder: border,
      rightBorder: border,
      diagonalBorder: border,
    ),
  );

  sheet.setColumnWidth(0, 50);

  // Create the example excel file in the current directory
  final outputFile = 'excel_custom.xlsx';

  final fileBytes = await excel.save();
  if (fileBytes != null) {
    File(outputFile)
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}
