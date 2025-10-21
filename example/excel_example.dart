import 'dart:io';

import 'package:excel/excel.dart';
import 'package:path/path.dart';

void main(List<String> args) async {
  //var file = "/Users/kawal/Desktop/excel/test/test_resources/example.xlsx";
  //var bytes = File(file).readAsBytesSync();
  final excel = Excel.createExcel();
  // or
  //var excel = Excel.decodeBytes(bytes);

  ///
  ///
  /// reading excel file values
  ///
  ///
  for (final table in excel.tables.keys) {
    print(table);
    print(excel.tables[table]!.maxColumns);
    print(excel.tables[table]!.maxRows);
    for (final row in excel.tables[table]!.rows) {
      print('${row.map((e) => e?.value)}');
    }
  }

  ///
  /// Change sheet from rtl to ltr and vice-versa i.e. (right-to-left -> left-to-right and vice-versa)
  ///
  final sheet1rtl = excel['Sheet1'].isRTL;
  excel['Sheet1'].isRTL = false;
  print('Sheet1: ((previous) isRTL: $sheet1rtl) ---> ((current) isRTL: ${excel['Sheet1'].isRTL})');

  final sheet2rtl = excel['Sheet2'].isRTL;
  excel['Sheet2'].isRTL = true;
  print('Sheet2: ((previous) isRTL: $sheet2rtl) ---> ((current) isRTL: ${excel['Sheet2'].isRTL})');

  ///
  ///
  /// declaring a cellStyle object
  ///
  ///
  final cellStyle = CellStyle(
    bold: true,
    italic: true,
    textWrapping: TextWrapping.WrapText,
    fontFamily: getFontFamily(FontFamily.Comic_Sans_MS),
  );

  var sheet = excel['mySheet'];

  final cell = sheet.cell(CellIndex.indexByString('A1'));
  cell.value = TextCellValue('Heya How are you I am fine ok goood night');
  cell.cellStyle = cellStyle;

  final cell2 = sheet.cell(CellIndex.indexByString('E5'));
  cell2.value = TextCellValue('Heya How night');
  cell2.cellStyle = cellStyle;

  /// printing cell-type
  print(
    'CellType: ${switch (cell.value) {
      null => 'empty',
      TextCellValue() => 'text',
      FormulaCellValue() => 'Formula',
      IntCellValue() => 'int',
      DoubleCellValue() => 'double',
      DateCellValue() => 'date',
      DateTimeCellValue() => 'date+time',
      TimeCellValue() => 'time',
      BoolCellValue() => 'bool',
    }}',
  );

  ///
  ///
  /// Iterating and changing values to desired type
  ///
  ///
  for (var row = 0; row < sheet.maxRows; row++) {
    sheet.row(row).forEach((Data? cell1) {
      if (cell1 != null) {
        cell1.value = TextCellValue(' My custom Value ');
      }
    });
  }

  excel.rename('mySheet', 'myRenamedNewSheet');

  final sheet1 = excel['Sheet1'];
  sheet1.cell(CellIndex.indexByString('A1')).value = TextCellValue('Sheet1');

  /// fromSheet should exist in order to sucessfully copy the contents
  excel.copy('Sheet1', 'newlyCopied');

  final sheet2 = excel['newlyCopied'];
  sheet2.cell(CellIndex.indexByString('A1')).value = TextCellValue('Newly Copied Sheet');

  /// renaming the sheet
  excel.rename('oldSheetName', 'newSheetName');

  /// deleting the sheet
  excel.delete('Sheet1');

  /// unlinking the sheet if any link function is used !!
  excel.unLink('sheet1');

  sheet = excel['sheet'];

  /// appending rows and checking the time complexity of it
  final stopwatch = Stopwatch()..start();
  final list = List<List<TextCellValue>>.generate(
    9000,
    (index) => List.generate(20, (index1) => TextCellValue('$index $index1')),
  );

  print('list creation executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  for (var i = 0; i < list.length; i++) {
    final row = list[i];
    sheet.appendRow(row);
  }
  print('appending executed in ${stopwatch.elapsed}');

  sheet.appendRow([
    const IntCellValue(8),
    const DoubleCellValue(999.62221),
    DateCellValue(year: DateTime.now().year, month: DateTime.now().month, day: DateTime.now().day),
    DateTimeCellValue.fromDateTime(DateTime.now()),
  ]);

  final isSet = excel.setDefaultSheet(sheet.sheetName);
  // isSet is bool which tells that whether the setting of default sheet is successful or not.
  if (isSet) {
    print('${sheet.sheetName} is set to default sheet.');
  } else {
    print('Unable to set ${sheet.sheetName} to default sheet.');
  }

  final columnIterableSheet = excel['ColumnIterables'];

  final columnIterables = ['A', 'B', 'C', 'D', 'E'];
  final columnIndex = 0;

  for (final columnValue in columnIterables) {
    columnIterableSheet
        .cell(CellIndex.indexByColumnRow(rowIndex: columnIterableSheet.maxRows, columnIndex: columnIndex))
        .value = TextCellValue(
      columnValue,
    );
  }

  // Saving the file

  final outputFile = '/Users/kawal/Desktop/git_projects/r.xlsx';

  //stopwatch.reset();
  final fileBytes = await excel.save();
  //print('saving executed in ${stopwatch.elapsed}');
  if (fileBytes != null) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}
