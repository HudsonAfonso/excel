import 'dart:io';

import 'package:excel/excel.dart';
import 'package:path/path.dart';

void main(List<String> args) async {
  //var file = "/Users/kawal/Desktop/excel/test/test_resources/example.xlsx";
  //var bytes = File(file).readAsBytesSync();
  final excel = Excel.createExcel();
  // or
  //var excel = Excel.decodeBytes(bytes);

  final sheet = excel['Sheet1'];

  sheet.appendRow([
    const IntCellValue(8),
    const DoubleCellValue(999.62221),
    DateCellValue(year: DateTime.now().year, month: DateTime.now().month, day: DateTime.now().day),
    DateTimeCellValue.fromDateTime(DateTime.now()),
  ]);

  // Saving the file

  // String outputFile = "/Users/kawal/Desktop/git_projects/r.xlsx";
  final outputFile = './example/example.xlsx';

  //stopwatch.reset();
  final fileBytes = await excel.save();
  //print('saving executed in ${stopwatch.elapsed}');
  if (fileBytes != null) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}
