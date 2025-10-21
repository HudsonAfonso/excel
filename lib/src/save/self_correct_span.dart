part of '../../excel.dart';

///Self correct the spanning of rows and columns by checking their cross-sectional relationship between if exists.
void _selfCorrectSpanMap(Excel excel) {
  for (final key in excel._mergeChangeLook) {
    if (excel._sheetMap[key] != null && excel._sheetMap[key]!._spanList.isNotEmpty) {
      final spanList = List<_Span?>.from(excel._sheetMap[key]!._spanList);

      for (var i = 0; i < spanList.length; i++) {
        final checkerPos = spanList[i];
        if (checkerPos == null) {
          continue;
        }
        var startRow = checkerPos.rowSpanStart,
            startColumn = checkerPos.columnSpanStart,
            endRow = checkerPos.rowSpanEnd,
            endColumn = checkerPos.columnSpanEnd;

        for (var j = i + 1; j < spanList.length; j++) {
          final spanObj = spanList[j];
          if (spanObj == null) {
            continue;
          }

          final locationChange = _isLocationChangeRequired(startColumn, startRow, endColumn, endRow, spanObj);
          if (locationChange.$1) {
            startColumn = locationChange.$2.$1;
            startRow = locationChange.$2.$2;
            endColumn = locationChange.$2.$3;
            endRow = locationChange.$2.$4;
            spanList[j] = null;
          } else {
            final locationChange2 = _isLocationChangeRequired(
              spanObj.columnSpanStart,
              spanObj.rowSpanStart,
              spanObj.columnSpanEnd,
              spanObj.rowSpanEnd,
              checkerPos,
            );

            if (locationChange2.$1) {
              startColumn = locationChange2.$2.$1;
              startRow = locationChange2.$2.$2;
              endColumn = locationChange2.$2.$3;
              endRow = locationChange2.$2.$4;
              spanList[j] = null;
            }
          }
        }
        final spanObj1 = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
        spanList[i] = spanObj1;
      }
      excel._sheetMap[key]!._spanList = List<_Span?>.from(spanList);
      excel._sheetMap[key]!._cleanUpSpanMap();
    }
  }
}
