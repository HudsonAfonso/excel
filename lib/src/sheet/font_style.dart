part of '../../excel.dart';

/// Styling class for cells
// ignore: must_be_immutable
class _FontStyle extends Equatable {
  ExcelColor? _fontColorHex = ExcelColor.black;
  String? fontFamily;
  FontScheme fontScheme = FontScheme.Unset;
  bool _bold = false, _italic = false;
  Underline underline = Underline.None;
  int? fontSize;

  _FontStyle({
    ExcelColor? fontColorHex = ExcelColor.black,
    this.fontSize,
    this.fontFamily,
    this.fontScheme = FontScheme.Unset,
    bool bold = false,
    this.underline = Underline.None,
    bool italic = false,
  }) {
    _bold = bold;

    _italic = italic;

    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex.colorHex).excelColor;
    } else {
      _fontColorHex = ExcelColor.black;
    }
  }

  /// Get Font Color
  ExcelColor get fontColor {
    return _fontColorHex ?? ExcelColor.black;
  }

  /// Set Font Color
  set fontColor(ExcelColor? fontColorHex) {
    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex.colorHex).excelColor;
    } else {
      _fontColorHex = ExcelColor.black;
    }
  }

  /// Get `Bold`
  bool get isBold {
    return _bold;
  }

  /// Set `Bold`
  set isBold(bool bold) {
    _bold = bold;
  }

  /// Get `Italic`
  bool get isItalic {
    return _italic;
  }

  /// Set `Italic`
  set isItalic(bool italic) {
    _italic = italic;
  }

  @override
  List<Object?> get props => [_bold, _italic, fontSize, underline, fontFamily, _fontColorHex];
}
