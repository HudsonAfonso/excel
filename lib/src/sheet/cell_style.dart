part of '../../excel.dart';

/// Styling class for cells
// ignore: must_be_immutable
class CellStyle extends Equatable {
  String _fontColorHex = ExcelColor.black.colorHex;
  String _backgroundColorHex = ExcelColor.none.colorHex;
  String? fontFamily;
  FontScheme fontScheme;
  HorizontalAlign horizontalAlignment = HorizontalAlign.Left;
  VerticalAlign verticalAlignment = VerticalAlign.Bottom;
  TextWrapping? wrap;
  bool _bold = false, _italic = false;
  Underline underline = Underline.None;
  int? fontSize;
  int _rotation = 0;
  Border _leftBorder;
  Border _rightBorder;
  Border _topBorder;
  Border _bottomBorder;
  Border _diagonalBorder;
  bool diagonalBorderUp = false;
  bool diagonalBorderDown = false;
  NumFormat numberFormat;

  CellStyle({
    ExcelColor fontColorHex = ExcelColor.black,
    ExcelColor backgroundColorHex = ExcelColor.none,
    this.fontSize,
    this.fontFamily,
    FontScheme? fontScheme,
    HorizontalAlign horizontalAlign = HorizontalAlign.Left,
    VerticalAlign verticalAlign = VerticalAlign.Bottom,
    TextWrapping? textWrapping,
    bool bold = false,
    this.underline = Underline.None,
    bool italic = false,
    int rotation = 0,
    Border? leftBorder,
    Border? rightBorder,
    Border? topBorder,
    Border? bottomBorder,
    Border? diagonalBorder,
    this.diagonalBorderUp = false,
    this.diagonalBorderDown = false,
    this.numberFormat = NumFormat.standard_0,
  }) : wrap = textWrapping,
       _bold = bold,
       _italic = italic,

       fontScheme = fontScheme ?? FontScheme.Unset,
       _rotation = rotation,
       _fontColorHex = _isColorAppropriate(fontColorHex.colorHex),
       _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex),
       verticalAlignment = verticalAlign,
       horizontalAlignment = horizontalAlign,
       _leftBorder = leftBorder ?? Border(),
       _rightBorder = rightBorder ?? Border(),
       _topBorder = topBorder ?? Border(),
       _bottomBorder = bottomBorder ?? Border(),
       _diagonalBorder = diagonalBorder ?? Border();

  CellStyle copyWith({
    ExcelColor? fontColorHexVal,
    ExcelColor? backgroundColorHexVal,
    String? fontFamilyVal,
    FontScheme? fontSchemeVal,
    HorizontalAlign? horizontalAlignVal,
    VerticalAlign? verticalAlignVal,
    TextWrapping? textWrappingVal,
    bool? boldVal,
    bool? italicVal,
    Underline? underlineVal,
    int? fontSizeVal,
    int? rotationVal,
    Border? leftBorderVal,
    Border? rightBorderVal,
    Border? topBorderVal,
    Border? bottomBorderVal,
    Border? diagonalBorderVal,
    bool? diagonalBorderUpVal,
    bool? diagonalBorderDownVal,
    NumFormat? numberFormat,
  }) {
    return CellStyle(
      fontColorHex: fontColorHexVal ?? _fontColorHex.excelColor,
      backgroundColorHex: backgroundColorHexVal ?? _backgroundColorHex.excelColor,
      fontFamily: fontFamilyVal ?? fontFamily,
      fontScheme: fontSchemeVal ?? fontScheme,
      horizontalAlign: horizontalAlignVal ?? horizontalAlignment,
      verticalAlign: verticalAlignVal ?? verticalAlignment,
      textWrapping: textWrappingVal ?? wrap,
      bold: boldVal ?? _bold,
      italic: italicVal ?? _italic,
      underline: underlineVal ?? underline,
      fontSize: fontSizeVal ?? fontSize,
      rotation: rotationVal ?? _rotation,
      leftBorder: leftBorderVal ?? _leftBorder,
      rightBorder: rightBorderVal ?? _rightBorder,
      topBorder: topBorderVal ?? _topBorder,
      bottomBorder: bottomBorderVal ?? _bottomBorder,
      diagonalBorder: diagonalBorderVal ?? _diagonalBorder,
      diagonalBorderUp: diagonalBorderUpVal ?? diagonalBorderUp,
      diagonalBorderDown: diagonalBorderDownVal ?? diagonalBorderDown,
      numberFormat: numberFormat ?? this.numberFormat,
    );
  }

  ///Get Font Color
  ///
  ExcelColor get fontColor {
    return _fontColorHex.excelColor;
  }

  ///Set Font Color
  ///
  set fontColor(ExcelColor fontColorHex) {
    _fontColorHex = _isColorAppropriate(fontColorHex.colorHex);
  }

  ///Get Background Color
  ///
  ExcelColor get backgroundColor {
    return _backgroundColorHex.excelColor;
  }

  ///Set Background Color
  ///
  set backgroundColor(ExcelColor backgroundColorHex) {
    _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex);
  }

  ///Get Rotation
  ///
  int get rotation {
    return _rotation;
  }

  ///Rotation varies from [90 to -90]
  ///
  set rotation(int rotate) {
    if (rotate > 90 || rotate < -90) {
      rotate = 0;
    }
    if (rotate < 0) {
      /// The value is from 0 to -90 so now make it absolute and add it to 90
      ///
      /// -(_rotate) + 90
      rotate = -rotate + 90;
    }
    _rotation = rotate;
  }

  ///Get `Bold`
  ///
  bool get isBold {
    return _bold;
  }

  ///Set `Bold`
  set isBold(bool bold) {
    _bold = bold;
  }

  ///Get `Italic`
  ///
  bool get isItalic {
    return _italic;
  }

  ///Set `Italic`
  ///
  set isItalic(bool italic) {
    _italic = italic;
  }

  ///Get `LeftBorder`
  ///
  Border get leftBorder {
    return _leftBorder;
  }

  ///Set `LeftBorder`
  ///
  set leftBorder(Border? leftBorder) {
    _leftBorder = leftBorder ?? Border();
  }

  ///Get `RightBorder`
  ///
  Border get rightBorder {
    return _rightBorder;
  }

  ///Set `RightBorder`
  ///
  set rightBorder(Border? rightBorder) {
    _rightBorder = rightBorder ?? Border();
  }

  ///Get `TopBorder`
  ///
  Border get topBorder {
    return _topBorder;
  }

  ///Set `TopBorder`
  ///
  set topBorder(Border? topBorder) {
    _topBorder = topBorder ?? Border();
  }

  ///Get `BottomBorder`
  ///
  Border get bottomBorder {
    return _bottomBorder;
  }

  ///Set `BottomBorder`
  ///
  set bottomBorder(Border? bottomBorder) {
    _bottomBorder = bottomBorder ?? Border();
  }

  ///Get `DiagonalBorder`
  ///
  Border get diagonalBorder {
    return _diagonalBorder;
  }

  ///Set `DiagonalBorder`
  ///
  set diagonalBorder(Border? diagonalBorder) {
    _diagonalBorder = diagonalBorder ?? Border();
  }

  @override
  List<Object?> get props => [
    _bold,
    _rotation,
    _italic,
    underline,
    fontSize,
    fontFamily,
    fontScheme,
    wrap,
    verticalAlignment,
    horizontalAlignment,
    _fontColorHex,
    _backgroundColorHex,
    _leftBorder,
    _rightBorder,
    _topBorder,
    _bottomBorder,
    _diagonalBorder,
    diagonalBorderUp,
    diagonalBorderDown,
    numberFormat,
  ];
}
