//
//  Worksheet.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent an Excel worksheet.
public struct Worksheet {
  private var lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>
  var name: String { String(cString: lxw_worksheet.pointee.name) }
  init(_ lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>) { self.lxw_worksheet = lxw_worksheet }
  /// Insert a chart object into a worksheet.
  public func insert(chart: Chart, _ pos: (row: Int, col: Int)) -> Worksheet {
    let r = UInt32(pos.row)
    let c = UInt16(pos.col)
    _ = worksheet_insert_chart(lxw_worksheet, r, c, chart.lxw_chart)
    return self
  }
  /// Insert a chart object into a worksheet, with options.
  public func insert(chart: Chart, _ pos: (row: Int, col: Int), scale: (x: Double, y: Double))
    -> Worksheet
  {
    let r = UInt32(pos.row)
    let c = UInt16(pos.col)
    var o = lxw_chart_options(
      x_offset: 0, y_offset: 0, x_scale: scale.x, y_scale: scale.y, object_position: 2,
      description: nil, decorative: 0)
    worksheet_insert_chart_opt(lxw_worksheet, r, c, chart.lxw_chart, &o)
    return self
  }
  /// Write a column of data starting from (row, col).
  @discardableResult public func write(column values: [Value], _ cell: Cell, format: Format? = nil)
    -> Worksheet
  {
    var r = cell.row
    let c = cell.col
    for value in values {
      write(value, .init(r, c), format: format)
      r += 1
    }
    return self
  }
  /// Write a row of data starting from (row, col).
  @discardableResult public func write(row values: [Value], _ cell: Cell, format: Format? = nil)
    -> Worksheet
  {
    let r = cell.row
    var c = cell.col
    for value in values {
      write(value, .init(r, c), format: format)
      c += 1
    }
    return self
  }
  /// Write a row of Double values starting from (row, col).
  @discardableResult public func write(
    _ numbers: [Double], row: Int, col: Int = 0, format: Format? = nil
  ) -> Worksheet {
    let f = format?.lxw_format
    let r = UInt32(row)
    var c = UInt16(col)
    for number in numbers {
      worksheet_write_number(lxw_worksheet, r, c, number, f)
      c += 1
    }

    return self
  }
  /// Write a row of String values starting from (row, col).
  @discardableResult public func write(
    _ strings: [String], row: Int, col: Int = 0, format: Format? = nil
  ) -> Worksheet {
    let f = format?.lxw_format
    let r = UInt32(row)
    var c = UInt16(col)
    for string in strings {
      let _ = string.withCString { s in worksheet_write_string(lxw_worksheet, r, c, s, f) }
      c += 1
    }

    return self
  }
  /// Write data to a worksheet cell by calling the appropriate
  /// worksheet_write_*() method based on the type of data being passed.
  @discardableResult public func write(_ value: Value, _ cell: Cell, format: Format? = nil)
    -> Worksheet
  {
    let r = cell.row
    let c = cell.col
    let f = format?.lxw_format
    let error: lxw_error
    switch value {
    case .number(let number): error = worksheet_write_number(lxw_worksheet, r, c, number, f)
    case .string(let string):
      error = string.withCString { s in worksheet_write_string(lxw_worksheet, r, c, s, f) }
    case .url(let url):
      error = url.absoluteString.withCString { s in worksheet_write_url(lxw_worksheet, r, c, s, f) }
    case .blank: error = worksheet_write_blank(lxw_worksheet, r, c, f)
    case .comment(let comment):
      error = comment.withCString { s in worksheet_write_comment(lxw_worksheet, r, c, s) }
    case .boolean(let boolean):
      error = worksheet_write_boolean(lxw_worksheet, r, c, Int32(boolean ? 1 : 0), f)
    case .formula(let formula):
      error = formula.withCString { s in worksheet_write_formula(lxw_worksheet, r, c, s, f) }
    case .datetime(let datetime):
      error = lxw_error(rawValue: 0)
      let num = (datetime.timeIntervalSince1970 / 86400) + 25569
      worksheet_write_number(lxw_worksheet, r, c, num, f)
    }
    if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }

    return self
  }
    
    // Added, RJ 2026-09-12: ----v    1️⃣ 2️⃣
    
    /// Freeze panes for scrolling; pass in the first does-scroll cell, so frozen is all to the left of, and all above, this cell
    @discardableResult public func freeze(_ cell: Cell) -> Worksheet {
        worksheet_freeze_panes(lxw_worksheet, cell.row, cell.col)
        return self
    }
    
    /// Selects the given cell. Use case: when user opens the workbook, this cell in this worksheet is where the cursor is visible
    @discardableResult public func selection(_ cell: Cell) -> Worksheet {
        worksheet_set_selection(lxw_worksheet, cell.row, cell.col, cell.row, cell.col)
        return self
    }
    
    /// Selects the given range. Use case: when user opens the workbook, this range in this worksheet is highlighted/visible as selected
    @discardableResult public func selection(_ range: Range) -> Worksheet {
        worksheet_set_selection(lxw_worksheet, range.row, range.col, range.row2, range.col2)
        return self
    }
        
    /// Populates both the formula and the initial (displayed before recalculation) value. Causes a fatalError immediately if the WRITE itself fails (e.g. invalid cell reference) — not checking the formula result.
    @discardableResult public func write(formula: String, result: Double, _ cell: Cell, format: Format? = nil) -> Worksheet {
        let error = formula.withCString { s in
            worksheet_write_formula_num(lxw_worksheet, cell.row, cell.col, s, format?.lxw_format, result)
        }
        if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
        return self
    }
    
    public struct CommentOptions {
        public var visible: UInt8
        public var author: String?
        public var width: UInt16
        public var height: UInt16
        public var xScale: Double
        public var yScale: Double
        public var color: Color?
        public var fontName: String?
        public var fontSize: Double
        public var fontFamily: UInt8
        public var start: Cell?
        public var xOffset: Int32
        public var yOffset: Int32
        
        public init(visible: UInt8 = 0, author: String? = nil, width: UInt16 = 128, height: UInt16 = 74,
                    xScale: Double = 1.0, yScale: Double = 1.0, color: Color? = nil,
                    fontName: String? = nil, fontSize: Double = 8, fontFamily: UInt8 = 0,
                    start: Cell? = nil, xOffset: Int32 = 0, yOffset: Int32 = 0) {
            self.visible = visible; self.author = author; self.width = width; self.height = height
            self.xScale = xScale; self.yScale = yScale; self.color = color
            self.fontName = fontName; self.fontSize = fontSize; self.fontFamily = fontFamily
            self.start = start; self.xOffset = xOffset; self.yOffset = yOffset
        }
    }
    
    private func withOptionalCString<R>(_ string: String?, _ body: (UnsafePointer<CChar>?) -> R) -> R {
        if let string = string { return string.withCString { body($0) } }
        return body(nil)
    }
    
    @discardableResult public func write(comment: String, _ cell: Cell, options: CommentOptions, format: Format? = nil) -> Worksheet {
        var error: lxw_error = LXW_NO_ERROR
        withOptionalCString(options.author) { authorPtr in
            withOptionalCString(options.fontName) { fontNamePtr in
                var cOptions = lxw_comment_options(
                    visible: options.visible,
                    author: authorPtr,
                    width: options.width,
                    height: options.height,
                    x_scale: options.xScale,
                    y_scale: options.yScale,
                    color: options.color?.hex ?? 0,
                    font_name: fontNamePtr,
                    font_size: options.fontSize,
                    font_family: options.fontFamily,
                    start_row: options.start?.row ?? 0,
                    start_col: options.start?.col ?? 0,
                    x_offset: options.xOffset,
                    y_offset: options.yOffset
                )
                error = comment.withCString { s in
                    worksheet_write_comment_opt(lxw_worksheet, cell.row, cell.col, s, &cOptions)
                }
            }
        }
        if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
        return self
    }
    
    public enum IgnoredErrorType: UInt8 {
        case numberStoredAsText = 1
        case evalError = 2
        case formulaDiffers = 3
        case formulaRange = 4
        case formulaUnlocked = 5
    }
    
    @discardableResult public func ignoreErrors(_ type: IgnoredErrorType, range: String) -> Worksheet {
        let error = range.withCString { s in
            worksheet_ignore_errors(lxw_worksheet, type.rawValue, s)
        }
        if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
        return self
    }

      /// Restricts the cell to a dropdown of the given values; anything else is rejected (Stop alert). Added RJ 2026-11-11
    @discardableResult public func validateList(_ values: [String], _ cell: Cell) -> Worksheet {
        let cStrings = values.map { strdup($0) }
        defer { cStrings.forEach { free($0) } }
        var list: [UnsafePointer<CChar>?] = cStrings.map { UnsafePointer($0) } + [nil]   // nil-terminated

        var dv = lxw_data_validation()
        dv.validate   = UInt8(LXW_VALIDATION_TYPE_LIST.rawValue)
        dv.error_type = UInt8(LXW_VALIDATION_ERROR_TYPE_STOP.rawValue)

        let error = list.withUnsafeMutableBufferPointer { buf -> lxw_error in
            dv.value_list = buf.baseAddress
            return worksheet_data_validation_cell(lxw_worksheet, cell.row, cell.col, &dv)
        }
        if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
        return self
    }

    /// Conditional format: when the cell's text equals `text`, apply `format` on top of its normal format. Added RJ 2026-11-12
    @discardableResult public func conditionalEquals(_ text: String, _ cell: Cell, format: Format) -> Worksheet {
        let error = "\"\(text)\"".withCString { s -> lxw_error in        // Excel wants the string quoted
            var cf = lxw_conditional_format()
            cf.type         = UInt8(LXW_CONDITIONAL_TYPE_CELL.rawValue)
            cf.criteria     = UInt8(LXW_CONDITIONAL_CRITERIA_EQUAL_TO.rawValue)
            cf.value_string = s
            cf.format       = format.lxw_format
            return worksheet_conditional_format_cell(lxw_worksheet, cell.row, cell.col, &cf)
        }
        if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
        return self
    }

  
    // Added, RJ 2026-09-12 and beyond  -----^    1️⃣ 2️⃣
    // And, for the ignored errors things, 9-18-2026 also.
    // And, for the validateList(), 2026-11-11 also.
    // And, with conditionalEquals() 2026-11-12
    
    
  /// Set a worksheet tab as selected.
  @discardableResult public func select() -> Worksheet {
    worksheet_select(lxw_worksheet)
    return self
  }
  /// Hide the current worksheet.
  @discardableResult public func hide() -> Worksheet {
    worksheet_hide(lxw_worksheet)
    return self
  }
  /// Make a worksheet the active, i.e., visible worksheet.
  @discardableResult public func activate() -> Worksheet {
    worksheet_activate(lxw_worksheet)
    return self
  }
  /// Hide zero values in worksheet cells.
  @discardableResult public func hide_zero() -> Worksheet {
    worksheet_hide_zero(lxw_worksheet)
    return self
  }
  /// Set the paper type for printing.
  @discardableResult public func paper(type: PaperType) -> Worksheet {
    worksheet_set_paper(lxw_worksheet, type.rawValue)
    return self
  }
  /// Set the properties for one or more columns of cells.
  @discardableResult public func column(_ cols: Cols, width: Double, format: Format? = nil)
    -> Worksheet
  {
    let first = cols.col
    let last = cols.col2
    let f = format?.lxw_format
    _ = worksheet_set_column(lxw_worksheet, first, last, width, f)
    return self
  }
  /// Set the properties for a row of cells
  @discardableResult public func row(_ row: UInt32, height: Double, format: Format? = nil) -> Worksheet
  {
    let f = format?.lxw_format
    _ = worksheet_set_row(lxw_worksheet, row, height, f)
    return self
  }
  /// Set the properties for one or more columns of cells.
  @discardableResult public func hide_columns(_ col: Int, width: Double = 8.43) -> Worksheet {
    let first = UInt16(col)
    let cols: Cols = "A:XFD"
    let last = cols.col2
    var o = lxw_row_col_options(hidden: 1, level: 0, collapsed: 0)
    _ = worksheet_set_column_opt(lxw_worksheet, first, last, width, nil, &o)
    return self
  }
  /// Set the color of the worksheet tab.
  @discardableResult public func tab(color: Color) -> Worksheet {
    worksheet_set_tab_color(lxw_worksheet, color.hex)
    return self
  }
  /// Set the default row properties.
  @discardableResult public func set_default(row_height: Double, hide_unused_rows: Bool = true)
    -> Worksheet
  {
    let hide: UInt8 = hide_unused_rows ? 1 : 0
    worksheet_set_default_row(lxw_worksheet, row_height, hide)
    return self
  }
  /// Set the print area for a worksheet.
  @discardableResult public func print_area(range: Range) -> Worksheet {
    let _ = worksheet_print_area(lxw_worksheet, range.row, range.col, range.row2, range.col2)
    return self
  }
  /// Set the autofilter area in the worksheet.
  @discardableResult public func autofilter(range: Range) -> Worksheet {
    let _ = worksheet_autofilter(lxw_worksheet, range.row, range.col, range.row2, range.col2)
    return self
  }
  /// Set the option to display or hide gridlines on the screen and the printed page.
  @discardableResult public func gridline(screen: Bool, print: Bool = false) -> Worksheet {
    worksheet_gridlines(lxw_worksheet, UInt8((print ? 2 : 0) + (screen ? 1 : 0)))
    return self
  }

  /// Set a table in the worksheet.
  @discardableResult public func table(
    range: Range, name: String? = nil, header: [(String, Format?)] = []
  ) -> Worksheet {
    table(
      range: range, name: name, header: header.map { $0.0 }, format: header.map { $0.1 },
      totalRow: [])
  }

  /// Merge a range of cells in the worksheet.
  @discardableResult public func merge(range: Range, string: String, format: Format? = nil)
    -> Worksheet
  {
    worksheet_merge_range(
      lxw_worksheet, range.row, range.col, range.row2, range.col2, string, format?.lxw_format)
    return self
  }

  /// Set a table in the worksheet.
  @discardableResult public func table(
    range: Range, name: String? = nil, header: [String] = [], format: [Format?] = [],
    totalRow: [TotalFunction] = []
  ) -> Worksheet {
    var options = lxw_table_options()
    if let name = name { options.name = makeCString(from: name) }
    options.style_type = UInt8(LXW_TABLE_STYLE_TYPE_MEDIUM.rawValue)
    options.style_type_number = 7
    options.total_row = totalRow.isEmpty ? UInt8(LXW_FALSE.rawValue) : UInt8(LXW_TRUE.rawValue)
    var table_columns = [lxw_table_column]()
    let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<lxw_table_column>?>.allocate(
      capacity: header.count + 1)
    defer { buffer.deallocate() }
    if !header.isEmpty {
      table_columns = Array(repeating: lxw_table_column(), count: header.count)
      for i in header.indices {
        table_columns[i].header = makeCString(from: header[i])
        if format.endIndex > i {
          table_columns[i].header_format = format[i]?.lxw_format
        }
        if totalRow.endIndex > i {
          table_columns[i].total_function = totalRow[i].rawValue
        }
        withUnsafeMutablePointer(to: &table_columns[i]) {
          buffer.baseAddress?.advanced(by: i).pointee = $0
        }
      }
      options.columns = buffer.baseAddress
    }
    _ = worksheet_add_table(
      lxw_worksheet, range.row, range.col, range.row2 + (totalRow.isEmpty ? 0 : 1), range.col2,
      &options)
    if let _ = name { options.name.deallocate() }
    table_columns.forEach { $0.header.deallocate() }
    return self
  }
    
    
    
    
    
    
    
}

private func makeCString(from str: String) -> UnsafePointer<CChar> {
  let count = str.utf8.count + 1
  let result = UnsafeMutablePointer<CChar>.allocate(capacity: count)
  str.withCString { result.initialize(from: $0, count: count) }
  return UnsafePointer(result)
}
