<?php

namespace Vtiful\Kernel;

/**
 * Class Excel
 *
 * @author  viest
 *
 * @package Vtiful\Kernel
 */
class Excel
{
    const TYPE_STRING = 0x01;
    const TYPE_INT = 0x02;
    const TYPE_DOUBLE = 0x04;
    const TYPE_TIMESTAMP = 0x08;

    const SKIP_NONE = 0x00;
    const SKIP_EMPTY_ROW = 0x01;
    const SKIP_EMPTY_CELLS = 0x02;

    /**
     * Excel constructor.
     *
     * @param array $config
     */
    public function __construct(array $config)
    {
        //
    }

    /**
     * File Name
     *
     * @param string $fileName
     * @param string $sheetName
     *
     * @return Excel
     *
     * @author viest
     */
    public function fileName(string $fileName, string $sheetName = 'Sheet1'): self
    {
        return $this;
    }

    /**
     * Const memory model
     *
     * @param string $fileName
     * @param string $sheetName
     *
     * @return Excel
     *
     * @author viest
     */
    public function constMemory(string $fileName, string $sheetName = 'Sheet1'): self
    {
        return $this;
    }

    /**
     * Add a new worksheet to a workbook.
     *
     * The worksheet name must be a valid Excel worksheet name, i.e. it must be
     * less than 32 character and it cannot contain any of the characters:
     *
     *     / \ [ ] : * ?
     *
     * In addition, you cannot use the same, case insensitive, `$sheetName` for more
     * than one worksheet.
     *
     * @param string|NULL $sheetName
     *
     * @return Excel
     *
     * @author viest
     */
    public function addSheet(?string $sheetName): self
    {
        return $this;
    }

    /**
     * Checkout worksheet
     *
     * @param string $sheetName
     *
     * @return Excel
     *
     * @author viest
     */
    public function checkoutSheet(string $sheetName): self
    {
        return $this;
    }

    /**
     * Insert data on the first line of the worksheet
     *
     * @param array $header
     *
     * @return Excel
     *
     * @author viest
     */
    public function header(array $header): self
    {
        return $this;
    }

    /**
     * Insert data on the worksheet
     *
     * @param array $data
     *
     * @return Excel
     *
     * @author viest
     */
    public function data(array $data): self
    {
        return $this;
    }

    /**
     * Generate file
     *
     * @return string
     *
     * @author viest
     */
    public function output(): string
    {
        return 'FilePath';
    }

    /**
     * Get file resource
     *
     * @return resource
     *
     * @author viest
     */
    public function getHandle()
    {
        //
    }

    /**
     * Auto filter on the worksheet
     *
     * @param string $range
     *
     * @return Excel
     *
     * @author viest
     */
    public function autoFilter(string $range): self
    {
        return $this;
    }

    /**
     * Insert data on the cell
     *
     * @param int               $row
     * @param int               $column
     * @param int|string|double $data
     * @param string|null       $format
     * @param resource|null     $formatHandle
     *
     * @return Excel
     *
     * @author viest
     */
    public function insertText(int $row, int $column, $data, string $format = NULL, $formatHandle = NULL): self
    {
        return $this;
    }

    /**
     * Insert date on the cell
     *
     * @param int           $row
     * @param int           $column
     * @param int           $timestamp
     * @param string|NULL   $format
     * @param resource|null $formatHandle
     *
     * @return Excel
     *
     * @author viest
     */
    public function insertDate(int $row, int $column, int $timestamp, string $format = NULL, $formatHandle = NULL): self
    {
        return $this;
    }

    /**
     * Insert chart on the cell
     *
     * @param int      $row
     * @param int      $column
     * @param resource $chartResource
     *
     * @return Excel
     *
     * @author viest
     */
    public function insertChart(int $row, int $column, $chartResource): self
    {
        return $this;
    }

    /**
     * Insert url on the cell
     *
     * @param int           $row
     * @param int           $column
     * @param string        $url
     * @param resource|null $formatHandle
     *
     * @return Excel
     *
     * @author viest
     */
    public function insertUrl(int $row, int $column, string $url, $formatHandle = NULL): self
    {
        return $this;
    }

    /**
     * Insert image on the cell
     *
     * @param int    $row
     * @param int    $column
     * @param string $imagePath
     * @param float  $width
     * @param float  $height
     *
     * @return Excel
     *
     * @author viest
     */
    public function insertImage(int $row, int $column, string $imagePath, float $width = 1, float $height = 1): self
    {
        return $this;
    }

    /**
     * Insert Formula on the cell
     *
     * @param int    $row
     * @param int    $column
     * @param string $formula
     *
     * @return Excel
     *
     * @author viest
     */
    public function insertFormula(int $row, int $column, string $formula): self
    {
        return $this;
    }

    /**
     * Merge cells
     *
     * @param string $range
     * @param string $data
     *
     * @return Excel
     *
     * @author viest
     */
    public function MergeCells(string $range, string $data): self
    {
        return $this;
    }

    /**
     * Set column cells width or format
     *
     * @param string        $range
     * @param float         $cellWidth
     * @param resource|null $formatHandle
     *
     * @return Excel
     *
     * @author viest
     */
    public function setColumn(string $range, float $cellWidth, $formatHandle = NULL): self
    {
        return $this;
    }

    /**
     * Set row cells height or format
     *
     * @param string        $range
     * @param float         $cellHeight
     * @param resource|null $formatHandle
     *
     * @return Excel
     *
     * @author viest
     */
    public function setRow(string $range, float $cellHeight, $formatHandle = NULL): self
    {
        return $this;
    }

    /**
     * Open xlsx file
     *
     * @param string $fileName
     *
     * @return Excel
     *
     * @author viest
     */
    public function openFile(string $fileName): self
    {
        return $this;
    }

    /**
     * Open sheet
     *
     * default open first sheet
     *
     * @param string|NULL $sheetName
     * @param int         skipFlag
     *
     * @return Excel
     *
     * @author viest
     */
    public function openSheet(string $sheetName = NULL, int $skipFlag = 0x00): self
    {
        return $this;
    }

    /**
     * Set row cell data type
     *
     * @param array $types
     *
     * @return Excel
     *
     * @author viest
     */
    public function setType(array $types): self
    {
        return $this;
    }

    /**
     * Read values from the sheet
     *
     * @return array
     *
     * @author viest
     */
    public function getSheetData(): array
    {
        return [];
    }

    /**
     * Read values from the sheet
     *
     * @return array
     *
     * @author viest
     */
    public function nextRow(): array
    {
        return [];
    }

    /**
     * Next Cell In Callback
     *
     * @param callable    $callback  function(int $row, int $cell, string $data)
     * @param string|NULL $sheetName sheet name
     *
     * @author viest
     */
    public function nextCellCallback(callable $callback, string $sheetName = NULL): void
    {
        //
    }

    /**
     * Freeze panes
     *
     * freezePanes(1, 0); // Freeze the first row.
     * freezePanes(0, 1); // Freeze the first column.
     * freezePanes(1, 1); // Freeze first row/column.
     *
     * @param int $row
     * @param int $column
     *
     * @return $this
     *
     * @author viest
     */
    public function freezePanes(int $row, int $column): self
    {
        return $this;
    }
}

/**
 * Class Format
 *
 * @author  viest
 *
 * @package Vtiful\Kernel
 */
class Format
{
    const UNDERLINE_SINGLE = 0x00;
    const UNDERLINE_DOUBLE = 0x00;
    const UNDERLINE_SINGLE_ACCOUNTING = 0x00;
    const UNDERLINE_DOUBLE_ACCOUNTING = 0x00;

    const FORMAT_ALIGN_LEFT = 0x00;
    const FORMAT_ALIGN_CENTER = 0x00;
    const FORMAT_ALIGN_RIGHT = 0x00;
    const FORMAT_ALIGN_FILL = 0x00;
    const FORMAT_ALIGN_JUSTIFY = 0x00;
    const FORMAT_ALIGN_CENTER_ACROSS = 0x00;
    const FORMAT_ALIGN_DISTRIBUTED = 0x00;
    const FORMAT_ALIGN_VERTICAL_TOP = 0x00;
    const FORMAT_ALIGN_VERTICAL_BOTTOM = 0x00;
    const FORMAT_ALIGN_VERTICAL_CENTER = 0x00;
    const FORMAT_ALIGN_VERTICAL_JUSTIFY = 0x00;
    const FORMAT_ALIGN_VERTICAL_DISTRIBUTED = 0x00;

    const COLOR_BLACK = 0x00;
    const COLOR_BLUE = 0x00;
    const COLOR_BROWN = 0x00;
    const COLOR_CYAN = 0x00;
    const COLOR_GRAY = 0x00;
    const COLOR_GREEN = 0x00;
    const COLOR_LIME = 0x00;
    const COLOR_MAGENTA = 0x00;
    const COLOR_NAVY = 0x00;
    const COLOR_ORANGE = 0x00;
    const COLOR_PINK = 0x00;
    const COLOR_PURPLE = 0x00;
    const COLOR_RED = 0x00;
    const COLOR_SILVER = 0x00;
    const COLOR_WHITE = 0x00;
    const COLOR_YELLOW = 0x00;

    const PATTERN_NONE = 0x00;
    const PATTERN_SOLID = 0x00;
    const PATTERN_MEDIUM_GRAY = 0x00;
    const PATTERN_DARK_GRAY = 0x00;
    const PATTERN_LIGHT_GRAY = 0x00;
    const PATTERN_DARK_HORIZONTAL = 0x00;
    const PATTERN_DARK_VERTICAL = 0x00;
    const PATTERN_DARK_DOWN = 0x00;
    const PATTERN_DARK_UP = 0x00;
    const PATTERN_DARK_GRID = 0x00;
    const PATTERN_DARK_TRELLIS = 0x00;
    const PATTERN_LIGHT_HORIZONTAL = 0x00;
    const PATTERN_LIGHT_VERTICAL = 0x00;
    const PATTERN_LIGHT_DOWN = 0x00;
    const PATTERN_LIGHT_UP = 0x00;
    const PATTERN_LIGHT_GRID = 0x00;
    const PATTERN_LIGHT_TRELLIS = 0x00;
    const PATTERN_GRAY_125 = 0x00;
    const PATTERN_GRAY_0625 = 0x00;

    const BORDER_THIN = 0x00;
    const BORDER_MEDIUM = 0x00;
    const BORDER_DASHED = 0x00;
    const BORDER_DOTTED = 0x00;
    const BORDER_THICK = 0x00;
    const BORDER_DOUBLE = 0x00;
    const BORDER_HAIR = 0x00;
    const BORDER_MEDIUM_DASHED = 0x00;
    const BORDER_DASH_DOT = 0x00;
    const BORDER_MEDIUM_DASH_DOT = 0x00;
    const BORDER_DASH_DOT_DOT = 0x00;
    const BORDER_MEDIUM_DASH_DOT_DOT = 0x00;
    const BORDER_SLANT_DASH_DOT = 0x00;

    /**
     * Format constructor.
     *
     * @param resource $fileHandle
     */
    public function __construct($fileHandle)
    {
        //
    }

    /**
     * Wrap
     *
     * @return Format
     *
     * @author viest
     */
    public function wrap(): self
    {
        return $this;
    }

    /**
     * Bold
     *
     * @return Format
     *
     * @author viest
     */
    public function bold(): self
    {
        return $this;
    }

    /**
     * Italic
     *
     * @return Format
     *
     * @author viest
     */
    public function italic(): self
    {
        return $this;
    }

    /**
     * Cells border
     *
     * @param int $style const BORDER_***
     *
     * @return Format
     *
     * @author viest
     */
    public function border(int $style): self
    {
        return $this;
    }

    /**
     * Align
     *
     * @param int ...$style const FORMAT_ALIGN_****
     *
     * @return Format
     *
     * @author viest
     */
    public function align(...$style): self
    {
        return $this;
    }

    /**
     * Number format
     *
     * @param string $format
     *
     * #,##0
     *
     * @return Format
     *
     * @author viest
     */
    public function number(string $format): self
    {
        return $this;
    }

    /**
     * Font color
     *
     * @param int $color const COLOR_****
     *
     * @return Format
     *
     * @author viest
     */
    public function fontColor(int $color): self
    {
        return $this;
    }

    /**
     * Font
     *
     * @param string $fontName
     *
     * @return Format
     *
     * @author viest
     */
    public function font(string $fontName): self
    {
        return $this;
    }

    /**
     * Font size
     *
     * @param float $size
     *
     * @return Format
     *
     * @author viest
     */
    public function fontSize(float $size): self
    {
        return $this;
    }

    /**
     * String strikeout
     *
     * @return Format
     *
     * @author viest
     */
    public function strikeout(): self
    {
        return $this;
    }

    /**
     * Underline
     *
     * @param int $style const UNDERLINE_****
     *
     * @return Format
     *
     * @author viest
     */
    public function underline(int $style): self
    {
        return $this;
    }

    /**
     * Cell background
     *
     * @param int $color   const COLOR_****
     * @param int $pattern const PATTERN_****
     *
     * @return Format
     *
     * @author viest
     */
    public function background(int $color, int $pattern = self::PATTERN_SOLID): self
    {
        return $this;
    }

    /**
     * Format to resource
     *
     * @return resource
     *
     * @author viest
     */
    public function toResource()
    {
        //
    }
}