"""Constants for HWP operations."""

# HWP COM Object ProgID
HWP_PROGID = "HWPFrame.HwpObject"

# HWP File Formats
class FileFormat:
    """HWP file format constants."""

    HWP = "HWP"  # 한글 문서 (.hwp)
    HWPX = "HWPX"  # 한글 문서 XML (.hwpx)
    HWT = "HWT"  # 한글 서식 (.hwt)
    HTML = "HTML"  # HTML 문서
    TEXT = "TEXT"  # 텍스트 문서
    PDF = "PDF"  # PDF 문서


# HWP Save Formats (for SaveAs)
class SaveFormat:
    """Save format codes for HWP."""

    HWP = "HWP"
    HWPX = "HWPX"
    HWT = "HWT"
    HTML = "HTML"
    TEXT = "TEXT"
    PDF = "PDF"


# HWP Actions
class Action:
    """HWP action names for HAction.Run()."""

    # Document actions
    FILE_NEW = "FileNew"
    FILE_OPEN = "FileOpen"
    FILE_SAVE = "FileSave"
    FILE_SAVE_AS = "FileSaveAs"
    FILE_CLOSE = "FileClose"

    # Cursor movement
    MOVE_DOC_BEGIN = "MoveDocBegin"
    MOVE_DOC_END = "MoveDocEnd"
    MOVE_LINE_BEGIN = "MoveLineBegin"
    MOVE_LINE_END = "MoveLineEnd"
    MOVE_PARA_BEGIN = "MoveParaBegin"
    MOVE_PARA_END = "MoveParaEnd"
    MOVE_LEFT = "MoveLeft"
    MOVE_RIGHT = "MoveRight"
    MOVE_UP = "MoveUp"
    MOVE_DOWN = "MoveDown"
    MOVE_NEXT_PARA = "MoveNextPara"
    MOVE_PREV_PARA = "MovePrevPara"

    # Selection
    SELECT_ALL = "SelectAll"

    # Text actions
    INSERT_TEXT = "InsertText"
    DELETE = "Delete"

    # Table actions
    TABLE_CREATE = "TableCreate"
    TABLE_APPEND_ROW = "TableAppendRow"
    TABLE_DELETE_ROW = "TableDeleteRow"
    TABLE_CELL_BLOCK = "TableCellBlock"

    # Find/Replace
    FIND = "Find"
    REPLACE = "Replace"
    REPLACE_ALL = "ReplaceAll"

    # Formatting
    CHAR_SHAPE = "CharShape"
    PARA_SHAPE = "ParaShape"

    # Export
    PRINT = "Print"


# Paragraph Alignment
class Alignment:
    """Paragraph alignment constants."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2
    JUSTIFY = 3
    DISTRIBUTE = 4


# Font Style
class FontStyle:
    """Font style constants."""

    NORMAL = 0
    BOLD = 1
    ITALIC = 2
    UNDERLINE = 4
    STRIKEOUT = 8


# Character Unit (in HWPUNIT, 1pt = 100 HWPUNIT)
HWPUNIT_PER_PT = 100
