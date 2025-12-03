using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal static class XlsFormulaReader
{
    private static readonly Dictionary<ushort, (string, int?)> FtabFunctionNames = new Dictionary<ushort, (string, int?)>
    {
        { 0x0000, ("COUNT", null) },
        { 0x0001, ("IF", null) },
        { 0x0002, ("ISNA", 1) },
        { 0x0003, ("ISERROR", 1) },
        { 0x0004, ("SUM", null) },
        { 0x0005, ("AVERAGE", null) },
        { 0x0006, ("MIN", null) },
        { 0x0007, ("MAX", null) },
        { 0x0008, ("ROW", null) },
        { 0x0009, ("COLUMN", null) },
        { 0x000A, ("NA", 0) },
        { 0x000B, ("NPV", null) },
        { 0x000C, ("STDEV", null) },
        { 0x000D, ("DOLLAR", null) },
        { 0x000E, ("FIXED", null) },
        { 0x000F, ("SIN", 1) },
        { 0x0010, ("COS", 1) },
        { 0x0011, ("TAN", 1) },
        { 0x0012, ("ATAN", 1) },
        { 0x0013, ("PI", 0) },
        { 0x0014, ("SQRT", 1) },
        { 0x0015, ("EXP", 1) },
        { 0x0016, ("LN", 1) },
        { 0x0017, ("LOG10", 1) },
        { 0x0018, ("ABS", 1) },
        { 0x0019, ("INT", 1) },
        { 0x001A, ("SIGN", 1) },
        { 0x001B, ("ROUND", 2) },
        { 0x001C, ("LOOKUP", null) },
        { 0x001D, ("INDEX", null) },
        { 0x001E, ("REPT", 2) },
        { 0x001F, ("MID", 3) },
        { 0x0020, ("LEN", 1) },
        { 0x0021, ("VALUE", 1) },
        { 0x0022, ("TRUE", 0) },
        { 0x0023, ("FALSE", 0) },
        { 0x0024, ("AND", null) },
        { 0x0025, ("OR", null) },
        { 0x0026, ("NOT", 1) },
        { 0x0027, ("MOD", 2) },
        { 0x0028, ("DCOUNT", 3) },
        { 0x0029, ("DSUM", 3) },
        { 0x002A, ("DAVERAGE", 3) },
        { 0x002B, ("DMIN", 3) },
        { 0x002C, ("DMAX", 3) },
        { 0x002D, ("DSTDEV", 3) },
        { 0x002E, ("VAR", null) },
        { 0x002F, ("DVAR", 3) },
        { 0x0030, ("TEXT", 2) },
        { 0x0031, ("LINEST", null) },
        { 0x0032, ("TREND", null) },
        { 0x0033, ("LOGEST", null) },
        { 0x0034, ("GROWTH", null) },
        { 0x0035, ("GOTO", 1) },
        { 0x0036, ("HALT", 0) },
        { 0x0037, ("RETURN", null) },
        { 0x0038, ("PV", null) },
        { 0x0039, ("FV", null) },
        { 0x003A, ("NPER", null) },
        { 0x003B, ("PMT", null) },
        { 0x003C, ("RATE", null) },
        { 0x003D, ("MIRR", 3) },
        { 0x003E, ("IRR", null) },
        { 0x003F, ("RAND", 0) },
        { 0x0040, ("MATCH", null) },
        { 0x0041, ("DATE", 3) },
        { 0x0042, ("TIME", 3) },
        { 0x0043, ("DAY", 1) },
        { 0x0044, ("MONTH", 1) },
        { 0x0045, ("YEAR", 1) },
        { 0x0046, ("WEEKDAY", 1) },
        { 0x0047, ("HOUR", 1) },
        { 0x0048, ("MINUTE", 1) },
        { 0x0049, ("SECOND", 1) },
        { 0x004A, ("NOW", 0) },
        { 0x004B, ("AREAS", 1) },
        { 0x004C, ("ROWS", 1) },
        { 0x004D, ("COLUMNS", 1) },
        { 0x004E, ("OFFSET", null) },
        { 0x004F, ("ABSREF", 2) },
        { 0x0050, ("RELREF", 2) },
        { 0x0051, ("ARGUMENT", null) },
        { 0x0052, ("SEARCH", null) },
        { 0x0053, ("TRANSPOSE", 1) },
        { 0x0054, ("ERROR", null) },
        { 0x0055, ("STEP", 0) },
        { 0x0056, ("TYPE", 1) },
        { 0x0057, ("ECHO", null) },
        { 0x0058, ("SET.NAME", null) },
        { 0x0059, ("CALLER", 0) },
        { 0x005A, ("DEREF", 1) },
        { 0x005B, ("WINDOWS", 0) },
        { 0x005C, ("SERIES", null) },
        { 0x005D, ("DOCUMENTS", 0) },
        { 0x005E, ("ACTIVE.CELL", 0) },
        { 0x005F, ("SELECTION", 0) },
        { 0x0060, ("RESULT", 1) },
        { 0x0061, ("ATAN2", 2) },
        { 0x0062, ("ASIN", 1) },
        { 0x0063, ("ACOS", 1) },
        { 0x0064, ("CHOOSE", null) },
        { 0x0065, ("HLOOKUP", 3) },
        { 0x0066, ("VLOOKUP", 3) }, // Has both fixed and variable parameter counts.
        { 0x0067, ("LINKS", null) },
        { 0x0068, ("INPUT", null) },
        { 0x0069, ("ISREF", 1) },
        { 0x006A, ("GET.FORMULA", 1) },
        { 0x006B, ("GET.NAME", 1) },
        { 0x006C, ("SET.VALUE", 2) },
        { 0x006D, ("LOG", null) },
        { 0x006E, ("EXEC", null) },
        { 0x006F, ("CHAR", 1) },
        { 0x0070, ("LOWER", 1) },
        { 0x0071, ("UPPER", 1) },
        { 0x0072, ("PROPER", 1) },
        { 0x0073, ("LEFT", null) },
        { 0x0074, ("RIGHT", null) },
        { 0x0075, ("EXACT", 2) },
        { 0x0076, ("TRIM", 1) },
        { 0x0077, ("REPLACE", 4) },
        { 0x0078, ("SUBSTITUTE", null) },
        { 0x0079, ("CODE", 1) },
        { 0x007A, ("NAMES", null) },
        { 0x007B, ("DIRECTORY", null) },
        { 0x007C, ("FIND", null) },
        { 0x007D, ("CELL", null) },
        { 0x007E, ("ISERR", 1) },
        { 0x007F, ("ISTEXT", 1) },
        { 0x0080, ("ISNUMBER", 1) },
        { 0x0081, ("ISBLANK", 1) },
        { 0x0082, ("T", 1) },
        { 0x0083, ("N", 1) },
        { 0x0084, ("FOPEN", null) },
        { 0x0085, ("FCLOSE", 1) },
        { 0x0086, ("FSIZE", 1) },
        { 0x0087, ("FREADLN", 1) },
        { 0x0088, ("FREAD", 2) },
        { 0x0089, ("FWRITELN", 2) },
        { 0x008A, ("FWRITE", 2) },
        { 0x008B, ("FPOS", null) },
        { 0x008C, ("DATEVALUE", 1) },
        { 0x008D, ("TIMEVALUE", 1) },
        { 0x008E, ("SLN", 3) },
        { 0x008F, ("SYD", 4) },
        { 0x0090, ("DDB", 4) }, // Has both fixed and variable parameter counts.
        { 0x0091, ("GET.DEF", null) },
        { 0x0092, ("REFTEXT", null) },
        { 0x0093, ("TEXTREF", null) },
        { 0x0094, ("INDIRECT", null) },
        { 0x0095, ("REGISTER", 3) },
        { 0x0096, ("CALL", null) },
        { 0x0097, ("ADD.BAR", 0) },
        { 0x0098, ("ADD.MENU", 2) },
        { 0x0099, ("ADD.COMMAND", 3) },
        { 0x009A, ("ENABLE.COMMAND", 4) },
        { 0x009B, ("CHECK.COMMAND", 4) },
        { 0x009C, ("RENAME.COMMAND", 4) },
        { 0x009D, ("SHOW.BAR", null) },
        { 0x009E, ("DELETE.MENU", 2) },
        { 0x009F, ("DELETE.COMMAND", 3) },
        { 0x00A0, ("GET.CHART.ITEM", null) },
        { 0x00A1, ("DIALOG.BOX", 1) },
        { 0x00A2, ("CLEAN", 1) },
        { 0x00A3, ("MDETERM", 1) },
        { 0x00A4, ("MINVERSE", 1) },
        { 0x00A5, ("MMULT", 2) },
        { 0x00A6, ("FILES", null) },
        { 0x00A7, ("IPMT", null) },
        { 0x00A8, ("PPMT", null) },
        { 0x00A9, ("COUNTA", null) },
        { 0x00AA, ("CANCEL.KEY", null) },
        { 0x00AB, ("FOR", null) },
        { 0x00AC, ("WHILE", 1) },
        { 0x00AD, ("BREAK", 0) },
        { 0x00AE, ("NEXT", 0) },
        { 0x00AF, ("INITIATE", 2) },
        { 0x00B0, ("REQUEST", 2) },
        { 0x00B1, ("POKE", 3) },
        { 0x00B2, ("EXECUTE", 2) },
        { 0x00B3, ("TERMINATE", 1) },
        { 0x00B4, ("RESTART", null) },
        { 0x00B5, ("HELP", null) },
        { 0x00B6, ("GET.BAR", 0) },
        { 0x00B7, ("PRODUCT", null) },
        { 0x00B8, ("FACT", 1) },
        { 0x00B9, ("GET.CELL", null) },
        { 0x00BA, ("GET.WORKSPACE", 1) },
        { 0x00BB, ("GET.WINDOW", null) },
        { 0x00BC, ("GET.DOCUMENT", null) },
        { 0x00BD, ("DPRODUCT", 3) },
        { 0x00BE, ("ISNONTEXT", 1) },
        { 0x00BF, ("GET.NOTE", null) },
        { 0x00C0, ("NOTE", null) },
        { 0x00C1, ("STDEVP", null) },
        { 0x00C2, ("VARP", null) },
        { 0x00C3, ("DSTDEVP", 3) },
        { 0x00C4, ("DVARP", 3) },
        { 0x00C5, ("TRUNC", 1) }, // Has both fixed and variable parameter counts.
        { 0x00C6, ("ISLOGICAL", 1) },
        { 0x00C7, ("DCOUNTA", 3) },
        { 0x00C8, ("DELETE.BAR", 1) },
        { 0x00C9, ("UNREGISTER", 1) }, // Mac only.

        // Excel 3.0
        { 0x00CC, ("USDOLLAR", null) }, // Renamed from YEN. Excel 4.0 for Mac.
        { 0x00CD, ("FINDB", null) }, // Excel 4.0 for Mac.
        { 0x00CE, ("SEARCHB", null) }, // Excel 4.0 for Mac.
        { 0x00CF, ("REPLACEB", 4) }, // Excel 4.0 for Mac.
        { 0x00D0, ("LEFTB", null) }, // Excel 4.0 for Mac.
        { 0x00D1, ("RIGHTB", null) }, // Excel 4.0 for Mac.
        { 0x00D2, ("MIDB", 3) }, // Excel 4.0 for Mac.
        { 0x00D3, ("LENB", 1) }, // Excel 4.0 for Mac.
        { 0x00D4, ("ROUNDUP", 2) }, // Excel 4.0 for Mac.
        { 0x00D5, ("ROUNDDOWN", 2) }, // Excel 4.0 for Mac.
        { 0x00D6, ("ASC", 1) }, // Excel 4.0 for Mac.
        { 0x00D7, ("DBCS", 1) }, // Renamed from JIS. Excel 4.0 for Mac.
        { 0x00D8, ("RANK", null) },
        { 0x00DB, ("ADDRESS", null) },
        { 0x00DC, ("DAYS360", 2) },
        { 0x00DD, ("TODAY", 0) },
        { 0x00DE, ("VDB", null) },
        { 0x00DF, ("ELSE", 0) },
        { 0x00E0, ("ELSE.IF", 1) },
        { 0x00E1, ("END.IF", 0) },
        { 0x00E2, ("FOR.CELL", null) },
        { 0x00E3, ("MEDIAN", null) },
        { 0x00E4, ("SUMPRODUCT", null) },
        { 0x00E5, ("SINH", 1) },
        { 0x00E6, ("COSH", 1) },
        { 0x00E7, ("TANH", 1) },
        { 0x00E8, ("ASINH", 1) },
        { 0x00E9, ("ACOSH", 1) },
        { 0x00EA, ("ATANH", 1) },
        { 0x00EB, ("DGET", 3) },
        { 0x00EC, ("CREATE.OBJECT", null) },
        { 0x00ED, ("VOLATILE", 0) },
        { 0x00EE, ("LAST.ERROR", 0) },
        { 0x00EF, ("CUSTOM.UNDO", null) },
        { 0x00F0, ("CUSTOM.REPEAT", null) },
        { 0x00F1, ("FORMULA.CONVERT", null) },
        { 0x00F2, ("GET.LINK.INFO", null) },
        { 0x00F3, ("TEXT.BOX", null) },
        { 0x00F4, ("INFO", 1) },
        { 0x00F5, ("GROUP", 0) },
        { 0x00F6, ("GET.OBJECT", null) },

        // Excel 4.0
        { 0x00F7, ("DB", null) },
        { 0x00F8, ("PAUSE", null) },
        { 0x00FB, ("RESUME", null) },
        { 0x00FC, ("FREQUENCY", 2) },
        { 0x00FD, ("ADD.TOOLBAR", null) },
        { 0x00FE, ("DELETE.TOOLBAR", 1) },
        { 0x00FF, ("USERDEFINED", null) },
        { 0x0100, ("RESET.TOOLBAR", 1) },
        { 0x0101, ("EVALUATE", 1) },
        { 0x0102, ("GET.TOOLBAR", null) },
        { 0x0103, ("GET.TOOL", null) },
        { 0x0104, ("SPELLING.CHECK", null) },
        { 0x0105, ("ERROR.TYPE", 1) },
        { 0x0106, ("APP.TITLE", null) },
        { 0x0107, ("WINDOW.TITLE", null) },
        { 0x0108, ("SAVE.TOOLBAR", null) },
        { 0x0109, ("ENABLE.TOOL", 3) },
        { 0x010A, ("PRESS.TOOL", 3) },
        { 0x010B, ("REGISTER.ID", null) },
        { 0x010C, ("GET.WORKBOOK", null) },
        { 0x010D, ("AVEDEV", null) },
        { 0x010E, ("BETADIST", null) },
        { 0x010F, ("GAMMALN", 1) },
        { 0x0110, ("BETAINV", null) },
        { 0x0111, ("BINOMDIST", 4) },
        { 0x0112, ("CHIDIST", 2) },
        { 0x0113, ("CHIINV", 2) },
        { 0x0114, ("COMBIN", 2) },
        { 0x0115, ("CONFIDENCE", 3) },
        { 0x0116, ("CRITBINOM", 3) },
        { 0x0117, ("EVEN", 1) },
        { 0x0118, ("EXPONDIST", 3) },
        { 0x0119, ("FDIST", 3) },
        { 0x011A, ("FINV", 3) },
        { 0x011B, ("FISHER", 1) },
        { 0x011C, ("FISHERINV", 1) },
        { 0x011D, ("FLOOR", 2) },
        { 0x011E, ("GAMMADIST", 4) },
        { 0x011F, ("GAMMAINV", 3) },
        { 0x0120, ("CEILING", 2) },
        { 0x0121, ("HYPGEOMDIST", 4) },
        { 0x0122, ("LOGNORMDIST", 3) },
        { 0x0123, ("LOGINV", 3) },
        { 0x0124, ("NEGBINOMDIST", 3) },
        { 0x0125, ("NORMDIST", 4) },
        { 0x0126, ("NORMSDIST", 1) },
        { 0x0127, ("NORMINV", 3) },
        { 0x0128, ("NORMSINV", 1) },
        { 0x0129, ("STANDARDIZE", 3) },
        { 0x012A, ("ODD", 1) },
        { 0x012B, ("PERMUT", 2) },
        { 0x012C, ("POISSON", 3) },
        { 0x012D, ("TDIST", 3) },
        { 0x012E, ("WEIBULL", 4) },
        { 0x012F, ("SUMXMY2", 2) },
        { 0x0130, ("SUMX2MY2", 2) },
        { 0x0131, ("SUMX2PY2", 2) },
        { 0x0132, ("CHITEST", 2) },
        { 0x0133, ("CORREL", 2) },
        { 0x0134, ("COVAR", 2) },
        { 0x0135, ("FORECAST", 3) },
        { 0x0136, ("FTEST", 2) },
        { 0x0137, ("INTERCEPT", 2) },
        { 0x0138, ("PEARSON", 2) },
        { 0x0139, ("RSQ", 2) },
        { 0x013A, ("STEYX", 2) },
        { 0x013B, ("SLOPE", 2) },
        { 0x013C, ("TTEST", 4) },
        { 0x013D, ("PROB", null) },
        { 0x013E, ("DEVSQ", null) },
        { 0x013F, ("GEOMEAN", null) },
        { 0x0140, ("HARMEAN", null) },
        { 0x0141, ("SUMSQ", null) },
        { 0x0142, ("KURT", null) },
        { 0x0143, ("SKEW", null) },
        { 0x0144, ("ZTEST", null) },
        { 0x0145, ("LARGE", 2) },
        { 0x0146, ("SMALL", 2) },
        { 0x0147, ("QUARTILE", 2) },
        { 0x0148, ("PERCENTILE", 2) },
        { 0x0149, ("PERCENTRANK", null) },
        { 0x014A, ("MODE", null) },
        { 0x014B, ("TRIMMEAN", 2) },
        { 0x014C, ("TINV", 2) },

        // Excel 5.0
        { 0x014E, ("MOVIE.COMMAND", null) }, // Unknown.
        { 0x014F, ("GET.MOVIE", null) }, // Unknown.
        { 0x0150, ("CONCATENATE", null) },
        { 0x0151, ("POWER", 2) },
        { 0x0152, ("PIVOT.ADD.DATA", null) },
        { 0x0153, ("GET.PIVOT.TABLE", null) },
        { 0x0154, ("GET.PIVOT.FIELD", null) },
        { 0x0155, ("GET.PIVOT.ITEM", null) },
        { 0x0156, ("RADIANS", 1) },
        { 0x0157, ("DEGREES", 1) },
        { 0x0158, ("SUBTOTAL", null) },
        { 0x0159, ("SUMIF", null) },
        { 0x015A, ("COUNTIF", 2) },
        { 0x015B, ("COUNTBLANK", 1) },
        { 0x015C, ("SCENARIO.GET", null) },
        { 0x015D, ("OPTIONS.LISTS.GET", 1) },
        { 0x015E, ("ISPMT", 4) },
        { 0x015F, ("DATEDIF", 3) },
        { 0x0160, ("DATESTRING", 1) },
        { 0x0161, ("NUMBERSTRING", 2) },
        { 0x0162, ("ROMAN", null) },
        { 0x0163, ("OPEN.DIALOG", null) },
        { 0x0164, ("SAVE.DIALOG", null) },

        // BIFF 8
        { 0x0165, ("VIEW.GET", null) },
        { 0x0166, ("GETPIVOTDATA", null) },
        { 0x0167, ("HYPERLINK", null) },
        { 0x0168, ("PHONETIC", 1) },
    };

    private static readonly Dictionary<ushort, string> MacroCommandNames = new Dictionary<ushort, string>
    {
        { 0x0000, "BEEP" },
        { 0x0001, "OPEN" },
        { 0x0002, "OPEN.LINKS" },
        { 0x0003, "CLOSE.ALL" },
        { 0x0004, "SAVE" },
        { 0x0005, "SAVE.AS" },
        { 0x0006, "FILE.DELETE" },
        { 0x0007, "PAGE.SETUP" },
        { 0x0008, "PRINT" },
        { 0x0009, "PRINTER.SETUP" },
        { 0x000A, "QUIT" },
        { 0x000B, "NEW.WINDOW" },
        { 0x000C, "ARRANGE.ALL" },
        { 0x000D, "SIZE" },
        { 0x000E, "MOVE" },
        { 0x000F, "FULL" },
        { 0x0010, "CLOSE" },
        { 0x0011, "RUN" },
        { 0x0016, "SET.PRINT.AREA" },
        { 0x0017, "SET.PRINT.TITLES" },
        { 0x0018, "SET.PAGE.BREAK" },
        { 0x0019, "REMOVE.PAGE.BREAK" },
        { 0x001A, "FONT" },
        { 0x001B, "DISPLAY" },
        { 0x001C, "PROTECT.DOCUMENT" },
        { 0x001D, "PRECISION" },
        { 0x001E, "A1.R1C1" },
        { 0x001F, "CALCULATE.NOW" },
        { 0x0020, "CALCULATION" },
        { 0x0022, "DATA.FIND" },
        { 0x0023, "EXTRACT" },
        { 0x0024, "DATA.DELETE" },
        { 0x0025, "SET.DATABASE" },
        { 0x0026, "SET.CRITERIA" },
        { 0x0027, "SORT" },
        { 0x0028, "DATA.SERIES" },
        { 0x0029, "TABLE" },
        { 0x002A, "FORMAT.NUMBER" },
        { 0x002B, "ALIGNMENT" },
        { 0x002C, "STYLE" },
        { 0x002D, "BORDER" },
        { 0x002E, "CELL.PROTECTION" },
        { 0x002F, "COLUMN.WIDTH" },
        { 0x0030, "UNDO" },
        { 0x0031, "CUT" },
        { 0x0032, "COPY" },
        { 0x0033, "PASTE" },
        { 0x0034, "CLEAR" },
        { 0x0035, "PASTE.SPECIAL" },
        { 0x0036, "EDIT.DELETE" },
        { 0x0037, "INSERT" },
        { 0x0038, "FILL.RIGHT" },
        { 0x0039, "FILL.DOWN" },
        { 0x003D, "DEFINE.NAME" },
        { 0x003E, "CREATE.NAMES" },
        { 0x003F, "FORMULA.GOTO" },
        { 0x0040, "FORMULA.FIND" },
        { 0x0041, "SELECT.LAST.CELL" },
        { 0x0042, "SHOW.ACTIVE.CELL" },
        { 0x0043, "GALLERY.AREA" },
        { 0x0044, "GALLERY.BAR" },
        { 0x0045, "GALLERY.COLUMN" },
        { 0x0046, "GALLERY.LINE" },
        { 0x0047, "GALLERY.PIE" },
        { 0x0048, "GALLERY.SCATTER" },
        { 0x0049, "COMBINATION" },
        { 0x004A, "PREFERRED" },
        { 0x004B, "ADD.OVERLAY" },
        { 0x004C, "GRIDLINES" },
        { 0x004D, "SET.PREFERRED" },
        { 0x004E, "AXES" },
        { 0x004F, "LEGEND" },
        { 0x0050, "ATTACH.TEXT" },
        { 0x0051, "ADD.ARROW" },
        { 0x0052, "SELECT.CHART" },
        { 0x0053, "SELECT.PLOT.AREA" },
        { 0x0054, "PATTERNS" },
        { 0x0055, "MAIN.CHART" },
        { 0x0056, "OVERLAY" },
        { 0x0057, "SCALE" },
        { 0x0058, "FORMAT.LEGEND" },
        { 0x0059, "FORMAT.TEXT" },
        { 0x005A, "EDIT.REPEAT" }, // Excel 3.0
        { 0x005B, "PARSE" },
        { 0x005C, "JUSTIFY" },
        { 0x005D, "HIDE" },
        { 0x005E, "UNHIDE" },
        { 0x005F, "WORKSPACE" },
        { 0x0060, "FORMULA" },
        { 0x0061, "FORMULA.FILL" },
        { 0x0062, "FORMULA.ARRAY" },
        { 0x0063, "DATA.FIND.NEXT" },
        { 0x0064, "DATA.FIND.PREV" },
        { 0x0065, "FORMULA.FIND.NEXT" },
        { 0x0066, "FORMULA.FIND.PREV" },
        { 0x0067, "ACTIVATE" },
        { 0x0068, "ACTIVATE.NEXT" },
        { 0x0069, "ACTIVATE.PREV" },
        { 0x006A, "UNLOCKED.NEXT" },
        { 0x006B, "UNLOCKED.PREV" },
        { 0x006C, "COPY.PICTURE" },
        { 0x006D, "SELECT" },
        { 0x006E, "DELETE.NAME" },
        { 0x006F, "DELETE.FORMAT" },
        { 0x0070, "VLINE" },
        { 0x0071, "HLINE" },
        { 0x0072, "VPAGE" },
        { 0x0073, "HPAGE" },
        { 0x0074, "VSCROLL" },
        { 0x0075, "HSCROLL" },
        { 0x0076, "ALERT" },
        { 0x0077, "NEW" },
        { 0x0078, "CANCEL.COPY" },
        { 0x0079, "SHOW.CLIPBOARD" },
        { 0x007A, "MESSAGE" },
        { 0x007C, "PASTE.LINK" },
        { 0x007D, "APP.ACTIVATE" },
        { 0x007E, "DELETE.ARROW" },
        { 0x007F, "ROW.HEIGHT" },
        { 0x0080, "FORMAT.MOVE" },
        { 0x0081, "FORMAT.SIZE" },
        { 0x0082, "FORMULA.REPLACE" },
        { 0x0083, "SEND.KEYS" },
        { 0x0084, "SELECT.SPECIAL" },
        { 0x0085, "APPLY.NAMES" },
        { 0x0086, "REPLACE.FONT" },
        { 0x0087, "FREEZE.PANES" },
        { 0x0088, "SHOW.INFO" },
        { 0x0089, "SPLIT" },
        { 0x008A, "ON.WINDOW" },
        { 0x008B, "ON.DATA" },
        { 0x008C, "DISABLE.INPUT" },
        { 0x008E, "OUTLINE" }, // Excel 3.0
        { 0x008F, "LIST.NAMES" },
        { 0x0090, "FILE.CLOSE" },
        { 0x0091, "SAVE.WORKSPACE" },
        { 0x0092, "DATA.FORM" },
        { 0x0093, "COPY.CHART" },
        { 0x0094, "ON.TIME" },
        { 0x0095, "WAIT" },
        { 0x0096, "FORMAT.FONT" },
        { 0x0097, "FILL.UP" },
        { 0x0098, "FILL.LEFT" },
        { 0x0099, "DELETE.OVERLAY" },
        { 0x009A, "NOTE" }, // =NOTE?(). Excel 3.0. Not documented.
        { 0x009B, "SHORT.MENUS" },
        { 0x009F, "SET.UPDATE.STATUS" }, // Excel 3.0.
        { 0x00A1, "COLOR.PALETTE" }, // Excel 3.0.
        { 0x00A2, "DELETE.STYLE" }, // Excel 3.0.
        { 0x00A3, "WINDOW.RESTORE" }, // Excel 4.0.
        { 0x00A4, "WINDOW.MAXIMIZE" }, // Excel 4.0.
        { 0x00A6, "CHANGE.LINK" },
        { 0x00A7, "CALCULATE.DOCUMENT" },
        { 0x00A8, "ON.KEY" },
        { 0x00A9, "APP.RESTORE" },
        { 0x00AA, "APP.MOVE" },
        { 0x00AB, "APP.SIZE" },
        { 0x00AC, "APP.MINIMIZE" },
        { 0x00AD, "APP.MAXIMIZE" },
        { 0x00AE, "BRING.TO.FRONT" },
        { 0x00AF, "SEND.TO.BACK" },
        { 0x00B9, "MAIN.CHART.TYPE" },
        { 0x00BA, "OVERLAY.CHART.TYPE" },
        { 0x00BB, "SELECT.END" }, // Excel 2.0.
        { 0x00BC, "OPEN.MAIL" }, // Excel 2.2. Mac only.
        { 0x00BD, "SEND.MAIL" }, // Excel 2.2. Mac only.
        { 0x00BE, "STANDARD.FONT" }, // Excel 2.2. Mac only.
        
        // Excel 3.0
        { 0x00BF, "CONSOLIDATE" },
        { 0x00C0, "SORT.SPECIAL" }, // Unknown.
        { 0x00C1, "GALLERY.3D.AREA" },
        { 0x00C2, "GALLERY.3D.COLUMN" },
        { 0x00C3, "GALLERY.3D.LINE" },
        { 0x00C4, "GALLERY.3D.PIE" },
        { 0x00C5, "VIEW.3D" },
        { 0x00C6, "GOAL.SEEK" },
        { 0x00C7, "WORKGROUP" },
        { 0x00C8, "FILL.WORKGROUP" }, // Excel 3.0. From Excel 4.0 this is "FILL.GROUP".
        { 0x00C9, "UPDATE.LINK" },
        { 0x00CA, "PROMOTE" },
        { 0x00CB, "DEMOTE" },
        { 0x00CC, "SHOW.DETAIL" },
        { 0x00CE, "UNGROUP" },
        { 0x00CF, "PLACEMENT" }, // Excel 3.0. From Excel 4.0 this is "OBJECT.PROPERTIES".
        { 0x00D0, "SAVE.NEW.OBJECT" },
        { 0x00D1, "SHARE" },
        { 0x00D2, "SHARE.NAME" },
        { 0x00D3, "DUPLICATE" },
        { 0x00D4, "APPLY.STYLE" },
        { 0x00D5, "ASSIGN.TO.OBJECT" },
        { 0x00D6, "OBJECT.PROTECTION" },
        { 0x00D7, "HIDE.OBJECT" },
        { 0x00D8, "SET.EXTRACT" },
        { 0x00D9, "CREATE.PUBLISHER" },
        { 0x00DA, "SUBSCRIBE.TO" },
        { 0x00DB, "ATTRIBUTES" },
        { 0x00DC, "SHOW.TOOLBAR" }, // Unknown.
        { 0x00DE, "PRINT.PREVIEW" },
        { 0x00DF, "EDIT.COLOR" },
        { 0x00E0, "SHOW.LEVELS" },
        { 0x00E1, "FORMAT.MAIN" },
        { 0x00E2, "FORMAT.OVERLAY" },
        { 0x00E3, "ON.RECALC" },
        { 0x00E4, "EDIT.SERIES" },
        { 0x00E5, "DEFINE.STYLE" },
        { 0x00F0, "LINE.PRINT" }, // Unknown.
        { 0x00F3, "ENTER.DATA" }, // Unknown.
        { 0x00F9, "GALLERY.RADAR" }, // Unknown.
        { 0x00FA, "MERGE.STYLES" },
        { 0x00FB, "EDITION.OPTIONS" },
        { 0x00FC, "PASTE.PICTURE" },
        { 0x00FD, "PASTE.PICTURE.LINK" },

        // Excel 4.0
        { 0x00FE, "SPELLING" },
        { 0x0100, "ZOOM" },
        { 0x0103, "INSERT.OBJECT" },
        { 0x0104, "WINDOW.MINIMIZE" },
        { 0x0109, "SOUND.NOTE" },
        { 0x010A, "SOUND.PLAY" },
        { 0x010B, "FORMAT.SHAPE" },
        { 0x010C, "EXTEND.POLYGON" },
        { 0x010D, "FORMAT.AUTO" },
        { 0x0110, "GALLERY.3D.BAR" },
        { 0x0111, "GALLERY.3D.SURFACE" },
        { 0x0112, "FILL.AUTO" },
        { 0x0114, "CUSTOMIZE.TOOLBAR" },
        { 0x0115, "ADD.TOOL" },
        { 0x0116, "EDIT.OBJECT" },
        { 0x0117, "ON.DOUBLECLICK" },
        { 0x0118, "ON.ENTRY" },
        { 0x0119, "WORKBOOK.ADD" },
        { 0x011A, "WORKBOOK.MOVE" },
        { 0x011B, "WORKBOOK.COPY" },
        { 0x011C, "WORKBOOK.OPTIONS" },
        { 0x011D, "SAVE.WORKSPACE" },
        { 0x0120, "CHART.WIZARD" },
        { 0x0121, "DELETE.TOOL" },
        { 0x0122, "MOVE.TOOL" },
        { 0x0123, "WORKBOOK.SELECT" },
        { 0x0124, "WORKBOOK.ACTIVATE" },
        { 0x0125, "ASSIGN.TO.TOOL" },
        { 0x0127, "COPY.TOOL" },
        { 0x0128, "RESET.TOOL" },
        { 0x0129, "CONSTRAIN.NUMERIC" },
        { 0x012A, "PASTE.TOOL" },
        { 0x012E, "WORKBOOK.NEW" },

        // Excel 5.0
        { 0x0131, "SCENARIO.CELLS" },
        { 0x0132, "SCENARIO.DELETE" },
        { 0x0133, "SCENARIO.ADD" },
        { 0x0134, "SCENARIO.EDIT" },
        { 0x0135, "SCENARIO.SHOW" },
        { 0x0136, "SCENARIO.SHOW.NEXT" },
        { 0x0137, "SCENARIO.SUMMARY" },
        { 0x0138, "PIVOT.TABLE.WIZARD" },
        { 0x0139, "PIVOT.FIELD.PROPERTIES" },
        { 0x013A, "PIVOT.FIELD" },
        { 0x013B, "PIVOT.ITEM" },
        { 0x013C, "PIVOT.ADD.FIELDS" },
        { 0x013E, "OPTIONS.CALCULATION" },
        { 0x013F, "OPTIONS.EDIT" },
        { 0x0140, "OPTIONS.VIEW" },
        { 0x0141, "ADDIN.MANAGER" },
        { 0x0142, "MENU.EDITOR" },
        { 0x0143, "ATTACH.TOOLBARS" },
        { 0x0144, "VBAActivate" },
        { 0x0145, "OPTIONS.CHART" },
        { 0x0148, "VBA.INSERT.FILE" },
        { 0x014A, "VBA.PROCEDURE.DEFINITION" },
        { 0x0150, "ROUTING.SLIP" },
        { 0x0152, "ROUTE.DOCUMENT" },
        { 0x0153, "MAIL.LOGON" },
        { 0x0156, "INSERT.PICTURE" },
        { 0x0157, "EDIT.TOOL" },
        { 0x0158, "GALLERY.DOUGHNUT" },
        { 0x015E, "CHART.TREND" },
        { 0x0160, "PIVOT.ITEM.PROPERTIES" },
        { 0x0162, "WORKBOOK.INSERT" },
        { 0x0163, "OPTIONS.TRANSITION" },
        { 0x0164, "OPTIONS.GENERAL" },
        { 0x0172, "FILTER.ADVANCED" },
        { 0x0175, "MAIL.ADD.MAILER" },
        { 0x0176, "MAIL.DELETE.MAILER" },
        { 0x0177, "MAIL.REPLY" },
        { 0x0178, "MAIL.REPLY.ALL" },
        { 0x0179, "MAIL.FORWARD" },
        { 0x017A, "MAIL.NEXT.LETTER" },
        { 0x017B, "DATA.LABEL" },
        { 0x017C, "INSERT.TITLE" },
        { 0x017D, "FONT.PROPERTIES" },
        { 0x017E, "MACRO.OPTIONS" },
        { 0x017F, "WORKBOOK.HIDE" },
        { 0x0180, "WORKBOOK.UNHIDE" },
        { 0x0181, "WORKBOOK.DELETE" },
        { 0x0182, "WORKBOOK.NAME" },
        { 0x0184, "GALLERY.CUSTOM" },
        { 0x0186, "ADD.CHART.AUTOFORMAT" },
        { 0x0187, "DELETE.CHART.AUTOFORMAT" },
        { 0x0188, "CHART.ADD.DATA" },
        { 0x0189, "AUTO.OUTLINE" },
        { 0x018A, "TAB.ORDER" },
        { 0x018B, "SHOW.DIALOG" },
        { 0x018C, "SELECT.ALL" },
        { 0x018D, "UNGROUP.SHEETS" },
        { 0x018E, "SUBTOTAL.CREATE" },
        { 0x018F, "SUBTOTAL.REMOVE" },
        { 0x0190, "RENAME.OBJECT" },
        { 0x019C, "WORKBOOK.SCROLL" },
        { 0x019D, "WORKBOOK.NEXT" },
        { 0x019E, "WORKBOOK.PREV" },
        { 0x019F, "WORKBOOK.TAB.SPLIT" },
        { 0x01A0, "FULL.SCREEN" },
        { 0x01A1, "WORKBOOK.PROTECT" },
        { 0x01A4, "SCROLLBAR.PROPERTIES" },
        { 0x01A5, "PIVOT.SHOW.PAGES" },
        { 0x01A6, "TEXT.TO.COLUMNS" },
        { 0x01A7, "FORMAT.CHARTTYPE" },
        { 0x01A8, "LINK.FORMAT" },
        { 0x01A9, "TRACER.DISPLAY" },
        { 0x01AE, "TRACER.NAVIGATE" },
        { 0x01AF, "TRACER.CLEAR" },
        { 0x01B0, "TRACER.ERROR" },
        { 0x01B1, "PIVOT.FIELD.GROUP" },
        { 0x01B2, "PIVOT.FIELD.UNGROUP" },
        { 0x01B3, "CHECKBOX.PROPERTIES" },
        { 0x01B4, "LABEL.PROPERTIES" },
        { 0x01B5, "LISTBOX.PROPERTIES" },
        { 0x01B6, "EDITBOX.PROPERTIES" },
        { 0x01B7, "PIVOT.REFRESH" },
        { 0x01B8, "LINK.COMBO" },
        { 0x01B9, "OPEN.TEXT" },
        { 0x01BA, "HIDE.DIALOG" },
        { 0x01BB, "SET.DIALOG.FOCUS" },
        { 0x01BC, "ENABLE.OBJECT" },
        { 0x01BD, "PUSHBUTTON.PROPERTIES" },
        { 0x01BE, "SET.DIALOG.DEFAULT" },
        { 0x01BF, "FILTER" },
        { 0x01C0, "FILTER.SHOW.ALL" },
        { 0x01C1, "CLEAR.OUTLINE" },
        { 0x01C2, "FUNCTION.WIZARD" },
        { 0x01C3, "ADD.LIST.ITEM" },
        { 0x01C4, "SET.LIST.ITEM" },
        { 0x01C5, "REMOVE.LIST.ITEM" },
        { 0x01C6, "SELECT.LIST.ITEM" },
        { 0x01C7, "SET.CONTROL.VALUE" },
        { 0x01C8, "SAVE.COPY.AS" },
        { 0x01CA, "OPTIONS.LISTS.ADD" },
        { 0x01CB, "OPTIONS.LISTS.DELETE" },
        { 0x01CC, "SERIES.AXES" },
        { 0x01CD, "SERIES.X" },
        { 0x01CE, "SERIES.Y" },
        { 0x01CF, "ERRORBAR.X" },
        { 0x01D0, "ERRORBAR.Y" },
        { 0x01D1, "FORMAT.CHART" },
        { 0x01D2, "SERIES.ORDER" },
        { 0x01D3, "MAIL.LOGOFF" },
        { 0x01D4, "CLEAR.ROUTING.SLIP" },
        { 0x01D5, "APP.ACTIVATE.MICROSOFT" },
        { 0x01D6, "MAIL.EDIT.MAILER" },
        { 0x01D7, "ON.SHEET" },
        { 0x01D8, "STANDARD.WIDTH" },
        { 0x01D9, "SCENARIO.MERGE" },
        { 0x01DA, "SUMMARY.INFO" },
        { 0x01DB, "FIND.FILE" },
        { 0x01DC, "ACTIVE.CELL.FONT" },
        { 0x01DD, "ENABLE.TIPWIZARD" },
        { 0x01DE, "VBA.MAKE.ADDIN" },
        { 0x01E0, "INSERTDATATABLE" }, // Unknown.
        { 0x01E1, "WORKGROUP.OPTIONS" }, // Unknown.
        { 0x01E2, "MAIL.SEND.MAILER" }, // Excel 4.0.
    };

    public static string ReadFormulaString(XlsBiffRecord record, int biffVersion, int offset, int cce, int rgbExtraOffset, XlsFormulaReaderContext context)
    {
        if (cce == 0)
        {
            return string.Empty;
        }

        var formulaString = new StringBuilder();

        // rgce (variable): An Rgce that specifies the sequence of Ptgs for
        // the formula. MUST NOT contain PtgRefN, PtgAreaN, or PtgSxName.
        int read = 0;

        // Indicates if the formula is a basic assignment (e.g. A1 = 5)
        bool isBasicAssignment = false;

        // Indicates if this is a array or table formula (e.g., or{=Sheet2!$B$1:$B$3} {=TABLE(...)})
        bool isArrayFormula = false;

        // If we're in an external sheet, we need to read from the
        // external names and not the defined names.
        bool isInExternalSheet = false;

        var operands = new Stack<string>();
        while (read < cce)
        {
            var ptg = (Ptg)record.ReadByte(offset + read);
            read++;

            switch (ptg)
            {
                case Ptg.PtgExp:
                    read = ParsePtgExp(record, biffVersion, offset, read, operands, context, out isArrayFormula);
                    break;

                case Ptg.PtgTbl:
                    read = ParsePtgTbl(record, biffVersion, offset, read, operands, context);
                    isArrayFormula = true;
                    break;

                case Ptg.PtgAdd:
                    // [MS-XLS] 2.5.198.26 PtgAdd
                    // The PtgAdd structure specifies a binary-value-operator that
                    // adds the second expression in a binary-value-expression
                    // to the first.
                    ParseBinaryOperator(operands, "+");
                    break;

                case Ptg.PtgSub:
                    // [MS-XLS] 2.5.198.90 PtgSub
                    // The PtgSub structure specifies a binary-value operator that
                    // subtracts the second expression in a binary-value-expression
                    // from the first.
                    ParseBinaryOperator(operands, "-");
                    break;

                case Ptg.PtgMul:
                    // [MS-XLS] 2.5.198.75 PtgMul
                    // The PtgMul structure specifies a binary-value-operator that
                    // multiplies the first and second expressions in a
                    // binary-value-expression.
                    ParseBinaryOperator(operands, "*");
                    break;

                case Ptg.PtgDiv:
                    // [MS-XLS] 2.5.198.45 PtgDiv
                    // The PtgDiv structure specifies a binary-value-operator that
                    // divides the first expression in a binary-value-expression
                    // by the second.
                    ParseBinaryOperator(operands, "/");
                    break;

                case Ptg.PtgPower:
                    // [MS-XLS] 2.5.198.82 PtgPower
                    // The PtgPower structure specifies a binary-value-operator
                    // that raises the first expression in a binary-value-expression
                    // to the power of the second.
                    ParseBinaryOperator(operands, "^");
                    break;

                case Ptg.PtgConcat:
                    // [MS-XLS] 2.5.198.43 PtgConcat
                    // The PtgConcat structure specifies a binary-value-operator
                    // that appends the second expression in binary-value-expression
                    // to the first.
                    ParseBinaryOperator(operands, "&");
                    break;

                case Ptg.PtgLt:
                    // [MS-XLS] 2.5.198.69 PtgLt
                    // The PtgLe structure specifies a binary-value-operator that
                    // compares whether the first expression in a binary-value-expression
                    // is less than the second.
                    ParseBinaryOperator(operands, "<");
                    break;

                case Ptg.PtgLe:
                    // [MS-XLS] 2.5.198.68 PtgLe
                    // The PtgLe structure specifies a binary-value-operator that
                    // compares whether the first expression in a binary-value-expression
                    // is less than or equal to the second.
                    ParseBinaryOperator(operands, "<=");
                    break;

                case Ptg.PtgEq:
                    // [MS-XLS] 2.5.198.56 PtgEq
                    // The PtgEq structure specifies the comparison of whether the
                    // first expression is equal to the second expression.
                    ParseBinaryOperator(operands, "=");
                    break;

                case Ptg.PtgGe:
                    // [MS-XLS] 2.5.198.64 PtgGe
                    // The PtgGe structure specifies a binary-value-operator that
                    // compares whether the first expression in a binary-value-expression
                    // is greater than or equal to the second.
                    ParseBinaryOperator(operands, ">=");
                    break;

                case Ptg.PtgGt:
                    // [MS-XLS] 2.5.198.65 PtgGt
                    // The PtgGt structure specifies a binary-value-operator that
                    // compares whether the first expression in a binary-value-expression
                    // is greater than the second.
                    ParseBinaryOperator(operands, ">");
                    break;

                case Ptg.PtgNe:
                    // [MS-XLS] 2.5.198.78 PtgNe
                    // The PtgNe structure specifies a binary-value-operator that
                    // compares whether the second expression in a binary-value-expression
                    // is not equal to the first.
                    ParseBinaryOperator(operands, "<> ");
                    break;

                case Ptg.PtgUnion:
                    // [MS-XLS] 2.5.198.94 PtgUnion
                    // The PtgUnion structure specifies a binary-reference-operator
                    // that specifies a union of the first expression in a
                    // binary-reference-expression with the second.
                    ParseBinaryOperator(operands, ",");
                    break;

                case Ptg.PtgRange:
                    // [MS-XLS] 2.5.198.83 PtgRange
                    // The PtgRange structure specifies the range operation,
                    // where the minimum bounding rectangle of the first
                    // expression and the second expression is generated.
                    ParseBinaryOperator(operands, ":");
                    break;

                case Ptg.PtgIsect:
                    // [MS-XLS] 2.5.198.67 PtgIsect
                    // The PtgIsect structure specifies a binary-reference-operator that
                    // intersects the first expression in a binary-reference-expression
                    // with the second.
                    ParseBinaryOperator(operands, " ");
                    break;

                case Ptg.PtgUplus:
                    // [MS-XLS] 2.5.198.95 PtgUplus
                    // The PtgUplus structure specifies a unary-operator which leaves a
                    // unary-expression unchanged.
                    ParseUnaryOperator(operands, "+", true);
                    break;

                case Ptg.PtgUminus:
                    // [MS-XLS] 2.5.198.93 PtgUminus
                    // The PtgUminus structure specifies a unary-operator which generates
                    // the additive inverse of a unary-expression.
                    ParseUnaryOperator(operands, "-", true);
                    break;

                case Ptg.PtgPercent:
                    // [MS-XLS] 2.5.198.81 PtgPercent
                    // The PtgPercent structure specifies a unary-operator which divides
                    // the expression in a unary-expression by 100.
                    ParseUnaryOperator(operands, "%", false);
                    break;

                case Ptg.PtgParen:
                    ParsePtgParen(operands);
                    break;

                case Ptg.PtgMissArg:
                    ParsePtgMissArg(operands);
                    break;

                case Ptg.PtgStr:
                    read = ParsePtgStr(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgSheet:
                    read = ParsePtgSheet(record, biffVersion, offset, read, operands, context, ref isInExternalSheet);
                    break;

                case Ptg.PtgEndSheet:
                    read = ParsePtgEndSheet(record, biffVersion, offset, read, operands);
                    isInExternalSheet = false;
                    break;

                case Ptg.PtgErr:
                    read = ParsePtgErr(record, offset, read, operands);
                    break;

                case Ptg.PtgBool:
                    read = ParsePtgBool(record, offset, read, operands);
                    break;

                case Ptg.PtgInt:
                    read = ParsePtgInt(record, offset, read, operands);
                    break;

                case Ptg.PtgNum:
                    read = ParsePtgNum(record, offset, read, operands);
                    break;

                case Ptg.PtgArrayR:
                case Ptg.PtgArrayV:
                case Ptg.PtgArrayA:
                    read = ParsePtgArray(record, biffVersion, offset, read, ref rgbExtraOffset, operands);
                    break;

                case Ptg.PtgFuncR:
                case Ptg.PtgFuncV:
                case Ptg.PtgFuncA:
                    read = ParsePtgFunc(record, biffVersion, offset, read, operands, isBasicAssignment);
                    break;

                case Ptg.PtgFuncVarR:
                case Ptg.PtgFuncVarV:
                case Ptg.PtgFuncVarA:
                    read = ParsePtgFuncVar(record, biffVersion, offset, read, operands, isBasicAssignment);
                    break;

                case Ptg.PtgNameR:
                case Ptg.PtgNameV:
                case Ptg.PtgNameA:
                    read = ParsePtgName(record, biffVersion, offset, read, operands, context, isInExternalSheet);
                    break;

                case Ptg.PtgRefR:
                case Ptg.PtgRefV:
                case Ptg.PtgRefA:
                    read = ParsePtgRef(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgAreaR:
                case Ptg.PtgAreaV:
                case Ptg.PtgAreaA:
                    read = ParsePtgArea(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgRefErrR:
                case Ptg.PtgRefErrV:
                case Ptg.PtgRefErrA:
                    read = ParsePtgRefErr(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgMemAreaR:
                case Ptg.PtgMemAreaV:
                case Ptg.PtgMemAreaA:
                    read = ParsePtgMemArea(record, biffVersion, offset, read, ref rgbExtraOffset);
                    break;

                case Ptg.PtgMemErrR:
                case Ptg.PtgMemErrV:
                case Ptg.PtgMemErrA:
                    read = ParsePtgMemErr(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgMemNoMemR:
                case Ptg.PtgMemNoMemV:
                case Ptg.PtgMemNoMemA:
                    read = ParsePtgMemNoMem(record, biffVersion, offset, read);
                    break;

                case Ptg.PtgMemFuncR:
                case Ptg.PtgMemFuncV:
                case Ptg.PtgMemFuncA:
                    read = ParsePtgMemFunc(record, biffVersion, offset, read);
                    break;

                case Ptg.PtgAreaErrR:
                case Ptg.PtgAreaErrV:
                case Ptg.PtgAreaErrA:
                    read = ParsePtgAreaErr(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgRefNR:
                case Ptg.PtgRefNV:
                case Ptg.PtgRefNA:
                    read = ParsePtgRefN(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgAreaNR:
                case Ptg.PtgAreaNV:
                case Ptg.PtgAreaNA:
                    read = ParsePtgAreaN(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgMemAreaNR:
                case Ptg.PtgMemAreaNV:
                case Ptg.PtgMemAreaNA:
                    read = ParsePtgMemAreaN(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgMemNoMemNR:
                case Ptg.PtgMemNoMemNV:
                case Ptg.PtgMemNoMemNA:
                    read = ParsePtgMemNoMemN(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgFuncCER:
                case Ptg.PtgFuncCEV:
                case Ptg.PtgFuncCEA:
                    read = ParsePtgFuncCE(record, biffVersion, offset, read, operands);
                    break;

                case Ptg.PtgNameXR:
                case Ptg.PtgNameXV:
                case Ptg.PtgNameXA:
                    read = ParsePtgNameX(record, biffVersion, offset, read, operands, context);
                    break;

                case Ptg.PtgRef3dR:
                case Ptg.PtgRef3dV:
                case Ptg.PtgRef3dA:
                    if (biffVersion <= 5)
                    {
                        read = ParsePtgRef3d5(record, biffVersion, offset, read, operands, context, isError: false);
                    }
                    else
                    {
                        read = ParsePtgRef3d8(record, biffVersion, offset, read, operands, context);
                    }

                    break;

                case Ptg.PtgArea3dR:
                case Ptg.PtgArea3dV:
                case Ptg.PtgArea3dA:
                    if (biffVersion <= 5)
                    {
                        read = ParsePtgArea3d5(record, biffVersion, offset, read, operands, context, isError: false);
                    }
                    else
                    {
                        read = ParsePtgArea3d8(record, biffVersion, offset, read, operands, context);
                    }

                    break;

                case Ptg.PtgAreaErr3dR:
                case Ptg.PtgAreaErr3dV:
                case Ptg.PtgAreaErr3dA:
                    if (biffVersion <= 5)
                    {
                        read = ParsePtgArea3d5(record, biffVersion, offset, read, operands, context, isError: true);
                    }
                    else
                    {
                        read = ParsePtgArea3d8(record, biffVersion, offset, read, operands, context);
                    }

                    break;

                case Ptg.PtgRefErr3dR:
                case Ptg.PtgRefErr3dV:
                case Ptg.PtgRefErr3dA:
                    if (biffVersion <= 5)
                    {
                        read = ParsePtgRef3d5(record, biffVersion, offset, read, operands, context, isError: true);
                    }
                    else
                    {
                        read = ParsePtgRef3d8(record, biffVersion, offset, read, operands, context);
                    }

                    break;

                case (Ptg)0x19:
                    read = ParsePtgAttr(record, biffVersion, offset, read, operands, formulaString, ref isBasicAssignment);
                    break;

                default:
                    throw new NotSupportedException($"PTG 0x{(int)ptg:X2} not supported in formula string parsing.");
            }
        }

        if (operands.Count != 1)
        {
            throw new InvalidOperationException($"Invalid formula parsing state - final operand count is not 1. Operands: {string.Join(", ", operands)}");
        }

        // Build the final formula string.
        if (isArrayFormula)
        {
            formulaString.Append('{');
        }
        else if (!isBasicAssignment)
        {
            // Prepend the '=' sign for normal formulas.
            formulaString.Append('=');
        }

        formulaString.Append(operands.Pop());
        if (isArrayFormula)
        {
            formulaString.Append('}');
        }

        return formulaString.ToString();
    }

    private static (ushort RowIndex, bool RowRelative, bool ColumnRelative) ReadBiff2RowIndex(XlsBiffRecord record, int offset, out int bytesRead)
    {
        // In the file format versions up to BIFF5, it is possible to use 16384 (214) rows.
        // A cell address contains the row index as a 14-bit value, the column index as an
        // 8-bit value, and two flags. The flags, encoded into the row index, specify whether
        // the row or column index is absolute or relative.
        // 13-0 3FFFH Index to row (0…16383) or row offset (method [B], -8192…8191)
        // 14 4000H 0 = Absolute column index 1 = Relative column index, or column offset
        // 15 8000H 0 = Absolute row index 1 = Relative row index, or row offset
        var value = record.ReadUInt16(offset);
        bytesRead = 2;

        var rowIndex = (ushort)(value & 0x3FFF);
        var columnRelative = (value & 0x4000) != 0;
        var rowRelative = (value & 0x8000) != 0;
        return (rowIndex, rowRelative, columnRelative);
    }

    private static (int ColIndex, bool ColRelative, bool RowRelative) ReadBiff8ColIndexRel(XlsBiffRecord record, int offset, out int bytesRead)
    {
        // [MS-XLS] 2.5.50 ColRelNegU
        // The ColRelNegU structure specifies the zero-based column index of a column
        // in a sheet offset information for this column index and a corresponding row
        // index.
        // col (14 bits): A signed integer that specifies the zero-based column index
        // or offset of a column in the sheet that contains this structure.  MUST be
        // greater than or equal to -255 be less than or equal to 255.
        // A - colRelative (1 bit): A bit that specifies whether col is an offset.
        // B - rowRelative (1 bit): bit that specifies whether a row index corresponding to col in the structure containing this structure is an offset.
        var value = record.ReadInt16(offset);
        bytesRead = 2;

        var colIndex = value & 0x3FFF;
        var colRelative = (value & 0x4000) != 0;
        var rowRelative = (value & 0x8000) != 0;
        return (colIndex, colRelative, rowRelative);
    }

    private static (ushort ColIndex, bool ColRelative, bool RowRelative) ReadBiff8ColIndexU(XlsBiffRecord record, int offset, out int bytesRead)
    {
        // [MS-XLS] 2.5.51 ColRelU
        // The ColRelU structure specifies the zero-based index of a column in a sheet
        // and relative reference information for this column index and a corresponding
        // row index.
        // col (14 bits): An unsigned integer that specifies the zero-based index of a
        // column in the sheet that contains this structure.  MUST be less than or equal to 0x00FF.
        // A - colRelative (1 bit): A bit that specifies whether col is a relative reference.
        // B - rowRelative (1 bit): A bit that specifies whether a row index corresponding to
        // col in the structure containing this structure is a relative reference.
        var value = record.ReadUInt16(offset);
        bytesRead = 2;

        var colIndex = (ushort)(value & 0x3FFF);
        var colRelative = (value & 0x4000) != 0;
        var rowRelative = (value & 0x8000) != 0;
        return (colIndex, colRelative, rowRelative);
    }

    private static void PushMacroCommandCall(int biffVersion, ushort ctab, Stack<string> operands, byte cparams)
    {
        // [MS-XLS] 2.5.198.4 Cetab
        // The Cetab structure specifies a function that can be called from a formula
        // (section 2.2.2). The definition of each function specifies the function
        // name and the valid sequence of arguments.
        if (!MacroCommandNames.TryGetValue(ctab, out var macroName))
        {
            throw new NotSupportedException($"Macro command 0x{ctab:X2} not supported in formula string parsing.");
        }

        if (biffVersion >= 4)
        {
            switch (macroName)
            {
                case "FILL.WORKGROUP":
                    // FILL.WORKGROUP was renamed to FILL.GROUP in BIFF4.
                    macroName = "FILL.GROUP";
                    break;
                case "MOVE":
                    // MOVE was renamed to WINDOW.MOVE in BIFF4.
                    macroName = "WINDOW.MOVE";
                    break;
                case "PLACEMENT":
                    // PLACEMENT was renamed to OBJECT.PROPERTIES in BIFF4.
                    macroName = "OBJECT.PROPERTIES";
                    break;
            }
        }

        // This is not documented, but cparams uses the high bit to indicate
        // a confirmation dialog should be shown.
        // E.g., ALIGNMENT? is encoded with cparams = 0x81.
        var isDialogFunction = (cparams & 0x80) != 0;
        var actualMacroName = isDialogFunction ? $"{macroName}?" : macroName;
        var actualCparams = (byte)(cparams & 0x7F);
        if (operands.Count < actualCparams)
        {
            throw new InvalidOperationException($"Can't read macro command \"{actualMacroName}\" 0x{ctab:X2}. Expected {actualCparams}, but got {operands.Count}.");
        }

        var paramArgs = new List<string>();
        for (byte i = 0; i < actualCparams; i++)
        {
            paramArgs.Add(operands.Pop());
        }

        paramArgs.Reverse();
        operands.Push($"{actualMacroName}({string.Join(",", paramArgs)})");
    }

    private static void PushFunctionCall(ushort iftab, Stack<string> operands, byte? cparams, bool isBasicAssignment)
    {
        // [MS-XLS] 2.5.198.17 Ftab
        // The Ftab structure specifies a function which can be called from a formula (section 2.2.2).
        // The definition of each function specifies the function name and the valid sequence of
        // arguments.
        if (!FtabFunctionNames.TryGetValue(iftab, out var function))
        {
            throw new NotSupportedException($"Function 0x{iftab:X2} (cparams = {cparams?.ToString(CultureInfo.InvariantCulture) ?? "null"}) not supported in formula string parsing.");
        }

        var isUserDefinedFunction = iftab == 0xFF;
        int actualParams = cparams ?? function.Item2 ?? throw new NotSupportedException($"Function 0x{iftab:X2} requires parameter count in formula string parsing.");
        if (isUserDefinedFunction)
        {
            // User-defined functions use the first operand as the function name.
            actualParams -= 1;
        }
        else if (isBasicAssignment && iftab == 0x0058)
        {
            // Translate SET.NAME(A, B) to <A> = <B>
            var value = operands.Pop();
            var fullName = operands.Pop();

            // Need to remove the surrounding quotes from the name.
            var name = fullName[1..^1];
            operands.Push($"{name}={value}");
            
            return;
        }

        if (actualParams > operands.Count)
        {
            throw new InvalidOperationException($"Can't read function \"{function.Item1}\" 0x{iftab:X2}. Expected {actualParams}, but got {operands.Count}.");
        }

        List<string> args = [];
        for (int i = 0; i < actualParams; i++)
        {
            args.Add(operands.Pop());
        }

        args.Reverse();

        // User-defined functions use the first operand as the function name.
        string functionName = isUserDefinedFunction ? operands.Pop() : function.Item1;
        operands.Push($"{functionName}({string.Join(",", args)})");
    }

    private static string GetAddressString(int columnIndex, bool columnRelative, int rowIndex, bool rowRelative)
    {
        // Convert zero-based column and row indices to their Excel cell address
        // representation (A1, $A$1, A$1, $A1).
        var columnPart = GetColumnString(columnIndex, columnRelative);
        var rowPart = GetRowString(rowIndex, rowRelative);
        return $"{columnPart}{rowPart}";
    }

    private static string GetRangeString(int biffVersion, int firstColumnIndex, bool firstColumnRelative, int firstRowIndex, bool firstRowRelative, int lastColumnIndex, bool lastColumnRelative, int lastRowIndex, bool lastRowRelative)
    {
        if (!firstRowRelative && !lastRowRelative)
        {
            if (firstRowIndex == 0 && lastRowIndex == 16383)
            {
                // This is a column range, e.g., A:B
                var firstColumnString = GetColumnString(firstColumnIndex, firstColumnRelative);
                var lastColumnString = GetColumnString(lastColumnIndex, lastColumnRelative);
                return $"{firstColumnString}:{lastColumnString}";
            }
        }

        if (!firstColumnRelative && !lastColumnRelative)
        {
            if (biffVersion <= 5 && firstColumnIndex == 0 && lastColumnIndex == 255)
            {
                // This is a row range, e.g., 1:16384
                var firstRowString = GetRowString(firstRowIndex, firstRowRelative);
                var lastRowString = GetRowString(lastRowIndex, lastRowRelative);
                return $"{firstRowString}:{lastRowString}";
            }
        }

        // This is a full range, e.g., AR:B2
        var firstCell = GetAddressString(firstColumnIndex, firstColumnRelative, firstRowIndex, firstRowRelative);
        var lastCell = GetAddressString(lastColumnIndex, lastColumnRelative, lastRowIndex, lastRowRelative);
        return $"{firstCell}:{lastCell}";
    }

    private static string GetColumnString(int columnIndex, bool columnRelative)
    {
        // Convert a zero-based column index to its Excel column string representation
        // (A, AB).
        var columnName = new StringBuilder();
        int dividend = columnIndex + 1;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName.Insert(0, Convert.ToChar(65 + modulo));
            dividend = (dividend - modulo) / 26;
        }

        if (!columnRelative)
        {
            columnName.Insert(0, '$');
        }

        return columnName.ToString();
    }

    private static string GetRowString(int rowIndex, bool rowRelative)
    {
        var rowString = (rowIndex + 1).ToString(CultureInfo.InvariantCulture);
        if (!rowRelative)
        {
            rowString = $"${rowString}";
        }

        return rowString;
    }

    private static string GetErrorCodeString(Berr berr)
    {
        return berr switch
        {
            Berr.NULL => "#NULL!",
            Berr.DIV0 => "#DIV/0!",
            Berr.VALUE => "#VALUE!",
            Berr.REF => "#REF!",
            Berr.NAME => "#NAME?",
            Berr.NUM => "#NUM!",
            Berr.NA => "#N/A",
            _ => "#UNKNOWN!"
        };
    }

    // Helper methods for parsing each Ptg type
    private static void ParseBinaryOperator(Stack<string> operands, string operatorSymbol)
    {
        var right = operands.Pop();
        var left = operands.Pop();
        operands.Push($"{left}{operatorSymbol}{right}");
    }

    private static void ParseUnaryOperator(Stack<string> operands, string operatorSymbol, bool prefix)
    {
        var expr = operands.Pop();
        operands.Push(prefix ? $"{operatorSymbol}{expr}" : $"{expr}{operatorSymbol}");
    }

    private static void ParsePtgParen(Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.80 PtgParen
        // The PtgParen display token specifies that parentheses are displayed
        // around the expression in a display-precedence-expression.
        var expr = operands.Pop();
        operands.Push($"({expr})");
    }

    private static void ParsePtgMissArg(Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.74 PtgMissArg
        // The PtgMissArg operand specifies a missing value.
        operands.Push(string.Empty);
    }

    private static int ParsePtgExp(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context, out bool isArrayFormula)
    {
        // [MS-XLS] 2.5.198.58 PtgExp
        // The PtgExp structure specifies that the containing Rgce is part
        // of an array formula (section 2.2.2) or shared formula and specifies
        // the row and column of the cell in which that formula exists.
        // The row and col fields of this structure specify a cell on the current sheet.
        // There MUST be a Formula record where the cell.rw field of that record is equal
        // to row, and cell.col.col field of that record is equal to col.
        // That Formula record MUST be followed by either a ShrFmla record or an Array record.
        // If that Formula record is followed by a ShrFmla, the row field of this structure
        // MUST be greater than or equal to the ref.rwFirst field and less than or equal
        // to the ref.rwLast field of the ShrFmla record, and the col field of this
        // structure MUST be greater than or equal to the ref.colFirst field and less
        // than or equal to the ref.colLast field of the ShrFmla record.
        // If that Formula record is followed by an Array, the row field of this structure
        // MUST be equal to the ref.rwFirst field of the Array record, and the col field
        // of this structure MUST be equal to the ref.colFirst field of the Array record.
        ushort row;
        ushort col;
        if (biffVersion == 2)
        {
            // In BIFF2, PtgExp uses a different format for row and column.
            row = record.ReadUInt16(offset + read);
            read += 2;

            col = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // row (2 bytes): A Rw that specifies the row of the cell that contains the
            // array formula or shared formula that the containing Rgce is a part of.
            row = record.ReadUInt16(offset + read);
            read += 2;

            // col (2 bytes): A Col that specifies the column of the cell that contains the
            // array formula or shared formula that the containing Rgce is a part of.
            col = record.ReadUInt16(offset + read);
            read += 2;
        }

        // First look for an array formula.
        XlsBiffArray foundArray = null;
        foreach (var arr in context.Arrays)
        {
            if (arr.RowFirst == row && arr.ColFirst == col)
            {
                foundArray = arr;
                break;
            }
        }

        if (foundArray != null)
        {
            var arrayFormulaString = foundArray.GetFormulaString(context);
            operands.Push(arrayFormulaString);
            isArrayFormula = true;
            return read;
        }

        // Next look for a shared formula.
        XlsBiffSharedFormula foundSharedFormula = null;
        foreach (var sf in context.SharedFormulas)
        {
            if (row >= sf.RowFirst && row <= sf.RowLast &&
                col >= sf.ColFirst && col <= sf.ColLast)
            {
                foundSharedFormula = sf;
                break;
            }
        }

        if (foundSharedFormula != null)
        {
            var sharedFormulaString = foundSharedFormula.GetFormulaString(context);
            operands.Push(sharedFormulaString);
            isArrayFormula = false;
            return read;
        }

        // If we reach here, we could not find either an array or shared formula.
        throw new NotSupportedException($"PTG EXP could not find matching array or shared formula at row {row}, col {col}.");
    }

    private static int ParsePtgTbl(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context)
    {
        // [MS-XLS] 2.5.198.92 PtgTbl
        // The PtgTbl structure specifies that the Rgce that contains this PtgTbl is part of
        // a data table (1) or an ObjectParsedFormula.
        // If the Rgce is not part of an ObjectParsedFormula, then there MUST be a Table
        // record in the current part where the ref.rwFirst field in Table is equal to row
        // and the ref.colFirst field in Table is equal to col.
        var row = record.ReadUInt16(offset + read);
        read += 2;

        int col;
        if (biffVersion == 2)
        {
            // In BIFF2, the column is stored as a single byte.
            col = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            col = record.ReadUInt16(offset + read);
            read += 2;
        }

        XlsBiffDataTable foundTable = null;
        foreach (var dt in context.DataTables)
        {
            if (dt.RowFirst == row && dt.ColFirst == col)
            {
                foundTable = dt;
                break;
            }
        }

        if (foundTable == null)
        {
            throw new NotSupportedException($"PTG TBL could not find matching data table at row {row}, col {col}.");
        }

        var sb = new StringBuilder();
        sb.Append("=TABLE(");
        sb.Append(GetAddressString(foundTable.ColInputFirst, true, foundTable.RowInputFirst, true));
        sb.Append(',');
        if (foundTable.ColInputSecond.HasValue && foundTable.RowInputSecond.HasValue)
        {
            sb.Append(GetAddressString(foundTable.ColInputSecond.Value, true, foundTable.RowInputSecond.Value, true));
        }

        sb.Append(')');
        operands.Push(sb.ToString());
        return read;
    }

    private static int ParsePtgStr(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.89 PtgStr
        // The PtgStr operand specifies a Unicode string value.
        string stringValue;
        if (biffVersion <= 5)
        {
            // In BIFF2-BIFF5, strings are stored as single-byte character strings.
            stringValue = record.ReadByteString(offset + read, out int bytesRead);
            read += bytesRead;
        }
        else
        {
            // [MS-XLS] 2.5.198.89 PtgStr
            // The PtgStr operand specifies a Unicode string value.
            // string (variable): A ShortXLUnicodeString value that specifies
            // the string.
            stringValue = record.ReadShortXLUnicodeString(offset + read, out int bytesRead);
            read += bytesRead;
        }

        operands.Push($"\"{stringValue}\"");
        return read;
    }

    private static int ParsePtgErr(XlsBiffRecord record, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.57 PtgErr
        // The PtgErr operand specifies an error code.
        // err (1 byte): A BErr that specifies the error code.
        var errValue = (Berr)record.ReadByte(offset + read);
        read += 1;
        operands.Push(GetErrorCodeString(errValue));
        return read;
    }

    private static int ParsePtgBool(XlsBiffRecord record, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.42 PtgBool
        // The PtgBool operand specifies a Boolean value.
        // boolean (1 byte):  A Boolean (section 2.5.14) that specifies the value.
        var boolValue = record.ReadByte(offset + read) != 0;
        read += 1;
        operands.Push(boolValue ? "TRUE" : "FALSE");
        return read;
    }

    private static int ParsePtgInt(XlsBiffRecord record, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.66 PtgInt
        // The PtgInt operand specifies an unsigned integer value.
        // integer (2 bytes): An unsigned integer that specifies the value.
        var intValue = record.ReadUInt16(offset + read);
        read += 2;
        operands.Push(intValue.ToString(CultureInfo.InvariantCulture));
        return read;
    }

    private static int ParsePtgNum(XlsBiffRecord record, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.79 PtgNum
        // The PtgNum operand specifies a IEEE 754 floating-point number.
        // num (8 bytes): An Xnum (section 2.5.342) that specifies the value.
        var numValue = record.ReadDouble(offset + read);
        read += 8;
        operands.Push(numValue.ToString(CultureInfo.InvariantCulture));
        return read;
    }

    private static int ParsePtgRef(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.84 PtgRef
        // The PtgRef operand specifies a reference to a single
        // cell as an RgceLoc.
        operands.Push(ParseCellAddress(record, biffVersion, offset + read, out int bytesRead));
        read += bytesRead;
        return read;
    }

    private static int ParsePtgArea(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.27 PtgArea
        // The PtgArea operand specifies a reference to a rectangular range of cells.
        operands.Push(ParseCellAddressRange(record, biffVersion, offset + read, out int bytesRead));
        read += bytesRead;
        return read;
    }

    private static int ParsePtgRefN(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.88 PtgRefN
        // The PtgRefN operand specifies a reference to a single cell as an RgceLocRel.
        operands.Push(ParseCellAddress(record, biffVersion, offset + read, out int bytesRead));
        read += bytesRead;
        return read;
    }

    private static int ParsePtgAreaN(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.31 PtgAreaN
        // The PtgAreaN operand specifies a reference to a rectangular range of cells 
        // as an RgceAreaRel.
        operands.Push(ParseCellAddressRange(record, biffVersion, offset + read, out int bytesRead));
        read += bytesRead;
        return read;
    }

    private static int ParsePtgMemAreaN(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // Not documented in MS-XLS.
        // But documented in https://github.com/xieguigang/sciBASIC/blob/master/mime/application%25vnd.openxmlformats-officedocument.spreadsheetml.sheet/Excel/XLS/BIFF/excel.txt#L633
        // Also documented in https://www.openoffice.org/sc/excelfileformat.pdf#page=26&zoom=100,114,662
        if (biffVersion == 2)
        {
            // BIFF2 PtgMemAreaN structure is different.
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgMemNoMemN(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // Not documented in MS-XLS.
        // But documented in https://github.com/xieguigang/sciBASIC/blob/master/mime/application%25vnd.openxmlformats-officedocument.spreadsheetml.sheet/Excel/XLS/BIFF/excel.txt#L633
        // Also documented in https://www.openoffice.org/sc/excelfileformat.pdf#page=26&zoom=100,114,662
        if (biffVersion == 2)
        {
            // BIFF2 ParsePtgMemNoMemN structure is different.
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgArray(XlsBiffRecord record, int biffVersion, int offset, int read, ref int rgbExtraOffset, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.27 PtgArray
        // The PtgArray operand specifies an array of values. There MUST be a PtgExtraArray
        // in the RgbExtra corresponding to this PtgArray. The correspondence between
        // PtgArray and PtgExtraArray structures is specified in RgbExtra.
        if (biffVersion == 2)
        {
            // BIFF2 PtgArray structure is different.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadUInt32(offset + read);
            read += 4;
        }
        else
        {
            // unused1 (1 byte): Undefined and MUST be ignored.
            _ = record.ReadByte(offset + read);
            read++;

            // unused2 (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            // unused3 (4 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt32(offset + read);
            read += 4;
        }

        operands.Push(ReadPtgExtraArray(record, biffVersion, ref rgbExtraOffset));
        return read;
    }

    private static string ReadPtgExtraArray(XlsBiffRecord record, int biffVersion, ref int rgbExtraOffset)
    {
        var stringBuilder = new StringBuilder();
        stringBuilder.Append('{');

        // [MS-XLS] 2.5.198.59 PtgExtraArray
        // The PtgExtraArray structure specifies the values for the corresponding
        // PtgArray as specified in RgbExtra.
        int cCols;
        int cRows;
        if (biffVersion <= 5)
        {
            // In BIFF2-5, the PtgExtraArray structure is different.
            // A column of zero indicates 256 columns.
            var cols = record.ReadByte(rgbExtraOffset);
            rgbExtraOffset += 1;
            cCols = cols == 0 ? 256 : cols;

            cRows = record.ReadUInt16(rgbExtraOffset);
            rgbExtraOffset += 2;
        }
        else
        {
            // cols (1 byte): A DColByteU that specifies one less than the number of
            // columns in the array.
            var cols = record.ReadByte(rgbExtraOffset);
            cCols = (int)cols + 1;
            rgbExtraOffset += 1;

            // rows (2 bytes):  A DRw that specifies one less than the number of rows
            // in the array.
            var rows = record.ReadUInt16(rgbExtraOffset);
            rgbExtraOffset += 2;
            cRows = rows + 1;
        }

        for (int r = 0; r < cRows; r++)
        {
            for (int c = 0; c < cCols; c++)
            {
                byte type = record.ReadByte(rgbExtraOffset);
                rgbExtraOffset += 1;

                switch (type)
                {
                    case 0x00:
                        {
                            // [MS-XLS] 2.5.198.115 SerNil
                            // unused1 (4 bytes): Undefined and MUST be ignored.
                            _ = record.ReadUInt32(rgbExtraOffset);
                            rgbExtraOffset += 4;

                            // unused2 (4 bytes): Undefined and MUST be ignored.
                            _ = record.ReadUInt32(rgbExtraOffset);
                            rgbExtraOffset += 4;
                            break;
                        }

                    case 0x01:
                        {
                            // [MS-XLS] 2.5.198.116 SerNum
                            // xnum (8 bytes): An Xnum (section 2.5.342) that
                            // specifies the value.
                            var numValue = record.ReadDouble(rgbExtraOffset);
                            rgbExtraOffset += 8;

                            stringBuilder.Append(numValue.ToString(CultureInfo.InvariantCulture));
                            break;
                        }

                    case 0x02:
                        {
                            // [MS-XLS] 2.5.198.117 SerStr
                            // The SerStr structure specifies a string in an array of values.
                            string stringValue;
                            if (biffVersion <= 5)
                            {
                                // In BIFF2-5, strings are stored as single-byte character strings.
                                stringValue = record.ReadByteString(rgbExtraOffset, out int bytesRead);
                                rgbExtraOffset += bytesRead;
                            }
                            else
                            {
                                // string (variable): An XLUnicodeString that specifies
                                // the string. The length of the string MUST be less than
                                // 256 characters.
                                stringValue = record.ReadXLUnicodeString(rgbExtraOffset, out int bytesRead);
                                rgbExtraOffset += bytesRead;
                            }

                            stringBuilder.Append('"');
                            stringBuilder.Append(stringValue);
                            stringBuilder.Append('"');
                            break;
                        }

                    case 0x04:
                        {
                            // [MS-XLS] 2.5.198.113 SerBool
                            // The SerBool structure specifies a Boolean (section 2.5.14)
                            // value in an array of values.

                            // f (1 byte):  A Boolean that specifies the value.
                            var boolValue = record.ReadByte(rgbExtraOffset) != 0;
                            rgbExtraOffset += 1;

                            // reserved2 (1 byte):  MUST be zero, and MUST be ignored.
                            _ = record.ReadByte(rgbExtraOffset);
                            rgbExtraOffset += 1;

                            // reserved3 (2 bytes):  MUST be zero, and MUST be ignored.
                            _ = record.ReadUInt16(rgbExtraOffset);
                            rgbExtraOffset += 2;

                            // reserved4 (4 bytes):  MUST be zero, and MUST be ignored.
                            _ = record.ReadUInt32(rgbExtraOffset);
                            rgbExtraOffset += 4;

                            stringBuilder.Append(boolValue ? "TRUE" : "FALSE");
                            break;
                        }

                    case 0x10:
                        {
                            // [MS-XLS] 2.5.198.114 SerErr
                            // The SerErr structure specifies an error value in an array of values.

                            // err (1 byte): A BErr that specifies the error value.
                            var errValue = (Berr)record.ReadByte(rgbExtraOffset);
                            rgbExtraOffset += 1;

                            // reserved2 (1 byte):  MUST be zero, and MUST be ignored.
                            _ = record.ReadByte(rgbExtraOffset);
                            rgbExtraOffset += 1;

                            // reserved3 (2 bytes):  MUST be zero, and MUST be ignored.
                            _ = record.ReadUInt16(rgbExtraOffset);
                            rgbExtraOffset += 2;

                            // reserved4 (4 bytes):  MUST be zero, and MUST be ignored.
                            _ = record.ReadUInt32(rgbExtraOffset);
                            rgbExtraOffset += 4;

                            stringBuilder.Append(GetErrorCodeString(errValue));
                            break;
                        }

                    default:
                        throw new NotSupportedException($"PTG Extra Array SerAr type 0x{type:X2} not supported in formula string parsing.");
                }

                if (c < cCols - 1)
                {
                    stringBuilder.Append(',');
                }
            }

            if (r < cRows - 1)
            {
                stringBuilder.Append(';');
            }
        }

        stringBuilder.Append('}');
        return stringBuilder.ToString();
    }

    private static int ParsePtgFunc(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, bool isBasicAssignment)
    {
        // [MS-XLS] 2.5.198.62 PtgFunc
        // The PtgFunc structure specifies a call to a function with a
        // fixed number of parameters, as defined in function-call.
        ushort iftab;
        if (biffVersion <= 3)
        {
            // In BIFF2 and BIFF3, iftab is a single byte.
            iftab = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // In BIFF4 and later, iftab is two bytes.
            // iftab (2 bytes): A Ftab that specifies the function to be called.
            // MUST specify a function with a fixed number of parameters.
            iftab = record.ReadUInt16(offset + read);
            read += 2;
        }

        PushFunctionCall(iftab, operands, null, isBasicAssignment);
        return read;
    }

    private static int ParsePtgFuncVar(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, bool isBasicAssignment)
    {
        // [MS-XLS] 2.5.198.63 PtgFuncVar
        // The PtgFuncVar structure specifies a call to a function with a
        // variable number of parameters as defined in function-call.
        byte cparams;
        ushort iftab;
        if (biffVersion <= 3)
        {
            // In BIFF2 and BIFF3, iftab is a single byte.
            cparams = record.ReadByte(offset + read);
            read += 1;

            iftab = record.ReadByte(offset + read);
            read += 1;

            PushFunctionCall(iftab, operands, cparams, isBasicAssignment);
        }
        else
        {
            // ptg (5 bits): Reserved. MUST be 0x02
            // A - type (2 bits): A PtgDataType that specifies the data
            // type for the value of this Ptg.
            // B - reserved (1 bit): MUST be 0, MUST be ignored.
            // cparams (1 byte): An unsigned integer that specifies the
            // number of parameters. MUST be within the range defined
            // for the function specified by tab.
            cparams = record.ReadByte(offset + read);
            read++;

            // tab (15 bits): A structure that specifies the function to
            // be called. If fCeFunc is 1, then this field specifies a
            // Cetab value. If fCeFunc is 0, then this field specifies
            // a Ftab value.
            // C - fCeFunc (1 bit): A bit that specifies whether tab
            // specifies a Cetab value or a Ftab value.
            var tabValue = record.ReadUInt16(offset + read);
            var fCeFunc = (tabValue & 0x8000) != 0;
            var tab = (ushort)(tabValue & 0x7FFF);
            read += 2;

            if (fCeFunc)
            {
                PushMacroCommandCall(biffVersion, tab, operands, cparams);
            }
            else
            {
                PushFunctionCall(tab, operands, cparams, isBasicAssignment);
            }
        }

        return read;
    }

    private static int ParsePtgName(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context, bool isInExternalSheet)
    {
        // [MS-XLS] 2.5.198.76 PtgName
        // The PtgName operand specifies a reference to a defined name in the same workbook
        // as the containing Rgce.
        // If the formula (section 2.2.2) containing this structure is part of a revision
        // as specified in the Formulas overview, then there MUST be a RevNameTabid in the
        // RgbExtra corresponding to this PtgName, which specifies those defined name.
        uint nameIndex;
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgName differently.
            nameIndex = record.ReadUInt16(offset + read);
            read += 2;

            // unused (5 bytes): Undefined and MUST be ignored.
            read += 5;
        }
        else if (biffVersion <= 4)
        {
            // BIFF3 and BIFF4 store PtgName differently.
            nameIndex = record.ReadUInt16(offset + read);
            read += 2;

            // unused (8 bytes): Undefined and MUST be ignored.
            read += 8;
        }
        else if (biffVersion == 5)
        {
            // BIFF5 stores PtgName differently.
            nameIndex = record.ReadUInt16(offset + read);
            read += 2;

            // unused (12 bytes): Undefined and MUST be ignored.
            read += 12;
        }
        else
        {
            // nameindex (4 bytes): If the formula containing this structure
            // is part of a revision as specified in the Formulas overview,
            // then this value is undefined and MUST be ignored. Otherwise
            // it is an unsigned integer that specifies a one-based index of
            // a Lbl record in the collection of Lbl records in the Globals
            // Substream. The referenced Lbl specifies the referenced defined name.
            // MUST be greater than 0 and less than or equal to the number of Lbl
            // records in the workbook.
            nameIndex = record.ReadUInt32(offset + read);
            read += 4;
        }

        if (!isInExternalSheet)
        {
            operands.Push(context.GetDefindName(nameIndex));
        }
        else
        {
            operands.Push(context.GetExternalName(nameIndex));
        }

        return read;
    }

    private static int ParsePtgSheet(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context, ref bool isInExternalSheet)
    {
        // Not documented in [MS-XLS].
        // But included in 1988 Microsoft Excel documentation
        // https://github.com/xieguigang/sciBASIC/blob/master/mime/application%25vnd.openxmlformats-officedocument.spreadsheetml.sheet/Excel/XLS/BIFF/excel.txt
        int ixals;
        if (biffVersion == 2)
        {
            // BIFF2 PtgSheet structure is different.
            _ = record.ReadUInt32(offset + read);
            read += 4;

            ixals = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            _ = record.ReadUInt32(offset + read);
            read += 4;

            _ = record.ReadUInt16(offset + read);
            read += 2;

            ixals = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        if (context.ExternalSheets == null || ixals == 0 || ixals > context.ExternalSheets.Count)
        {
            throw new InvalidOperationException("External sheets information is missing or invalid for formula string parsing.");
        }

        XlsBiffExternalSheet sheet = context.ExternalSheets[ixals - 1];
        isInExternalSheet = !sheet.IsSelf;
        operands.Push(sheet.Name);
        return read;
    }

    private static int ParsePtgEndSheet(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // Not documented in [MS-XLS].
        // But included in 1988 Microsoft Excel documentation
        // https://github.com/xieguigang/sciBASIC/blob/master/mime/application%25vnd.openxmlformats-officedocument.spreadsheetml.sheet/Excel/XLS/BIFF/excel.txt
        read += biffVersion == 2 ? 3 : 4;
        var cellRef = operands.Pop();
        var sheetRef = operands.Pop();
        operands.Push($"{sheetRef}!{cellRef}");
        return read;
    }

    private static int ParsePtgRefErr(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.86 PtgRefErr
        // The PtgRefErr operand specifies an erroneous reference to a single cell.
        if (biffVersion <= 5)
        {
            // BIFF2-BIFF5 stores PtgRefErr differently.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // unused1 (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            // unused2 (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        operands.Push("#REF!");
        return read;
    }

    private static int ParsePtgMemArea(XlsBiffRecord record, int biffVersion, int offset, int read, ref int rgbExtraOffset)
    {
        // [MS-XLS] 2.5.198.70 PtgMemArea
        // The PtgMemArea mem token specifies the cached result of a
        // binary-reference-expression in a mem-area-expression.
        if (biffVersion == 2)
        {
            _ = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadByte(offset + read);
            read += 1;

            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // unused (4 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt32(offset + read);
            read += 4;

            // cce (2 bytes): An unsigned integer that specifies the count of bytes in the
            // binary-reference-expression following this structure.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        ReadPtgExtraMem(record, biffVersion, ref rgbExtraOffset);
        return read;
    }

    private static void ReadPtgExtraMem(XlsBiffRecord record, int biffVersion, ref int rgbExtraOffset)
    {
        // [MS-XLS] 2.5.198.61 PtgExtraMem
        // The PtgExtraMem structure specifies a range that corresponds to a PtgMemArea
        // as specified in RgbExtra.

        // count (2 bytes): An unsigned integer that specifies the areas within the range.
        var count = record.ReadUInt16(rgbExtraOffset);
        rgbExtraOffset += 2;

        // array (variable): An array of Ref8U that specifies the range. The number of
        // elements MUST be equal to count.
        for (int i = 0; i < count; i++)
        {
            // [MS-XLS] 2.5.209 Ref8U
            // rwFirst (2 bytes): A RwU structure that specifies the zero-based index of the
            // first row in the range. The value MUST be less than or equal to rwLast.
            _ = record.ReadUInt16(rgbExtraOffset);
            rgbExtraOffset += 2;

            // rwLast (2 bytes): A RwU structure that specifies the zero-based index of
            // the last row in the range. The value MUST be greater than or equal to rwFirst.
            _ = record.ReadUInt16(rgbExtraOffset);
            rgbExtraOffset += 2;

            // colFirst (2 bytes): A ColU structure that specifies the zero-based index of
            // the first column in the range. The value MUST be less than or equal to colLast,
            // and MUST be less than or equal to 0x00FF.
            if (biffVersion <= 5)
            {
                // BIFF2-5 ColU is stored as a single byte.
                _ = record.ReadByte(rgbExtraOffset);
                rgbExtraOffset += 1;
            }
            else
            {
                _ = record.ReadUInt16(rgbExtraOffset);
                rgbExtraOffset += 2;
            }

            // colLast (2 bytes): A ColU structure that specifies the zero-based index of
            // the last column in the range. The value MUST be greater than or equal to
            // colFirst, and MUST be less than or equal to 0x00FF.
            if (biffVersion <= 5)
            {
                // BIFF2-5 ColU is stored as a single byte.
                _ = record.ReadByte(rgbExtraOffset);
                rgbExtraOffset += 1;
            }
            else
            {
                _ = record.ReadUInt16(rgbExtraOffset);
                rgbExtraOffset += 2;
            }
        }
    }

    private static int ParsePtgMemErr(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.71 PtgMemErr
        // The PtgMemErr mem token specifies that the result of a binary-reference-expression
        // in a mem-area-expression is an error code.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgMemErr differently.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadByte(offset + read);
            read += 1;

            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // err (1 byte): A BErr that specifies the error code value.
            _ = record.ReadByte(offset + read);
            read += 1;

            // unused1 (1 byte): Undefined and MUST be ignored.
            _ = record.ReadByte(offset + read);
            read += 1;

            // unused2 (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            // cce (2 bytes): An unsigned integer that specifies the count of bytes in the
            // binary-reference-expression following this structure.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgMemNoMem(XlsBiffRecord record, int biffVersion, int offset, int read)
    {
        // [MS-XLS] 2.5.198.73 PtgMemNoMem
        // The PtgMemNoMem mem token specifies that the result of the binary-reference-expression
        // in a mem-area-expression failed to cache.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgMemNoMem differently.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadByte(offset + read);
            read += 1;

            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // unused (4 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt32(offset + read);
            read += 4;
        }
        
        return read;
    }

    private static int ParsePtgMemFunc(XlsBiffRecord record, int biffVersion, int offset, int read)
    {
        // [MS-XLS] 2.5.198.72 PtgMemFunc
        // The PtgMemFunc mem token specifies a result of an expression of type
        // binary-reference-expression in a function-expression.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgMemFunc differently.
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // cce (2 bytes): An unsigned integer that specifies the count of bytes in the
            // binary-reference-expression following this structure.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgAreaErr(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.29 PtgAreaErr
        // The PtgAreaErr operand specifies an erroneous reference to
        // a rectangular range of cells.
        if (biffVersion <= 5)
        {
            _ = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadUInt16(offset + read);
            read += 2;

            _ = record.ReadUInt16(offset + read);
            read += 2;
        }
        else
        {
            // unused1 (2 bytes): Undefined and MUST be ignored.        
            _ = record.ReadUInt16(offset + read);
            read += 2;

            // unused2 (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            // unused3 (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;

            // unused4 (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        operands.Push("#REF!");
        return read;
    }

    private static int ParsePtgFuncCE(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // Not documented in [MS-XLS].
        // But included in 1988 Microsoft Excel documentation
        // https://github.com/xieguigang/sciBASIC/blob/master/mime/application%25vnd.openxmlformats-officedocument.spreadsheetml.sheet/Excel/XLS/BIFF/excel.txt#L2388
        byte cparams = record.ReadByte(offset + read);
        read++;

        byte index = record.ReadByte(offset + read);
        read++;

        PushMacroCommandCall(biffVersion, index, operands, cparams);
        return read;
    }

    private static int ParsePtgNameX(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context)
    {
        // [MS-XLS] 2.5.198.77 PtgNameX
        // The PtgNameX structure specifies a reference to a defined
        // name in an external workbook.
        // If the formula (section 2.2.2) containing this structure
        // is part of a revision as specified in the Formulas overview,
        // then there MUST be a RevName in the RgbExtra corresponding
        // to this PtgNameX that specifies the defined name.
        // If the formula containing this structure is not part of a
        // revision as specified in the Formulas overview (section 2.2.2),
        // then the referenced defined name is specified by an XtiIndex.]
        bool isDefinedName;
        uint nameIndex;
        if (biffVersion <= 5)
        {
            // This is always a negative value to indicate an internal name.
            // The absolute value is the onebased index to EXTERNSHEET record
            // (➜5.41) in the Local Link Table (➜4.10.2).
            int sheetIndex = record.ReadInt16(offset + read);
            read += 2;

            // unused (8 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt32(offset + read);
            read += 4;

            _ = record.ReadUInt32(offset + read);
            read += 4;

            // One-based index to DEFINEDNAME record (➜5.33) in the Global Link Table (➜4.10.2)
            // One-based index to EXTERNALNAME record (➜5.39)
            nameIndex = record.ReadUInt16(offset + read);
            read += 2;

            // unused (12 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt32(offset + read);
            read += 4;

            _ = record.ReadUInt32(offset + read);
            read += 4;

            _ = record.ReadUInt32(offset + read);
            read += 4;

            XlsBiffExternalSheet sheet = context.GetExternalSheet5(sheetIndex);
            isDefinedName = sheet.IsSelf;
        }
        else
        {
            // ixti (2 bytes): If the formula containing this structure is
            // not part of a revision as specified in the Formulas overview,
            // this value is an XtiIndex that specifies the XTI that specifies
            // the referenced defined name.
            // If the formula containing this structure is part of a revision
            // as specified in the Formulas overview, this value is undefined
            // and MUST be ignored.
            int sheetIndex = record.ReadUInt16(offset + read);
            read += 2;

            // nameindex (4 bytes): If the formula containing this structure
            // is not part of a revision as specified in the Formulas
            // overview, this value is an unsigned integer that specifies
            // the one-based index of an ExternName record in the collection
            // of ExternName records directly following the SupBook record
            // referenced by ixti. The referenced ExternName and its associated
            // records specify the referenced defined name.
            nameIndex = record.ReadUInt32(offset + read);
            read += 4;

            (XlsBiffSupBook workbook, XlsBiffXti _) = context.GetExternalSheet8(sheetIndex);
            isDefinedName = workbook.Type == XlsBiffSupBookType.Internal3DReference;
        }

        if (isDefinedName)
        {
            operands.Push(context.GetDefindName(nameIndex));
        }
        else
        {
            operands.Push(context.GetExternalName(nameIndex));
        }

        return read;
    }

    private static int ParsePtgRef3d5(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context, bool isError)
    {
        // Token tRef3d for 3D references, BIFF5:
        // This is always a negative value to indicate a 3D reference. The absolute
        // value is the onebased index to EXTERNSHEET record (➜5.41) in the Local
        // Link Table (➜4.10.2) containing the name of the first referenced sheet.
        // This is always a positive value to indicate an external reference. One-based
        // index to EXTERNSHEET record (➜5.41) in the Local Link Table (➜4.10.2).
        int sheetIndex = record.ReadInt16(offset + read);
        read += 2;

        // Not used.
        _ = record.ReadUInt32(offset + read);
        read += 4;

        _ = record.ReadUInt32(offset + read);
        read += 4;

        int? firstReferencedSheetIndex = null;
        int? lastReferencedSheetIndex = null;
        if (sheetIndex < 0)
        {
            // Zero-based index to first referenced sheet (FFFFH = deleted sheet)
            firstReferencedSheetIndex = record.ReadUInt16(offset + read);
            read += 2;

            // Zero-based index to last referenced sheet (FFFFH = deleted sheet)
            lastReferencedSheetIndex = record.ReadUInt16(offset + read);
            read += 2;
        }
        else
        {
            // Not used.
            _ = record.ReadUInt32(offset + read);
            read += 4;
        }

        string address = ParseCellAddress(record, biffVersion, offset + read, out var bytesRead);
        read += bytesRead;

        string addressString = !isError ? address : "#REF!";
        operands.Push($"{Get3dString5(sheetIndex, firstReferencedSheetIndex, lastReferencedSheetIndex, context)}!{addressString}");
        return read;
    }

    private static int ParsePtgRef3d8(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context)
    {
        // [MS-XLS] 2.5.198.85 PtgRef3d
        // The PtgRef3d operand specifies a reference to a single cell in an external workbook.
        // ixti (2 bytes): If the formula containing this structure is not part of a
        // revision as specified in the Formulas overview, then this value is an
        // XtiIndex that specifies the XTI which specifies those sheets. Otherwise
        // it is undefined and MUST be ignored.
        int sheetIndex = record.ReadUInt16(offset + read);
        read += 2;

        // loc (4 bytes):  A value that specifies coordinates of the referenced cell.
        // If this PtgRef3d is part of a NameParsedFormula then this is a RgceLocRel
        // value. Otherwise it is a RgceLoc value.
        string address = ParseCellAddress(record, biffVersion, offset + read, out var bytesRead);
        read += bytesRead;

        operands.Push($"{Get3dString8(sheetIndex, context)}!{address}");
        return read;
    }

    private static int ParsePtgArea3d5(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context, bool isError)
    {
        // BIFF5 stores PtgArea3d differently.
        // Not documented in MS-XLS.
        // Also documented in https://www.openoffice.org/sc/excelfileformat.pdf#page=26&zoom=100,114,662
        // Token tArea3d for 3D references, BIFF5:
        // This is always a negative value to indicate a 3D reference. The absolute value
        // is the onebased index to EXTERNSHEET record (➜5.41) in the Local Link Table
        // (➜4.10.2) containing the name of the first referenced sheet.
        // Token tArea3d for 3D references, BIFF5:
        // This is always a positive value to indicate an external reference. One-based
        // index to EXTERNSHEET record (➜5.41) in the Local Link Table (➜4.10.2).
        int sheetIndex = record.ReadInt16(offset + read);
        read += 2;

        // Not used.
        _ = record.ReadUInt32(offset + read);
        read += 4;

        _ = record.ReadUInt32(offset + read);
        read += 4;

        int? firstReferencedSheetIndex = null;
        int? lastReferencedSheetIndex = null;
        if (sheetIndex < 0)
        {
            // Zero-based index to first referenced sheet (FFFFH = deleted sheet)
            firstReferencedSheetIndex = record.ReadUInt16(offset + read);
            read += 2;

            // Zero-based index to last referenced sheet (FFFFH = deleted sheet)
            lastReferencedSheetIndex = record.ReadUInt16(offset + read);
            read += 2;
        }
        else
        {
            // Not used.
            _ = record.ReadUInt32(offset + read);
            read += 4;
        }

        string address = ParseCellAddressRange(record, biffVersion, offset + read, out var bytesRead);
        read += bytesRead;

        string addressString = !isError ? address : "#REF!";
        operands.Push($"{Get3dString5(sheetIndex, firstReferencedSheetIndex, lastReferencedSheetIndex, context)}!{addressString}");
        return read;
    }

    private static int ParsePtgArea3d8(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, XlsFormulaReaderContext context)
    {
        // [MS-XLS] 2.5.198.28 PtgArea3d
        // The PtgArea3d operand specifies a reference to a rectangular range of
        // cells in an external workbook.
        // ixti (2 bytes): If the formula containing this structure is not part of a revision
        // as specified in the Formulas overview, then this value is an XtiIndex that
        // specifies the XTI which specifies those sheets. Otherwise it is undefined and MUST
        // be ignored.
        var ixti = record.ReadUInt16(offset + read);
        read += 2;

        // area (8 bytes):  A value that specifies coordinates of the referenced range of
        // cells. If this PtgArea3d is part of a NameParsedFormula then this is an RgceAreaRel
        // value. Otherwise it is an RgceArea value.
        string address = ParseCellAddressRange(record, biffVersion, offset + read, out var bytesRead);
        read += bytesRead;

        operands.Push($"{Get3dString8(ixti, context)}!{address}");
        return read;
    }

    private static int ParsePtgAttr(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands, StringBuilder formulaString, ref bool isBasicAssignment)
    {
        byte ptgSubType = record.ReadByte(offset + read);
        read++;

        switch (ptgSubType)
        {
            case 0x00:
                {
                    // Not documented but this appears when
                    // IF({TRUE,FALSE}, 1, 0) is used.
                    _ = record.ReadByte(offset + read);
                    read += 1;
                    break;
                }

            case 0x01:
                read = ParsePtgAttrSemi(record, biffVersion, offset, read);
                break;

            case 0x02:
                read = ParsePtgAttrIf(record, biffVersion, offset, read);
                break;

            case 0x04:
                read = ParsePtgAttrChoose(record, biffVersion, offset, read);
                break;

            case 0x08:
                read = ParsePtgAttrGoto(record, biffVersion, offset, read);
                break;

            case 0x10:
                read = ParsePtgAttrSum(record, biffVersion, offset, read, operands);
                break;

            case 0x20:
                read = ParsePtgAttrBaxcel(record, biffVersion, offset, read, ref isBasicAssignment);
                break;

            case 0x40:
                read = ParsePtgAttrSpace(record, offset, read, formulaString);
                break;

            default:
                throw new NotSupportedException($"PTG 0x19 subtype 0x{ptgSubType:X2} not supported in formula string parsing.");
        }

        return read;
    }

    private static int ParsePtgAttrSemi(XlsBiffRecord record, int biffVersion, int offset, int read)
    {
        // [MS-XLS] 2.5.198.37 PtgAttrSemi
        // The PtgAttrSemi structure specifies that this Rgce is volatile.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgAttrSemi differently.
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // unused (2 bytes): Undefined and MUST be ignored.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgAttrIf(XlsBiffRecord record, int biffVersion, int offset, int read)
    {
        // [MS-XLS] 2.5.198.36 PtgAttrIf
        // The PtgAttrIf structure specifies a control token.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgAttrIf differently.
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // offset (2 bytes): An unsigned integer that specifies the byte offset.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgAttrChoose(XlsBiffRecord record, int biffVersion, int offset, int read)
    {
        // [MS-XLS] 2.5.198.34 PtgAttrChoose
        // The PtgAttrChoose structure specifies a control token.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgAttrChoose differently.
            byte cOffset = record.ReadByte(offset + read);
            read += 1;

            for (int i = 0; i < cOffset; i++)
            {
                _ = record.ReadByte(offset + read);
                read += 1;
            }

            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // cOffset (2 bytes): An unsigned integer that specifies
            // a value which is 1 less than the number of elements
            // in rgOffset.
            ushort cOffset = record.ReadUInt16(offset + read);
            read += 2;

            // rgOffset (variable): An array of 2-byte unsigned
            // integers that specifies the byte offsets.
            for (int i = 0; i < cOffset; i++)
            {
                _ = record.ReadUInt16(offset + read);
                read += 2;
            }

            // Not documented in MS-XLS, but there is an extra two bytes here.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgAttrGoto(XlsBiffRecord record, int biffVersion, int offset, int read)
    {
        // [MS-XLS] The PtgAttrGoto structure specifies a control token.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgAttrGoto differently.
            // Distance (number of bytes) from start of next token
            // to destination position, decreased by 1
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            // A - reserved1 (1 bit):  MUST be zero, and MUST be ignored.
            // B - reserved2 (3 bits):  MUST be zero, and MUST be ignored.
            // C - bitGoto (1 bit):  If the formula (section 2.2.2) containing
            // this structure is not part of a ArrayParsedFormula then the
            // bit is reserved and MUST be 1. If the formula containing this
            // structure is part of an ArrayParsedFormula, then the bit is
            // undefined and MUST be ignored.
            // D - reserved3 (4 bits):  MUST be zero, and MUST be ignored.
            //
            // offset (2 bytes): An unsigned integer that specifies a value 1
            // less than the byte offset.
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        return read;
    }

    private static int ParsePtgAttrSum(XlsBiffRecord record, int biffVersion, int offset, int read, Stack<string> operands)
    {
        // [MS-XLS] 2.5.198.41 PtgAttrSum
        // The PtgAttrSum structure specifies the sum
        // of an expression as defined in function-call.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgAttrSum differently.
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        operands.Push($"SUM({operands.Pop()})");
        return read;
    }

    private static int ParsePtgAttrBaxcel(XlsBiffRecord record, int biffVersion, int offset, int read, ref bool isBasicAssignment)
    {
        // [MS-XLS] 2.5.198.33 PtgAttrBaxcel
        // The PtgAttrBaxcel structure specifies that the result of
        // the Rgce is to be assigned to a local variable used in a
        // macro sheet.
        if (biffVersion == 2)
        {
            // BIFF2 stores PtgAttrBaxcel differently.
            _ = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            _ = record.ReadUInt16(offset + read);
            read += 2;
        }

        isBasicAssignment = true;
        return read;
    }

    private static int ParsePtgAttrSpace(XlsBiffRecord record, int offset, int read, StringBuilder formulaString)
    {
        // [MS-XLS] 2.5.198.38 PtgAttrSpace
        // The PtgAttrSpace display token specifies a number of space or carriage
        // return characters that are displayed around the expression in a
        // display-precedence-expression.
        // reserved2 (6 bits):  MUST be zero, and MUST be ignored.
        // B - bitSpace (1 bit): Reserved. MUST be 1.
        // C - reserved3 (1 bit):  MUST be zero, and MUST be ignored.
        _ = record.ReadByte(offset + read);

        // type (2 bytes): A PtgAttrSpaceType that specifies a number of space
        // or carriage return characters and the position of those characters
        // [MS-XLS] 2.5.198.40 PtgAttrSpaceType
        // The PtgAttrSpaceType structure specifies the number of space or
        // carriage return characters and position of those characters.
        // type (1 byte):  An unsigned integer that specifies the character
        // and position of the character. MUST be a value from the following
        // table:
        // 0x00 Specifies space characters before a base-expression.
        // 0x01 Specifies carriage return characters before a base-expression.
        // 0x02 Specifies space characters before the open parenthesis specified
        // by PtgParen in a display-precedence-specifier.
        // 0x03 Specifies carriage return characters before the open parenthesis
        // specified by PtgParen in a display-precedence-specifier.
        // 0x04 Specifies space characters before the close parenthesis specified
        // by PtgParen in a display-precedence-specifier.
        // 0x05 Specifies carriage return characters before the close parenthesis
        // specified by PtgParen in a display-precedence-specifier.
        // 0x06 Specifies space characters before an expression.
        var spaceType = record.ReadByte(offset + read);
        read++;

        // cch (1 byte): An unsigned integer that specifies
        // the number of characters.
        var charCount = record.ReadByte(offset + read);
        read++;

        switch (spaceType)
        {
            case 0x00:
            case 0x02:
            case 0x04:
            case 0x06:
                formulaString.Append(' ', charCount);
                break;
            case 0x01:
            case 0x03:
            case 0x05:
                formulaString.Append('\r', charCount);
                break;
            default:
                throw new NotSupportedException($"PTG PtgAttrSpaceType with type 0x{spaceType:X2} not supported in formula string parsing.");
        }

        return read;
    }

    private static string ParseCellAddress(XlsBiffRecord record, int biffVersion, int offset, out int bytesRead)
    {
        int row;
        bool rowRelative;
        int col;
        bool colRelative;
        int read = 0;

        if (biffVersion <= 5)
        {
            (row, rowRelative, colRelative) = ReadBiff2RowIndex(record, offset + read, out var bytesRead1);
            read += bytesRead1;

            col = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            row = record.ReadUInt16(offset + read);
            read += 2;

            (col, colRelative, rowRelative) = ReadBiff8ColIndexU(record, offset + read, out var bytesRead1);
            read += bytesRead1;
        }

        bytesRead = read;
        return GetAddressString(col, colRelative, row, rowRelative);
    }

    private static string ParseCellAddressRange(XlsBiffRecord record, int biffVersion, int offset, out int bytesRead)
    {
        int rowFirst;
        bool rowFirstRelative;
        int rowLast;
        bool rowLastRelative;
        int colFirst;
        bool colFirstRelative;
        int colLast;
        bool colLastRelative;
        int read = 0;

        if (biffVersion <= 5)
        {
            (rowFirst, rowFirstRelative, colFirstRelative) = ReadBiff2RowIndex(record, offset + read, out var bytesRead1);
            read += bytesRead1;

            (rowLast, rowLastRelative, colLastRelative) = ReadBiff2RowIndex(record, offset + read, out var bytesRead2);
            read += bytesRead2;

            colFirst = record.ReadByte(offset + read);
            read += 1;

            colLast = record.ReadByte(offset + read);
            read += 1;
        }
        else
        {
            rowFirst = record.ReadUInt16(offset + read);
            read += 2;

            rowLast = record.ReadUInt16(offset + read);
            read += 2;

            (colFirst, colFirstRelative, rowFirstRelative) = ReadBiff8ColIndexU(record, offset + read, out var bytesRead1);
            read += bytesRead1;

            (colLast, colLastRelative, rowLastRelative) = ReadBiff8ColIndexU(record, offset + read, out var bytesRead2);
            read += bytesRead2;
        }

        bytesRead = read;
        return GetRangeString(biffVersion, colFirst, colFirstRelative, rowFirst, rowFirstRelative, colLast, colLastRelative, rowLast, rowLastRelative);
    }

    private static string Get3dString5(int sheetIndex, int? firstReferencedSheetIndex, int? lastReferencedSheetIndex, XlsFormulaReaderContext context)
    {
        XlsBiffExternalSheet sheet = context.GetExternalSheet5(sheetIndex);
        if (sheet.IsSelf && firstReferencedSheetIndex.HasValue && lastReferencedSheetIndex.HasValue)
        {
            // Internal sheet reference.
            string firstSheetName;
            if (firstReferencedSheetIndex.Value == 0xFFFF)
            {
                firstSheetName = "#REF";
            }
            else
            {
                firstSheetName = context.Sheets[firstReferencedSheetIndex.Value].GetSheetName(context.Encoding);
            }

            string lastSheetName;
            if (lastReferencedSheetIndex.Value == 0xFFFF)
            {
                lastSheetName = "#REF";
            }
            else
            {
                lastSheetName = context.Sheets[lastReferencedSheetIndex.Value].GetSheetName(context.Encoding);
            }

            return firstSheetName == lastSheetName
                ? firstSheetName
                : $"{firstSheetName}:{lastSheetName}";
        }
        else
        {
            // External sheet reference.
            // If the sheet's name is in the format [Workbook]Sheet
            // and Workbook and Sheet are the same, just display
            // Sheet.
            var index = sheet.Name.IndexOf(']');
            if (sheet.Name.Length > 0 && sheet.Name[0] == '[' && index > 0)
            {
                var workbookName = sheet.Name[1..index];
                var sheetName = sheet.Name[(index + 1)..];
                if (workbookName == sheetName)
                {
                    return sheetName;
                }
            }
            
            return sheet.Name;
        }
    }

    private static string Get3dString8(int ixti, XlsFormulaReaderContext context)
    {
        (XlsBiffSupBook workbook, XlsBiffXti externalSheet) = context.GetExternalSheet8(ixti);
        string firstSheetName = GetXtiSheetName(externalSheet.FirstSheetIndex, workbook, context);
        string lastSheetName = GetXtiSheetName(externalSheet.LastSheetIndex, workbook, context);

        if (firstSheetName == lastSheetName || string.IsNullOrEmpty(lastSheetName))
        {
            return $"{workbook.Url}!{firstSheetName}";
        }
        else
        {
            return $"{workbook.Url}!{firstSheetName}:{lastSheetName}";
        }
    }

    private static string GetXtiSheetName(int sheetIndex, XlsBiffSupBook workbook, XlsFormulaReaderContext context)
    {
        switch (sheetIndex)
        {
            case -2:
                return string.Empty;
            case -1:
                return "#REF!";
            default:
                if (workbook.Type == XlsBiffSupBookType.Internal3DReference)
                {
                    if (sheetIndex < 0 || sheetIndex >= context.Sheets.Count)
                    {
                        throw new InvalidOperationException($"Invalid sheet index {sheetIndex} in internal workbook.");
                    }

                    return context.Sheets[sheetIndex].GetSheetName(context.Encoding);
                }
                else
                {
                    if (sheetIndex < 0 || sheetIndex >= workbook.Sheets.Count)
                    {
                        throw new InvalidOperationException($"Invalid sheet index {sheetIndex} in external workbook '{workbook.Url}'.");
                    }

                    return workbook.Sheets[sheetIndex];
                }
        }
    }
}
