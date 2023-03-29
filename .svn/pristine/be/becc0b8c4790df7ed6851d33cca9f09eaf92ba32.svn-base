using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

#pragma warning disable 0168

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Categorization of SheetChanged Events
    /// </summary>
    [ComVisible(false)]
  public enum SheetEventType
  {
    ROW_ADDED,
    ROW_DELETED,
    ROW_MOVED,
    COLUMN_ADDED,
    COLUMN_DELETED,
    COLUMN_MOVED,
    CONTENTS_CHANGED
  };

  /// <summary>
  /// A SheetEvent represents an Excel SheetChanged event with
  /// a more detailed classification of the event.
  /// </summary>
  [ComVisible(false)]
  public struct SheetEvent
  {
    public SheetEventType eventType;
    public int row;
    public int column;
    public int width;
    public int height;
    public bool entireRow;
    public bool entireColumn;

    public override string ToString()
    {
        StringBuilder builder = new StringBuilder("SheetEvent: ");
        builder.Append("Type: ").Append(eventType);
        builder.Append("; Row: ").Append(row);
        builder.Append("; Column: ").Append(column);
        builder.Append("; Height: ").Append(height);
        builder.Append("; Width: ").Append(width);
        builder.Append("; EntireRow: ").Append(entireRow);
        builder.Append("; EntireColumn: ").Append(entireColumn);
        return builder.ToString();
    }
      
    public static SheetEvent AsSheetEvent(Excel.Worksheet aSheet, System.Drawing.Rectangle aRectangle)
    {
        SheetEvent tmpEvent = new SheetEvent();
        tmpEvent.column = aRectangle.X;
        tmpEvent.row = aRectangle.Y;
        tmpEvent.height = aRectangle.Height;
        tmpEvent.width = aRectangle.Width;
        tmpEvent.eventType = SheetEventType.CONTENTS_CHANGED;
        Excel.Worksheet tmpSheet = aSheet;
        tmpEvent.entireRow = tmpEvent.width == PDCListObject.max_column;
        tmpEvent.entireColumn = tmpEvent.height == PDCListObject.max_row;
        return tmpEvent;
    }
  }
}
