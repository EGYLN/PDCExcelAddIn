using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;
using BBS.ST.IVY.Chemistry.Util;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;


namespace BBS.ST.BHC.BSP.PDC.Lib.Util
{
    public class ExcelShapeUtils
  {
      private Process RendererProcess { get; set; }
    #region Singleton
    protected static Object missing = Type.Missing;
      private static ExcelShapeUtils utils = new ExcelShapeUtils();

    public static ExcelShapeUtils TheUtils
    {
      get
      {
        return utils;
      }
    }
    #endregion
    #region IsisDraw_MDL
    #region UseIsisOrMdl
    /// <summary>
    /// Returns true if IsisDraw or MDL is present
    /// </summary>
    /// <returns></returns>
    public bool UseIsisOrMdl()
    {
      try
      {
        return VersionUtil.IsRequiredDrawVersionInstalled();
      }
      catch (Exception)
      {
        return false;
      }
    }
    #endregion
    #region InsertISISObject
    /// <summary>
    ///   Inserts an ISIS object to the specified cell.
    /// </summary>
    /// <param name="sheet">
    ///   The Excel sheet, where the ISIS object will be inserted.
    /// </param>
    /// <param name="sheetInfo">
    ///   The SheetInfo object of the Excel sheet.
    /// </param>
    /// <param name="structureDrawingColumn">
    ///   The column in the Excel sheet, where the ISIS object will be inserted.
    /// </param>
    /// <param name="row">
    ///   The row in the Excel sheet, where the ISIS object will be inserted.
    /// </param>
    /// <param name="molFile">
    ///   The mol file of the structure.
    /// </param>
    public string InsertISISObject(Worksheet sheet, int structureDrawingColumn, int row, string molFile, UserSettings userSettings, bool exitRenderer)
    {
        if (molFile == null) return null;
        if (molFile.Trim().Equals("")) return null;
        PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject", "InsertISISObject - Method start");
        try
        {
            // insert Structure OLE object

            object[] parameters = new object[1];
            Range cellRange = (Range) sheet.Cells[row, structureDrawingColumn];
            parameters[0] = cellRange;
            RenderMolfile2Clipboard(molFile, userSettings, exitRenderer);
            Shapes shapesOld = sheet.Shapes;
            int shapeCountOld = shapesOld.Count;
            List<string> tmpOldShapeNames = new List<string>();
            Shapes shapes = sheet.Shapes;
            for (int i = 1; i <= shapeCountOld; i++)
            {
                tmpOldShapeNames.Add(shapes.Item(i).Name);
            }
            PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.DuplicateShape", "InsertISISObject.DuplicateShape - Biovia workaround");
            sheet.Paste(cellRange);
            PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.DuplicateShape", "InsertISISObject.DuplicateShape - Biovia workaround");
            shapes = sheet.Shapes;
            // the inserted shape is the shape with the highest index
            int shapeCount = sheet.Shapes.Count;
            if (shapeCount <= shapeCountOld)
            {
                PDCLogger.TheLogger.LogWarning(PDCLogger.LOG_NAME_EXCEL, "Paste of ole object failed");
                return null;
            }
            Shape shape = null;

            PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.Findshape", "InsertISISObject.Findshape - Lookup new shape. Total shape count:" + shapeCount);
            for (int j = 1; j <= shapeCount; j++)
            {
                shape = shapes.Item(j);
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, shape.Name);
                if (!tmpOldShapeNames.Contains(shape.Name))
                {
                    break;
                }

            }
            PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.Findshape", "InsertISISObject.Findshape - Looked up new shape");
            Debug.Assert(shape != null);

            PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.ResizePosition", "InsertISISObject.ResizePosition - Resize, position, ... new shape.");

            //Resize
            shape.Visible = Office.MsoTriState.msoTrue;
            shape.Select(missing);

            //shape.Placement = Excel.XlPlacement.xlFreeFloating;
            ResizeShape(shape, userSettings, cellRange);
            ResizeCellToShapeSize(cellRange, shape);
            shape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

            shape.Select(missing);
            shape.Visible = Office.MsoTriState.msoTrue;

            // because of a pasting problem from the clipboard to Excel we need to set explicitly the transparent background
            if (userSettings.TransparentBackground)
            {
                shape.Select(missing);
                shape.Fill.Visible = Office.MsoTriState.msoFalse;
            }
            shape.Select(missing);
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Select(missing);

            PDCLogger.TheLogger.LogStoptime("PDCLib.InsertISISObject.ResizePosition", "InsertISISObject.ResizePosition - Resized, positioned, ... new shape.");

            var shapeName = "IvyChemistry" + shapeCount;
            try
            {
                //Shape duplication because OLE object renders itself new (causes all Biowia settings are lost) after the set of 'Placement' property.
                PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.DuplicateShape", "InsertISISObject.DuplicateShape - Biovia workaround");

                var shapeDuplicate = shape.Duplicate();

                shapeDuplicate.Top = shape.Top;
                shapeDuplicate.Left = shape.Left;

                shapeDuplicate.Name = shapeName;
                shapeDuplicate.Placement = XlPlacement.xlMoveAndSize;

                shape.Delete();

                return shapeName;
            }
            catch (Exception exception)
            {
                PDCLogger.TheLogger.LogDebugMessage("Duplicate Shape", "Duplicate of shape is faulted Exception:" + exception);
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("PDCLib.InsertISISObject.DuplicateShape", "InsertISISObject.DuplicateShape - Biovia workaround");
            }

            try
            {
                shape.Name = shapeName;
            }
            catch (Exception)
            {
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL,
                    "Useless Exception during setting of shapename '" + shapeName + "'");
            }

            return shape.Name;
        }
        finally
        {
            PDCLogger.TheLogger.LogStoptime("PDCLib.InsertISISObject", "InsertISISObject - Method end");
        }
    }

    private void RenderMolfile2Clipboard(string molFile, UserSettings userSettings, bool exitRenderer)
    {
        if (RendererProcess == null || RendererProcess.HasExited)
        {
            RendererProcess = StartRendererProcess(userSettings);
        }
            
        RendererProcess.StandardInput.Write(molFile);
        if (!molFile.Trim().EndsWith("$$$$"))
        {
            RendererProcess.StandardInput.WriteLine();
            RendererProcess.StandardInput.WriteLine("$$$$");
        }
        string status = RendererProcess.StandardOutput.ReadLine();
        if (exitRenderer)
        {
            RendererProcess.StandardInput.Close();
        }
        if (!string.Equals(status, "ok", StringComparison.OrdinalIgnoreCase))
        {
            throw new Exception(status);
        }
    }

      private static Process StartRendererProcess(UserSettings userSettings)
      {
          
          StringWriter writer = new StringWriter();
          new XmlSerializer(typeof(UserSettings)).Serialize(writer, userSettings);
          string serializedSettings = writer.ToString();
          string settingsFile = Path.GetTempFileName();
          File.WriteAllText(settingsFile, writer.ToString(),Encoding.Unicode);
          string localPath = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
          string executable = Path.Combine(localPath, "Molfile2Clipboard.exe");
            if (!File.Exists(executable))
          {
              throw new Exception("ClipboardRenderer executable for molfile not found.");
          }
          ProcessStartInfo info = new ProcessStartInfo(executable)
          {
              Arguments = settingsFile,
              RedirectStandardInput = true,
              RedirectStandardOutput = true,
              UseShellExecute = false
          };
          Process rendererProcess = Process.Start(info);
          return rendererProcess;
      }

      private DisplayPreferences DisplayPreferencesByUserSetting(UserSettings userSettings)
    {
      DisplayPreferences displayPreferences =DisplayPreferences.GetSystemDefaultPreferences();
      // set the display preferences as selected in the user settings
      displayPreferences.BondLength = userSettings.BondLength / 2.54 * 720;
      displayPreferences.ChemLabelFont = userSettings.ChemLabelFont;
      displayPreferences.DisplayCarbonLabels = userSettings.DisplayCarbonLabels;
      displayPreferences.HydrogenDisplayMode = userSettings.HydrogenDisplayMode;
      displayPreferences.TextFont = userSettings.TextFont;
      displayPreferences.TransparentBackground = userSettings.TransparentBackground;
      displayPreferences.ColorAtomsByType = userSettings.AtomColor;
      return displayPreferences;
    }

    public static  void DebugSizes(Shape shape, Range cellRange)
    {
      if (shape != null)
      {
        float shapeTop = shape.Top;
        float shapeLeft = shape.Left;
        float shapeWidth = shape.Width;
        float shapeHeight = shape.Height;
        Debug.WriteLine("shapeTop:" + shapeTop);
        Debug.WriteLine("shapeLeft:" + shapeLeft);
        Debug.WriteLine("shapeWidth:" + shapeWidth);
        Debug.WriteLine("shapeHeight:" + shapeHeight);
      }
      if (cellRange != null)
      {
        double cellTop = (double)cellRange.Top;
        double cellLeft = (double)cellRange.Left;
        double cellHeight = (double)cellRange.RowHeight;
        double cellWidth = (double)cellRange.Width;
        double cellColumnWidth = (double)cellRange.ColumnWidth;

        
       
        Debug.WriteLine("cellTop:" + cellTop);
        Debug.WriteLine("cellLeft:" + cellLeft);
        Debug.WriteLine("cellWidth:" + cellWidth);
        Debug.WriteLine("cellHeight:" + cellHeight);
        Debug.WriteLine("cellColumnWidth:" + cellColumnWidth);
        if (cellColumnWidth == 0)
        {
          Debug.WriteLine("ratio: infinite");
        }
        else
        {
          Debug.WriteLine("ratio:" + cellWidth / cellColumnWidth);
        }

        

      }

    }

    /// <summary>
    /// Resizes the cell according to the user settings
    /// </summary>
    /// <param name="cellRange"></param>
    /// <param name="userSettings"></param>
    private void ResizeCell(Range cellRange, UserSettings userSettings, Shape shape)
    {
      double shapeWidthInColumns = 0;
      Range row = cellRange.EntireRow;
      Range column = cellRange.EntireColumn;

      switch (userSettings.ResizeMode)
      {
        case UserSettings.ResizeModes.Maximum:
              double  height = shape.Height+2;
              double width = (shape.Width+5)*(double)cellRange.ColumnWidth/(double)cellRange.Width;

          if ((double)cellRange.RowHeight < height)
          {
            cellRange.RowHeight = height;
          }
          if ((double)cellRange.ColumnWidth < width)
          {
            cellRange.ColumnWidth = width;
          }
          break;
        case UserSettings.ResizeModes.FixedWidth:
          if ((double)cellRange.ColumnWidth < userSettings.ColumnWidth)
          {
            cellRange.ColumnWidth = userSettings.ColumnWidth;
          }
          if ((double)cellRange.RowHeight < shape.Height)
          {
            cellRange.RowHeight = shape.Height +1;
          }

          break;
        case UserSettings.ResizeModes.FixedHeight:
          if ((double)cellRange.RowHeight < userSettings.RowHeight)
          {
            cellRange.RowHeight = userSettings.RowHeight +1;
          }
              
          double rw = (double) column.Width;
          double rcw = (double) column.ColumnWidth;
          
          shapeWidthInColumns = shape.Width *((double)column.ColumnWidth) / (double)column.Width;
          if ((double)column.ColumnWidth < shapeWidthInColumns)
          {
            
            column.ColumnWidth = Math.Round(shapeWidthInColumns)+1;
            double nw = (double)cellRange.Width;
            Debug.WriteLine("CellRange New Width:" + nw);
            DebugSizes(shape, cellRange);
          }
          break;
        case UserSettings.ResizeModes.StructureDefault:
          if ((double)cellRange.RowHeight < shape.Height)
          {
            cellRange.RowHeight = shape.Height + 1.0f;
          }
          shapeWidthInColumns = shape.Width * ((double) cellRange.ColumnWidth) / ((double)cellRange.Width);
          if ((double) cellRange.ColumnWidth < shapeWidthInColumns)
          {
            cellRange.ColumnWidth = Math.Round(shapeWidthInColumns + 1);
          }
          return; 
      }
    }
    #endregion

    #endregion
    #region SetPlacement4Shapes
    /// <summary>
    ///   Sets the placement property of all IVY Chemistry objects to xlMoveAndSize
    /// </summary>
    /// <param name="sheet">
    ///   The sheet with the IVY Chemistry objects.
    /// </param>
    public void SetPlacement4Shapes(Worksheet sheet)
    {
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape.Name != null && shape.Name.StartsWith("IvyChemistry"))
            {
                shape.Select(missing);
                shape.Placement = XlPlacement.xlMoveAndSize;
            }
        }
    }
    #endregion

    #region InsertStructureDrawing
    public string InsertStructureDrawing(Worksheet aSheet, TestStruct aStruct, int aColumn, int aRow, UserSettings userSettings, String servletPath, bool useIsisOrMdl, bool exitRenderer)
    {
        if (!useIsisOrMdl)
        {
            return InsertImageShape(aSheet, aColumn, aRow, aStruct, servletPath, userSettings);
        }
        return  InsertISISObject(aSheet,aColumn, aRow, aStruct.molfile, userSettings, exitRenderer);
    }
    #endregion
    #region InsertImageShape
    public string InsertImageShape(Worksheet tmpSheet, int aColumn, int aRow, TestStruct tmpCis, string servletPath, UserSettings userSettings)
    {

      if (tmpCis.molimagearray != null)
        {
          string tmpFileName = Path.GetTempFileName();
          FileStream tmpStream = File.OpenWrite(tmpFileName);
          tmpStream.Write(tmpCis.molimagearray, 0, tmpCis.molimagearray.Length);
          tmpStream.Flush();
          tmpStream.Close();
          //das Feld Tag enthält dann den Link auf das Image File 
        
          Range tmpCellRange = (Range)tmpSheet.Cells[aRow, aColumn];
          double tmpLeft = (double)tmpCellRange.Left;
          double tmpTop = (double)tmpCellRange.Top;
          double tmpWidth = (double)tmpCellRange.Width;
          double tmpHeight = (double)tmpCellRange.Height;

          Shape tmpShape = tmpSheet.Shapes.AddPicture(tmpFileName,
                                                        Office.MsoTriState.msoFalse,
                                                        Office.MsoTriState.msoTrue,
                                                        (float)tmpLeft,
                                                        (float)tmpTop,
                                                        (float)tmpWidth,
                                                        (float)tmpHeight);
          // delete tempFile from disk
          if (File.Exists(tmpFileName))
          {
            File.Delete(tmpFileName);
          }
          tmpShape.AlternativeText = "Structure Drawing of " + tmpCis.compoundno;
          tmpShape.Visible = Office.MsoTriState.msoTrue;
          tmpShape.Select(missing);

          string tmpImageName = tmpShape.Name;

          tmpShape.Placement = XlPlacement.xlFreeFloating;
          tmpShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

          ResizeShape(tmpShape, userSettings, tmpCellRange);
          ResizeCell(tmpCellRange, userSettings, tmpShape);
          ResizeShapeToCellSize(tmpCellRange, tmpShape);
          ResizeCellToShapeSize(tmpCellRange, tmpShape);
          tmpShape.Select(missing);
          tmpShape.Visible = Office.MsoTriState.msoTrue;
          tmpShape.Select(missing);
          tmpShape.Placement = XlPlacement.xlMoveAndSize;
          if (servletPath != null)
          {
            string tmpHL = PDCService.ThePDCService.ServerURL +
              servletPath + "?compoundno=" +
              tmpCis.compoundno.Replace(" ", "%20");

            //System.Web.HttpUtility.UrlEncode(tmpCis.CompoundNo);
            tmpSheet.Hyperlinks.Add(tmpShape, tmpHL, missing, "Click to follow MolFileService Hyperlink", tmpCis.compoundno);
          }
          return tmpImageName;
        }
        tmpCis.filename = string.Empty;
        return null;
    }
    #endregion
    public float MaxShapeWidth(Range aRange)
    {
      Range range = aRange.EntireColumn;
      double rw = (double) range.Width;
      double cw = (double) range.ColumnWidth;
      return (float) ((UserSettings.MAX_COLUMN_WIDTH-1.0)* (rw/cw));

    }
    private void ResizeShapeToCellSize(Range cellRange, Shape shape)
    {
        double ch = ((double) cellRange.EntireRow.RowHeight)-1.0;
        double cw = ((double) cellRange.EntireColumn.Width)-1.0;

        double sh = shape.Height;
        double sw = shape.Width;
        double scale = ch / sh;
        double scale2 = cw / sw;
        scale = Math.Min(scale, scale2);
        shape.Width = (float) (sw * scale);
        if (shape.Height == sh)
        {
            shape.Height = (float)(sh * scale);
        }
    }
    /// <summary>
    /// resizes the cell to the size of the shape
    /// </summary>
    /// <param name="cellRange">the range to be resize</param>
    /// <param name="shape">the size giving shape</param>
    public void ResizeCellToShapeSize(Range cellRange, Shape shape)
    {
      float maxShapeWidth = MaxShapeWidth(cellRange);
      if (shape.Width > maxShapeWidth || shape.Height >= UserSettings.MAX_ROW_HEIGHT)
      {
        double xMax = (double) cellRange.Width / (double) cellRange.ColumnWidth * (maxShapeWidth);
        double yMax = (double)cellRange.Height / (double)cellRange.RowHeight * (UserSettings.MAX_ROW_HEIGHT-1);
        ResizeShapeToGivenMaximum(shape, xMax, yMax);

      }
      double minWidth = (shape.Width+1) * ((double)cellRange.ColumnWidth) / ((double)cellRange.Width) ;
      double minHeigth = (shape.Height+1) * ((double)cellRange.RowHeight) / ((double)cellRange.Height) ;
      if (minWidth > UserSettings.MAX_COLUMN_WIDTH)
      {
          minWidth = UserSettings.MAX_COLUMN_WIDTH;
      }
      if (minHeigth > UserSettings.MAX_ROW_HEIGHT)
      {
          minHeigth = UserSettings.MAX_ROW_HEIGHT;
      }
      if ((double)cellRange.Width <= shape.Width)
      {
        cellRange.EntireColumn.ColumnWidth = minWidth;
      }
      if ((double)cellRange.Height <= shape.Height)
      {
        cellRange.EntireRow.RowHeight = minHeigth;
      }
      // if after resizing the cell the image is still bigger, resize the image now!
      // Adjust the size of the shape.... SUPERSTUPID FINETUNING Workaround as setting cellRange.EntireColumn.ColumnWidth  
      // sets the number of characters and Excel does stuff itsown which is dump
      // resize
      int smallerizer = 0;
      while ((double)cellRange.Width <= shape.Width)
      {
        ResizeShapeToGivenMaximum(shape, (double)cellRange.Width - smallerizer, (double)cellRange.Height-smallerizer);
        smallerizer++;
      }
      smallerizer = 1;
      while ((double)cellRange.Height <= shape.Height)
      {
        ResizeShapeToGivenMaximum(shape, (double)cellRange.Width - smallerizer, (double)cellRange.Height- smallerizer);
        smallerizer++;
      }
    }

    public void ResizeShapeToGivenMaximum(Shape shape, double xMax,double yMax)
    {
      double sw = shape.Width;

        if (sw > xMax)
        {
            shape.Width = (float) xMax;
        };
        double sh = shape.Height;
        if (sh > yMax)
        {
            shape.Height = (float) yMax;
        }
    }

    /// <summary>
    /// Constraints the scale such that current is not scaled above maximum
    /// </summary>
    /// <param name="aScale"></param>
    /// <param name="current"></param>
    /// <param name="maximum"></param>
    /// <returns></returns>
    private double ConstraintScale(double aScale, double current, double maximum)
    {
      if (aScale * current > maximum)
      {
        return maximum / current;
      }
      return aScale;
    }
    /// <summary>
    /// Returns the maximum shape width as specified by the user settings.
    /// </summary>
    /// <param name="cellRange"></param>
    /// <param name="userSettings"></param>
    /// <returns></returns>
    public float MaxUserWidth(Range cellRange, UserSettings userSettings)
    {
      Range range = cellRange.EntireColumn;
      //Range.Width is the width in pix.
      //Range.ColumnWidth is the width in characters
      return (float) ((userSettings.ColumnWidth-0.5) * ((double)range.Width / (double)range.ColumnWidth));
    }
    #region ResizeShape
    public void ResizeShape(Shape shape, UserSettings userSettings, Range cellRange)
    {
        try
        {
            //      DebugSizes(shape, cellRange);
            if (!UseIsisOrMdl())
            {
                shape.ScaleHeight(1.0f, Office.MsoTriState.msoTrue, missing);
                shape.ScaleWidth(1.0f, Office.MsoTriState.msoTrue, missing);
                shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            DebugSizes(shape, cellRange);
            double scale = 1.0;
            float h = shape.Height;
            float w = shape.Width;
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "X: " + w + ", Y: " + h);
            switch (userSettings.ResizeMode)
            {
                case UserSettings.ResizeModes.StructureDefault:
                    scale = 1.0;
                    scale = ConstraintScale(scale, w, MaxShapeWidth(cellRange));
                    scale = ConstraintScale(scale, h, UserSettings.MAX_ROW_HEIGHT);
                    break; ;
                case UserSettings.ResizeModes.FixedWidth:
                    scale = MaxUserWidth(cellRange, userSettings) / shape.Width;
                    scale = ConstraintScale(scale, h, UserSettings.MAX_ROW_HEIGHT);
                    break;
                case UserSettings.ResizeModes.FixedHeight:
                    scale = userSettings.RowHeight / shape.Height;
                    scale = ConstraintScale(scale, w, MaxShapeWidth(cellRange));
                    break;
                case UserSettings.ResizeModes.Maximum:
                    scale = CalcScale(cellRange, w, h, MaxUserWidth(cellRange, userSettings), userSettings.RowHeight);
                    break;
            }
            shape.Width = (float)(shape.Width * scale);
            if (shape.Height == h)
            {
                shape.Height = (float)(shape.Height * scale);
            }
            //shape.ScaleHeight((float) scale, Office.MsoTriState.msoTrue, missing);
            DebugSizes(shape, cellRange);
            //shape.ScaleWidth((float) scale, Office.MsoTriState.msoTrue, missing);
            //shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
        }
        catch (Exception e)
        {
            PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Fehler beim Resizen von " + shape.Name, e);
        }
    }

    private double CalcScale(Range cellRange, double currentX, double currentY, double targetX, double targetY)
    {
      double scale = targetX / currentX;
      scale = Math.Max(scale, targetY / currentY);
      if (scale * currentX > MaxShapeWidth(cellRange))
      {
        scale = MaxShapeWidth(cellRange) / currentX;
      }
      if (scale * currentY > UserSettings.MAX_ROW_HEIGHT)
      {
        scale = UserSettings.MAX_ROW_HEIGHT / currentY;
      }
      return scale;
    }
    #endregion
  }
    
}
