namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Util
{
    public class VersionMigrator
  {
    string oldVersion = null;
    public VersionMigrator(string aFromVersion)
    {
      oldVersion = aFromVersion;
    }
    public Predefined.PredefinedParameterHandler Migrate(Predefined.PredefinedParameterHandler aHandler)
    {
      if (aHandler is Predefined.MeasurementCopyHandler)
      {
        Predefined.MeasurementCopyHandler tmpMH = (Predefined.MeasurementCopyHandler)aHandler;

        Predefined.MultipleMeasurementTableHandler tmpMultiMeasurementTableHandler = new Predefined.MultipleMeasurementTableHandler(tmpMH.testDefinition);
        tmpMultiMeasurementTableHandler.setInternalState(tmpMH.baseSheetName,
                                                          tmpMH.nrOfTables,
                                                          tmpMH.measurementTableMap,
                                                          tmpMH.sheets,
                                                          tmpMH.tableLinks,
                                                          tmpMH.firstTable,
                                                          tmpMH.initialSheetId);
        return tmpMultiMeasurementTableHandler;
      }
      if (aHandler is Predefined.MeasurementHandler)
      {
      }
      return aHandler;
    }
  }
}
