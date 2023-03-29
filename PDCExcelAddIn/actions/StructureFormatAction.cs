using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    class StructureFormatAction : PDCAction
  {
    public const string ACTION_TAG = "PDC_StructureFormatAction";
    private UserSettings myUserSettings;

    #region constructor
    public StructureFormatAction(bool beginGroup, UserSettings userSettings)
        : base(Properties.Resources.Action_StructureFormat_Caption, ACTION_TAG, beginGroup)
    {
      this.myUserSettings = userSettings;
    }
    #endregion

    #region SafeForCellEditingMode
    protected override bool SafeForCellEditingMode()
    {
        return true;
    }
    #endregion

    #region PerformAction
    internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
    {
      StructureFormatDialog structureFormatDialog = new StructureFormatDialog(this.myUserSettings);
      structureFormatDialog.ShowDialog();
      return new ActionStatus();
    }
    #endregion
  }
}
