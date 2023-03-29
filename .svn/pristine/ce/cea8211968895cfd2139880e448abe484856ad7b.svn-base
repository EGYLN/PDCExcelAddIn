using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using BBS.ST.IVY.Chemistry.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    ///   The dialog with the structure format settings.
    /// </summary>
    [ComVisible(false)]
  public partial class StructureFormatDialog : Form
  {
    private Font myChemLabelFont;
    private Font myTextLabelFont;
    private UserSettings myUserSettings;

    #region constructor
    /// <summary>
    ///   The constructor of the format structure dialog.
    /// </summary>
    /// <param name="userSettings">
    ///   The user settings from the registry.
    /// </param>
    public StructureFormatDialog(UserSettings userSettings)
    {
      InitializeComponent();
      myChemLabelFont = userSettings.ChemLabelFont;
      myTextLabelFont = userSettings.TextFont;
      myUserSettings = userSettings;
      this.InitializeDialog();
    }
    #endregion

    #region methods

    #region GetText4Font
    /// <summary>
    ///    Returns a semicolon seperated text, describing the given font.
    /// </summary>
    /// <param name="font">
    ///    A font.
    /// </param>
    /// <returns>
    ///    Null, if the font is null. Otherwise a semicolon seperated text, describing the given font.
    /// </returns>
    private string GetText4Font(Font font)
    {
      if (font == null) return null;
      return font.Name + "; " + font.Size.ToString() + "; " + font.Style.ToString();
    }
    #endregion

    #region InitializeDialog
    /// <summary>
    ///    This funtion initialzes the structure format dialog
    /// </summary>
    private void InitializeDialog()
    {

      ToolTip toolTip = new ToolTip();
      toolTip.AutoPopDelay = 100000;
      this.chkDisplayPrepno.Checked = myUserSettings.DisplayPrepno;
      toolTip.SetToolTip(chkDisplayPrepno, Tooltips.ChemicalDataSettings);

      this.chkDisplayStructure.Checked = myUserSettings.DisplayStructure;
      toolTip.SetToolTip(chkDisplayStructure, Tooltips.ChemicalDataSettings);
      
      this.chkDisplayMolweight.Checked = myUserSettings.DisplayMolweight;
      toolTip.SetToolTip(chkDisplayMolweight, Tooltips.ChemicalDataSettings);
      
      this.chkDisplayMolformula.Checked = myUserSettings.DisplayMolformula;
      toolTip.SetToolTip(chkDisplayMolformula, Tooltips.ChemicalDataSettings);


      this.txtAtomFont.Text = this.GetText4Font(myUserSettings.ChemLabelFont);
      toolTip.SetToolTip(txtAtomFont, Tooltips.AtomFont);
      toolTip.SetToolTip(lblFontAtom, Tooltips.AtomFont);
      toolTip.SetToolTip(btnAtomFont, Tooltips.AtomFont);

      this.txtTextFont.Text = this.GetText4Font(myUserSettings.TextFont);
      toolTip.SetToolTip(txtTextFont, Tooltips.TextFont);
      toolTip.SetToolTip(lblFontText, Tooltips.TextFont); 
      toolTip.SetToolTip(btnTextFont, Tooltips.TextFont);

      this.txtBondLength.Text = myUserSettings.BondLength.ToString(new CultureInfo("en-US"));
      toolTip.SetToolTip(txtBondLength, Tooltips.BondLength);
      toolTip.SetToolTip(rbtnBondLength, Tooltips.BondLength);

      this.chkDisplayCarbon.Checked = myUserSettings.DisplayCarbonLabels;
      toolTip.SetToolTip(chkDisplayCarbon, Tooltips.DisplayCarbon);

      this.chkTransparentBackground.Checked = myUserSettings.TransparentBackground;
      toolTip.SetToolTip(chkTransparentBackground, Tooltips.TransparentBackground);

      this.chkAtomColor.Checked = myUserSettings.AtomColor;
      toolTip.SetToolTip(chkAtomColor, Tooltips.Color);

      this.txtColumnWidth.Text = myUserSettings.ColumnWidth.ToString();
      toolTip.SetToolTip(txtColumnWidth, Tooltips.ColumnWidth);
      toolTip.SetToolTip(rbtnColumnWidth, Tooltips.ColumnWidth);
      

      this.txtRowHeight.Text = myUserSettings.RowHeight.ToString();
      toolTip.SetToolTip(txtRowHeight, Tooltips.RowHeigth);
      toolTip.SetToolTip(rbtnRowHeight, Tooltips.RowHeigth);

      this.txtHorizontalOffset.Value = myUserSettings.HorizontalOffset;
      toolTip.SetToolTip(txtHorizontalOffset, Tooltips.ShiftRows);
      toolTip.SetToolTip(label2, Tooltips.ShiftRows);

      this.txtVerticalOffset.Value = myUserSettings.VerticalOffset;
      toolTip.SetToolTip(txtVerticalOffset, Tooltips.ShiftCols);
      toolTip.SetToolTip(label3, Tooltips.ShiftCols);

      switch (myUserSettings.HydrogenDisplayMode)
      {
        case HydrogenDisplayMode.All:
          this.chkDisplayHydrogens.Checked = true;
          this.rbtnAll.Checked = true;
          break;
        case HydrogenDisplayMode.Hetero:
          this.chkDisplayHydrogens.Checked = true;
          this.rbtnHetero.Checked = true;
          break;
        case HydrogenDisplayMode.HeteroOrTerminal:
          this.chkDisplayHydrogens.Checked = true;
          this.rbtnHeteroOrTerminal.Checked = true;
          break;
        case HydrogenDisplayMode.None:
          this.chkDisplayHydrogens.Checked = true;
          this.chkDisplayHydrogens.Checked = false;
          EnableHydrogens(false);
          break;
        case HydrogenDisplayMode.Terminal:
          this.chkDisplayHydrogens.Checked = true;
          this.rbtnTerminal.Checked = true;
          break;
      }
      toolTip.SetToolTip(chkDisplayHydrogens, Tooltips.DisplayHydrogens);

      if (myUserSettings.Orientation == UserSettings.Direction.Vertical)
      {
          this.rbtnDownFromOffset.Checked = true;
      }
      else
      {
          this.rbtnRightToOffset.Checked = true;
      }
      toolTip.SetToolTip(rbtnDownFromOffset, Tooltips.OffsetBelow);
      toolTip.SetToolTip(rbtnRightToOffset, Tooltips.OffsetRight);

      switch (myUserSettings.ResizeMode)
      {
        case UserSettings.ResizeModes.Maximum: 
          //(A) Maximum is only used by OpenLib and should not be persisted
          goto case UserSettings.ResizeModes.FixedWidth;
        case UserSettings.ResizeModes.StructureDefault:
          this.rbtnBondLength.Checked = true;
          txtColumnWidth.Enabled = false;
          txtRowHeight.Enabled = false;
          break;
        case UserSettings.ResizeModes.FixedWidth:
          this.rbtnColumnWidth.Checked = true;
          this.txtRowHeight.Enabled = false;
          this.txtBondLength.Enabled = false;
          break;
        case UserSettings.ResizeModes.FixedHeight:
          this.rbtnRowHeight.Checked = true;
          this.txtBondLength.Enabled = false;
          this.txtColumnWidth.Enabled = false;
          break;
      }
    }
    #endregion

    #endregion

    #region events

    /// <summary>
    /// 
    /// </summary>
    /// <param name="enabled"></param>
    private void EnableHydrogens(bool enabled)
    {
        this.rbtnAll.Enabled = enabled;
        this.rbtnHetero.Enabled = enabled;
        this.rbtnHeteroOrTerminal.Enabled = enabled;
        this.rbtnTerminal.Enabled = enabled;
        if (enabled && !rbtnTerminal.Checked && !rbtnAll.Checked && !rbtnHetero.Checked && !rbtnHeteroOrTerminal.Checked)
        {
            rbtnHeteroOrTerminal.Checked = true;
        }
        if (!enabled)
        {
            rbtnHetero.Checked = false;
            rbtnHeteroOrTerminal.Checked = false;
            rbtnTerminal.Checked = false;
            rbtnAll.Checked = false;
        }
    }
    #region btnAtomFont_Click
    private void btnAtomFont_Click(object sender, EventArgs e)
    {
      FontDialog fontDialog = new FontDialog();
      if (fontDialog.ShowDialog(this) == DialogResult.OK)
      {
        this.txtAtomFont.Text = this.GetText4Font(fontDialog.Font);
        myChemLabelFont = fontDialog.Font;
      }
    }
    #endregion

    #region btnTextFont_Click
    private void btnTextFont_Click(object sender, EventArgs e)
    {
      FontDialog fontDialog = new FontDialog();
      if (fontDialog.ShowDialog(this) == DialogResult.OK)
      {
        this.txtTextFont.Text = this.GetText4Font(fontDialog.Font);
        myTextLabelFont = fontDialog.Font;
      }
    }
    #endregion

    private string DialogTitle() {
      return Properties.Resources.StructureFormatValidation_Title;
    }
    /// <summary>
    /// Checks the dialog fields for valid user input
    /// </summary>
    /// <returns></returns>
    private bool CheckUserInput()
    {
        //Check that at least one Compoundinfo is selected
        if (!(chkDisplayPrepno.Checked || chkDisplayStructure.Checked || chkDisplayMolformula.Checked || chkDisplayMolweight.Checked)) {
            MessageBox.Show(this, Properties.Resources.MSG_NO_COMPOUNDINFO_SELECTED, DialogTitle(), MessageBoxButtons.OK,MessageBoxIcon.Error);
            chkDisplayPrepno.Select();
            return false;
        }
        //Check or Valid input in number fields
        double tmpForget;
        if (rbtnBondLength.Checked)
        {
            var bondLength = txtBondLength.Text.Replace(',', '.');
            if (!double.TryParse(bondLength, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out tmpForget))
            {
                MessageBox.Show(this, string.Format(Properties.Resources.MSG_INVALID_FORMAT, txtBondLength.Text), DialogTitle(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtBondLength.Select();
                return false;
            }
            if (tmpForget < UserSettings.MIN_BOND_LENGTH || tmpForget> UserSettings.MAX_BOND_LENGTH)
            {
                MessageBox.Show(this, string.Format(Properties.Resources.MSG_OUT_OF_RANGE, UserSettings.MIN_BOND_LENGTH, UserSettings.MAX_BOND_LENGTH), DialogTitle(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtBondLength.Select();
                return false;
            }
          
        }
        if (rbtnColumnWidth.Checked)
        {
            if (!double.TryParse(txtColumnWidth.Text, out tmpForget))
            {
              MessageBox.Show(this, string.Format(Properties.Resources.MSG_INVALID_FORMAT, txtColumnWidth.Text), DialogTitle(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtColumnWidth.Select();
                return false;
            }
            if (tmpForget < Lib.Util.UserSettings.MIN_COLUMN_WIDTH || tmpForget > Lib.Util.UserSettings.MAX_COLUMN_WIDTH)
            {
              MessageBox.Show(this, string.Format(Properties.Resources.MSG_OUT_OF_RANGE, Lib.Util.UserSettings.MIN_COLUMN_WIDTH, Lib.Util.UserSettings.MAX_COLUMN_WIDTH), DialogTitle(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtColumnWidth.Select();
                return false;
            }
        }
        if (rbtnRowHeight.Checked)
        {
            if (!double.TryParse(txtRowHeight.Text, out tmpForget))
            {
              MessageBox.Show(this, string.Format(Properties.Resources.MSG_INVALID_FORMAT, txtRowHeight.Text), DialogTitle(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtRowHeight.Select();
                return false;
            }
            if (tmpForget < Lib.Util.UserSettings.MIN_ROW_HEIGHT || tmpForget > Lib.Util.UserSettings.MAX_ROW_HEIGHT)
            {
              MessageBox.Show(this, string.Format(Properties.Resources.MSG_OUT_OF_RANGE, Lib.Util.UserSettings.MIN_ROW_HEIGHT, Lib.Util.UserSettings.MAX_ROW_HEIGHT), DialogTitle(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtRowHeight.Select();
                return false;
            }

        }
        return true;

    }
    #region btnOk_Click
    private void btnOk_Click(object sender, EventArgs e)
    {
      if (!CheckUserInput())
      {
        return;
      }
      myUserSettings.ChemLabelFont = myChemLabelFont;
      myUserSettings.TextFont = myTextLabelFont;
      if (!this.chkDisplayHydrogens.Checked) { 
          myUserSettings.HydrogenDisplayMode = HydrogenDisplayMode.None; 
      } else
      {
          if (this.rbtnAll.Checked) myUserSettings.HydrogenDisplayMode = HydrogenDisplayMode.All;
          if (this.rbtnHetero.Checked) myUserSettings.HydrogenDisplayMode = HydrogenDisplayMode.Hetero;
          if (this.rbtnHeteroOrTerminal.Checked) myUserSettings.HydrogenDisplayMode = HydrogenDisplayMode.HeteroOrTerminal;
          if (this.rbtnTerminal.Checked) myUserSettings.HydrogenDisplayMode = HydrogenDisplayMode.Terminal;
      }
      if (this.rbtnBondLength.Checked)
      {
          myUserSettings.ResizeMode = UserSettings.ResizeModes.StructureDefault;
          myUserSettings.BondLength = (float) Convert.ToDouble(txtBondLength.Text.Replace(',', '.'), new CultureInfo("en-US"));
      }
      if (this.rbtnColumnWidth.Checked)
      {
          myUserSettings.ResizeMode = UserSettings.ResizeModes.FixedWidth;
          myUserSettings.ColumnWidth = (float) Convert.ToDouble(txtColumnWidth.Text);
      }
      if (this.rbtnRowHeight.Checked)
      {
          myUserSettings.ResizeMode = UserSettings.ResizeModes.FixedHeight;
          myUserSettings.RowHeight = (float) Convert.ToDouble(txtRowHeight.Text);
      }

      myUserSettings.HorizontalOffset = (int) txtHorizontalOffset.Value;
      myUserSettings.VerticalOffset = (int) txtVerticalOffset.Value;
      if (this.rbtnDownFromOffset.Checked)
      {
          myUserSettings.Orientation = UserSettings.Direction.Vertical;
      }
      else
      {
          myUserSettings.Orientation = UserSettings.Direction.Horizontal;
      }
      myUserSettings.DisplayCarbonLabels = this.chkDisplayCarbon.Checked;
      myUserSettings.TransparentBackground = this.chkTransparentBackground.Checked;
      myUserSettings.AtomColor = this.chkAtomColor.Checked;
      myUserSettings.DisplayPrepno = chkDisplayPrepno.Checked;
      myUserSettings.DisplayStructure = chkDisplayStructure.Checked;
      myUserSettings.DisplayMolformula = chkDisplayMolformula.Checked;
      myUserSettings.DisplayMolweight = chkDisplayMolweight.Checked;

      myUserSettings.WriteSettings();
      Dispose();
    }
    #endregion

    private void chkDisplayHydrogens_CheckedChanged(object sender, EventArgs e)
    {
        EnableHydrogens(chkDisplayHydrogens.Checked);
    }

    #endregion

    private void rbtnBondLength_CheckedChanged(object sender, EventArgs e)
    {
        txtBondLength.Enabled = rbtnBondLength.Checked;
    }

    private void rbtnRowHeight_CheckedChanged(object sender, EventArgs e)
    {
        txtRowHeight.Enabled = rbtnRowHeight.Checked;
    }

    private void rbtnColumnWidth_CheckedChanged(object sender, EventArgs e)
    {
        txtColumnWidth.Enabled = rbtnColumnWidth.Checked;
    }
  }
}
