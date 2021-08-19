using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.IO;
using Syncfusion.GridExcelConverter;

namespace HRMF0723
{
    public partial class HRMF0723 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0723()
        {
            InitializeComponent();
        }

        public HRMF0723(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- User Make Methods ----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void DefaultDate()
        {
            if (DateTime.Today.Month <= 2)
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(iDate.ISDate_Add(string.Format("{0}-01-01", DateTime.Today.Year), -1));
            }
            else
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            }
        }

        private DateTime GetDateTime()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        private void SEARCH_DB()
        {
            string vMessage = string.Empty;
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_STD_YYYYMM.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
            
            try
            {
                string vPERSON_NUM = iString.ISNull(IGR_SMALL_BIZ_REDUTION.GetCellValue("PERSON_NUM"));
                int vIDX_Col = IGR_SMALL_BIZ_REDUTION.GetColumnToIndex("PERSON_NUM");

                IDA_SMALL_BIZ_REDUTION.Fill();
                if (IGR_SMALL_BIZ_REDUTION.RowCount > 0)
                {
                    for (int vRow = 0; vRow < IGR_SMALL_BIZ_REDUTION.RowCount; vRow++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_SMALL_BIZ_REDUTION.GetCellValue(vRow, vIDX_Col)))
                        {
                            IGR_SMALL_BIZ_REDUTION.CurrentCellActivate(vRow, 0);
                            IGR_SMALL_BIZ_REDUTION.CurrentCellMoveTo(vRow, 0);
                        }
                    }
                }
                IGR_SMALL_BIZ_REDUTION.Focus();
            }
            catch (System.Exception ex)
            {
                vMessage = string.Format("Adapter Fill Error\n{0}", ex.Message);
                MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
         
        private void SetCommon(object pGROUP_CODE, object pENABLED_FLAG_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pENABLED_FLAG_YN);
        }
         
        #endregion;


        #region ----- Excel Export -----

        private void ExcelExport()
        {
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.Title = "Save File Name";
            saveFileDialog1.Filter = "Excel Files(*.xls)|*.xls";
            saveFileDialog1.DefaultExt = ".xls";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                vExport.GridToExcel(IGR_SMALL_BIZ_REDUTION.BaseGrid, saveFileDialog1.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.RowHeaders);

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog1.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion

        #region ----- Main Button Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_SMALL_BIZ_REDUTION.IsFocused)
                    {
                        IDA_SMALL_BIZ_REDUTION.AddOver();
                        IGR_SMALL_BIZ_REDUTION.Focus();
                    }          
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_SMALL_BIZ_REDUTION.IsFocused)
                    {
                        IDA_SMALL_BIZ_REDUTION.AddUnder();
                        IGR_SMALL_BIZ_REDUTION.Focus();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        System.Windows.Forms.SendKeys.Send("{TAB}");
                        IDA_SMALL_BIZ_REDUTION.Update();
                    }
                    catch (Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_SMALL_BIZ_REDUTION.IsFocused)
                    {
                        IDA_SMALL_BIZ_REDUTION.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_SMALL_BIZ_REDUTION.IsFocused)
                    {
                        IDA_SMALL_BIZ_REDUTION.Delete();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                }
            }
        }

        #endregion;

        #region ----- This Form Events -----

        private void HRMF0723_Load(object sender, EventArgs e)
        {
            IDA_SMALL_BIZ_REDUTION.FillSchema(); 
        }

        private void HRMF0723_Shown(object sender, EventArgs e)
        {
            DefaultDate();
            DefaultCorporation();   
        }

        private void IGR_SMALL_BIZ_REDUTION_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            if (e.ColIndex == IGR_SMALL_BIZ_REDUTION.GetColumnToIndex("ORI_JOIN_DATE"))
            {
                if (iString.ISNull(IGR_SMALL_BIZ_REDUTION.GetCellValue("REDUTION_DATE_FR")) == string.Empty)
                {
                    IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_DATE_FR", e.CellValue);

                    IDC_GET_REDUTION_DATE_TO_P.SetCommandParamValue("P_REDUTION_DATE_FR", e.CellValue);
                    IDC_GET_REDUTION_DATE_TO_P.ExecuteNonQuery();
                    IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_DATE_TO", IDC_GET_REDUTION_DATE_TO_P.GetCommandParamValue("O_REDUTION_DATE_TO"));
                }

                if (iString.ISNumtoZero(IGR_SMALL_BIZ_REDUTION.GetCellValue("JOIN_AGE"), 0) == 0)
                {
                    //취업시 나이
                    IDC_GET_AGE_P.SetCommandParamValue("P_PERSON_ID", IGR_SMALL_BIZ_REDUTION.GetCellValue("PERSON_ID"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ORI_JOIN_DATE", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));
                    IDC_GET_AGE_P.ExecuteNonQuery();
                    IGR_SMALL_BIZ_REDUTION.SetCellValue("JOIN_AGE", IDC_GET_AGE_P.GetCommandParamValue("O_AGE"));
                }
                if (iString.ISNumtoZero(IGR_SMALL_BIZ_REDUTION.GetCellValue("AGE"), 0) == 0)
                {
                    //병역제외한 나이 
                    IDC_GET_AGE_P.SetCommandParamValue("P_PERSON_ID", IGR_SMALL_BIZ_REDUTION.GetCellValue("PERSON_ID"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ORI_JOIN_DATE", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_FR", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_FR"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_TO", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_TO"));
                    IDC_GET_AGE_P.ExecuteNonQuery();
                    IGR_SMALL_BIZ_REDUTION.SetCellValue("AGE", IDC_GET_AGE_P.GetCommandParamValue("O_AGE"));
                }
            }
            else if (e.ColIndex == IGR_SMALL_BIZ_REDUTION.GetColumnToIndex("JOIN_AGE"))
            {
                if (iString.ISNumtoZero(IGR_SMALL_BIZ_REDUTION.GetCellValue("AGE"), 0) == 0)
                {
                    IGR_SMALL_BIZ_REDUTION.SetCellValue("AGE", e.CellValue);
                }
            }
            else if (e.ColIndex == IGR_SMALL_BIZ_REDUTION.GetColumnToIndex("ARMY_PERIOD_FR"))
            {
                if (iString.ISNumtoZero(IGR_SMALL_BIZ_REDUTION.GetCellValue("JOIN_AGE"), 0) == 0)
                {
                    //병역제외한 나이 
                    IDC_GET_AGE_P.SetCommandParamValue("P_PERSON_ID", IGR_SMALL_BIZ_REDUTION.GetCellValue("PERSON_ID"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ORI_JOIN_DATE", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_FR", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_FR"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_TO", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_TO"));
                    IDC_GET_AGE_P.ExecuteNonQuery();
                    IGR_SMALL_BIZ_REDUTION.SetCellValue("AGE", IDC_GET_AGE_P.GetCommandParamValue("O_AGE"));
                }
            }
            else if (e.ColIndex == IGR_SMALL_BIZ_REDUTION.GetColumnToIndex("ARMY_PERIOD_TO"))
            {
                if (iString.ISNumtoZero(IGR_SMALL_BIZ_REDUTION.GetCellValue("JOIN_AGE"), 0) == 0)
                {
                    //병역제외한 나이 
                    IDC_GET_AGE_P.SetCommandParamValue("P_PERSON_ID", IGR_SMALL_BIZ_REDUTION.GetCellValue("PERSON_ID"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ORI_JOIN_DATE", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_FR", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_FR"));
                    IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_TO", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_TO"));
                    IDC_GET_AGE_P.ExecuteNonQuery();
                    IGR_SMALL_BIZ_REDUTION.SetCellValue("AGE", IDC_GET_AGE_P.GetCommandParamValue("O_AGE"));
                }
            }
            else if (e.ColIndex == IGR_SMALL_BIZ_REDUTION.GetColumnToIndex("REDUTION_DATE_FR"))
            {
                IDC_GET_REDUTION_DATE_TO_P.SetCommandParamValue("P_REDUTION_DATE_FR", IGR_SMALL_BIZ_REDUTION.GetCellValue("REDUTION_DATE_FR"));
                IDC_GET_REDUTION_DATE_TO_P.ExecuteNonQuery();
                IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_DATE_TO", IDC_GET_REDUTION_DATE_TO_P.GetCommandParamValue("O_REDUTION_DATE_TO"));
            }
        }


        #endregion;

        #region ----- Lookup Event -----

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_DEPT_ID", W_DEPT_ID.EditValue);
            ildPERSON.SetLookupParamValue("W_FLOOR_ID", W_FLOOR_ID.EditValue); 
        }

        private void ILA_PERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_DEPT_ID", DBNull.Value);
            ildPERSON.SetLookupParamValue("W_FLOOR_ID", DBNull.Value); 
        }

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_EMPLOYE_TYPE", "Y");
        }

        private void ILA_SMALL_BIZ_JOIN_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("SMALL_BIZ_JOIN_TYPE", "Y");
        }

        private void ILA_SMALL_BIZ_JOIN_TYPE_SelectedRowData(object pSender)
        {
            IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_RATE_CODE", null);
            IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_RATE_NAME", null);
            IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_RATE", null);
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("FLOOR", "Y");
        }

        private void ILA_PERSON_SelectedRowData(object pSender)
        {
            if (iString.ISNull(IGR_SMALL_BIZ_REDUTION.GetCellValue("REDUTION_DATE_FR")) == string.Empty)
            {
                IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_DATE_FR", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));

                IDC_GET_REDUTION_DATE_TO_P.SetCommandParamValue("P_REDUTION_DATE_FR", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));
                IDC_GET_REDUTION_DATE_TO_P.ExecuteNonQuery();
                IGR_SMALL_BIZ_REDUTION.SetCellValue("REDUTION_DATE_TO", IDC_GET_REDUTION_DATE_TO_P.GetCommandParamValue("O_REDUTION_DATE_TO"));
            }

            if (iString.ISNumtoZero(IGR_SMALL_BIZ_REDUTION.GetCellValue("JOIN_AGE"), 0) == 0)
            {
                //취업시 나이
                IDC_GET_AGE_P.SetCommandParamValue("P_PERSON_ID", IGR_SMALL_BIZ_REDUTION.GetCellValue("PERSON_ID"));
                IDC_GET_AGE_P.SetCommandParamValue("P_ORI_JOIN_DATE", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));
                IDC_GET_AGE_P.ExecuteNonQuery();
                IGR_SMALL_BIZ_REDUTION.SetCellValue("JOIN_AGE", IDC_GET_AGE_P.GetCommandParamValue("O_AGE"));
            }
            if (iString.ISNumtoZero(IGR_SMALL_BIZ_REDUTION.GetCellValue("AGE"), 0) == 0)
            {
                //병역제외한 나이 
                IDC_GET_AGE_P.SetCommandParamValue("P_PERSON_ID", IGR_SMALL_BIZ_REDUTION.GetCellValue("PERSON_ID"));
                IDC_GET_AGE_P.SetCommandParamValue("P_ORI_JOIN_DATE", IGR_SMALL_BIZ_REDUTION.GetCellValue("ORI_JOIN_DATE"));
                IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_FR", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_FR"));
                IDC_GET_AGE_P.SetCommandParamValue("P_ARMY_PERIOD_TO", IGR_SMALL_BIZ_REDUTION.GetCellValue("ARMY_PERIOD_TO"));
                IDC_GET_AGE_P.ExecuteNonQuery();
                IGR_SMALL_BIZ_REDUTION.SetCellValue("AGE", IDC_GET_AGE_P.GetCommandParamValue("O_AGE"));
            }
        }

        private void ILA_REDUTION_RATE_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_REDUTION_RATE_CODE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaPERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=[Person No]"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ORI_JOIN_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[최초입사일]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["JOIN_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[취업유형]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["JOIN_AGE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[취업시 나이]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;  
            }
            if (iString.ISNull(e.Row["REDUTION_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show("[감면기간 :: 시작일자]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REDUTION_DATE_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show("[감면기간 :: 종료일자]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}