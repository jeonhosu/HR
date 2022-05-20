using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.IO;
using Syncfusion.GridExcelConverter;
using Syncfusion.XlsIO;

namespace HRMF0313
{
    public partial class HRMF0313 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0313(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)   //파견직관리
            {
                G_CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
        }

        #endregion;

        #region ----- Corp Type -----

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

        #endregion

        #region ----- Private Methods -----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            
            CORP_NAME_0.BringToFront();
            G_CORP_GROUP.BringToFront();
            G_CORP_GROUP.Visible = false;

            if (iString.ISNull(G_CORP_TYPE.EditValue) == "ALL")
            {
                G_CORP_GROUP.Visible = true; //.Show(); 
                V_RB_ALL.RadioButtonValue = "A";
                G_CORP_TYPE.EditValue = "A";

            }
            else if (iString.ISNull(G_CORP_TYPE.EditValue) == "1")
            {
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null&& G_CORP_TYPE.EditValue.ToString() == "1")
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            icb_SELECT_YN.CheckBoxValue = "N";
            isAppInterfaceAdv1.OnAppMessage("");
            igrDAY_LEAVE.LastConfirmChanges();
            idaDAY_LEAVE.OraSelectData.AcceptChanges();
            idaDAY_LEAVE.Refillable = true;

            idaDAY_LEAVE.Fill();
            igrDAY_LEAVE.Focus();
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pWork_Date)
        {            
            if (iConv.ISNull(pWork_Date) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = pWork_Date;

            idaWORK_CALENDAR.SetSelectParamValue("W_END_DATE", pWork_Date);
            idaDAY_HISTORY.Fill();
            idaDUTY_PERIOD.Fill();
            idaWORK_CALENDAR.Fill();
        }

        private void isSearch_Day_History(int pAdd_Day)
        {
            if (iConv.ISNull(WORK_DATE_8.EditValue) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = Convert.ToDateTime(WORK_DATE_8.EditValue).AddDays(pAdd_Day);
            idaDAY_HISTORY.Fill();
        }

        private bool Check_Holy_Type(object pWork_Date, object pDuty_ID, object pHoly_Type, object pOpen_Time, object pClose_Time)
        {     
            bool mCheck_Value = false;

            if (iConv.ISNull(pWork_Date) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iConv.ISNull(pDuty_ID) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iConv.ISNull(pHoly_Type) == string.Empty)
            {
                return (mCheck_Value);
            }

            idcHOLY_TYPE_CHECK_P.SetCommandParamValue("W_WORK_DATE", pWork_Date);
            idcHOLY_TYPE_CHECK_P.SetCommandParamValue("W_DUTY_ID", pDuty_ID);
            idcHOLY_TYPE_CHECK_P.SetCommandParamValue("W_HOLY_TYPE", pHoly_Type);
            idcHOLY_TYPE_CHECK_P.SetCommandParamValue("W_OPEN_TIME", pOpen_Time);
            idcHOLY_TYPE_CHECK_P.SetCommandParamValue("W_CLOSE_TIME", pClose_Time);
            idcHOLY_TYPE_CHECK_P.ExecuteNonQuery();
            string vStatus = iConv.ISNull(idcHOLY_TYPE_CHECK_P.GetCommandParamValue("O_STATUS"));
            string vMessage = iConv.ISNull(idcHOLY_TYPE_CHECK_P.GetCommandParamValue("O_MESSAGE"));
            if (idcHOLY_TYPE_CHECK_P.ExcuteError || vStatus == "F")
            {
                mCheck_Value = false;
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                mCheck_Value = true;
            }
            return (mCheck_Value);
        }

        private bool Check_Work_Date_time(object pHoly_Type, object IO_Flag, object pWork_Date, object pNew_Work_Date)
        {
            bool mCheck_Value = false;

            if (iConv.ISNull(pHoly_Type) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iConv.ISNull(IO_Flag) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iConv.ISNull(pWork_Date) == string.Empty)
            {
                return (mCheck_Value); 
            }
            if (iConv.ISNull(pNew_Work_Date) == string.Empty)
            {
                return true;
            }

            if ((pHoly_Type.ToString() == "0".ToString() || pHoly_Type.ToString() == "1".ToString() || pHoly_Type.ToString() == "2".ToString()
                || pHoly_Type.ToString() == "D".ToString() || pHoly_Type.ToString() == "S".ToString())
                && IO_Flag.ToString() == "IN".ToString())
            {// 주간, 무휴, 유휴, DAY, SWING --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date == Convert.ToDateTime(pNew_Work_Date).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "3".ToString() || pHoly_Type.ToString() == "N".ToString() || pHoly_Type.ToString() == "G".ToString()
                  || pHoly_Type.ToString() == "Y".ToString())
                && IO_Flag.ToString() == "IN".ToString())
            {// 주간, 야간, 무휴, 유휴, DAY, NIGHT --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                    && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "0".ToString() || pHoly_Type.ToString() == "1".ToString() || pHoly_Type.ToString() == "2".ToString()
         || pHoly_Type.ToString() == "D".ToString() || pHoly_Type.ToString() == "S".ToString())
              && IO_Flag.ToString() == "OUT".ToString())
            {// 주간, 무휴, 유휴, DAY, SWING --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                    && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "3".ToString() || pHoly_Type.ToString() == "N".ToString() || pHoly_Type.ToString() == "G".ToString()
                  || pHoly_Type.ToString() == "Y".ToString())
           && IO_Flag.ToString() == "OUT".ToString())
            {// 주간, 야간, 무휴, 유휴, DAY, NIGHT --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                    && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            return (mCheck_Value);
        }

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                //xls 저장방법
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 파일 저장 방법
                GridExcelConverterControl converter = new GridExcelConverterControl();
                ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2010;
                IWorkbook workBook = ExcelUtils.CreateWorkbook(1);
                workBook.Version = ExcelVersion.Excel2010;
                IWorksheet sheet = workBook.Worksheets[0];
                //used to convert grid to excel 
                converter.GridToExcel(pGrid.BaseGrid, sheet, ConverterOptions.ColumnHeaders);
                //used to save the file
                workBook.SaveAs(saveFileDialog.FileName);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                        "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----      
          
        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaDAY_LEAVE.Update();  
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaDAY_LEAVE.IsFocused)
                    {
                        idaDAY_LEAVE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaDAY_LEAVE.IsFocused)
                    {
                        idaDAY_LEAVE.Delete();
                    }
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(igrDAY_LEAVE);
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0313_Load(object sender, EventArgs e)
        {
            this.Visible = true;

            START_DATE_0.EditValue = DateTime.Today;
            END_DATE_0.EditValue = DateTime.Today;
            
            // CORP SETTING
            DefaultCorporation();

            // LEAVE CLOSE TYPE SETTING
            ildLEAVE_CLOSE_TYPE_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ildLEAVE_CLOSE_TYPE_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            LEAVE_CLOSE_TYPE_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            LEAVE_CLOSE_TYPE_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();
             
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
            idaDAY_LEAVE.FillSchema();

        }

        private void START_DATE_0_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            END_DATE_0.EditValue = e.EditValue;
        }
         
        private void ibtLEAVE_DATETIME_UPDATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null && G_CORP_TYPE.EditValue.ToString() =="1")
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 근무 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 근무 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mStatus = "F";
            string mMessage = string.Empty;

            idcLEAVE_DATETIME_UPDATE.ExecuteNonQuery();
            mStatus= idcLEAVE_DATETIME_UPDATE.GetCommandParamValue("O_STATUS").ToString();
            mMessage = iConv.ISNull(idcLEAVE_DATETIME_UPDATE.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (idcLEAVE_DATETIME_UPDATE.ExcuteError || mStatus == "F")
            {
                if (mMessage != string.Empty)
                {
                    MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // refill.
            Search_DB();
        }
        
        private void ibtSET_OT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null&& G_CORP_TYPE.EditValue.ToString() == "1")

            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 근무 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 근무 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents(); 

            string mStatus = "F";
            string mMessage = string.Empty;
            idcSET_OT.ExecuteNonQuery();
            mStatus = iConv.ISNull(idcSET_OT.GetCommandParamValue("O_STATUS"));
            mMessage = iConv.ISNull(idcSET_OT.GetCommandParamValue("O_MESSAGE"));
            if (idcSET_OT.ExcuteError || mStatus == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                if (mMessage != string.Empty)
                {
                    MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (mMessage != string.Empty)
            {
                MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // refill.
            Search_DB();
        }

        private void ibtCLOSE_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (igrDAY_LEAVE.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor; 
            Application.DoEvents();
             
            igrDAY_LEAVE.LastConfirmChanges();
            idaDAY_LEAVE.OraSelectData.AcceptChanges();
            idaDAY_LEAVE.Refillable = true;

            int vIDX_SELECT_YN = igrDAY_LEAVE.GetColumnToIndex("SELECT_YN");
            int vIDX_WORK_DATE = igrDAY_LEAVE.GetColumnToIndex("WORK_DATE");
            int vIDX_PERSON_ID = igrDAY_LEAVE.GetColumnToIndex("PERSON_ID");

            string mStatus = "F";
            string mMessage = string.Empty;

            for (int i = 0; i < igrDAY_LEAVE.RowCount; i++)
            {
                if (iConv.ISNull(igrDAY_LEAVE.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                {
                    idcDATA_CLOSE_PROC.SetCommandParamValue("W_WORK_DATE", igrDAY_LEAVE.GetCellValue(i, vIDX_WORK_DATE));
                    idcDATA_CLOSE_PROC.SetCommandParamValue("W_PERSON_ID", igrDAY_LEAVE.GetCellValue(i, vIDX_PERSON_ID));
                    idcDATA_CLOSE_PROC.ExecuteNonQuery();
                    mStatus = idcDATA_CLOSE_PROC.GetCommandParamValue("O_STATUS").ToString();
                    mMessage = iConv.ISNull(idcDATA_CLOSE_PROC.GetCommandParamValue("O_MESSAGE"));
                    
                    Application.DoEvents();

                    if (mStatus == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        if (mMessage != string.Empty)
                        {
                            MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    } 
                } 
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // refill.
            Search_DB();
        }

        private void ibtCLOSE_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (igrDAY_LEAVE.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();


            igrDAY_LEAVE.LastConfirmChanges();
            idaDAY_LEAVE.OraSelectData.AcceptChanges();
            idaDAY_LEAVE.Refillable = true;

            int vIDX_SELECT_YN = igrDAY_LEAVE.GetColumnToIndex("SELECT_YN");
            int vIDX_WORK_DATE = igrDAY_LEAVE.GetColumnToIndex("WORK_DATE");
            int vIDX_PERSON_ID = igrDAY_LEAVE.GetColumnToIndex("PERSON_ID");

            string mStatus = "F";
            string mMessage = string.Empty;
            for (int i = 0; i < igrDAY_LEAVE.RowCount; i++)
            {
                if (iConv.ISNull(igrDAY_LEAVE.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                {
                    idcDATA_CLOSE_CANCEL.SetCommandParamValue("W_WORK_DATE", igrDAY_LEAVE.GetCellValue(i, vIDX_WORK_DATE));
                    idcDATA_CLOSE_CANCEL.SetCommandParamValue("W_PERSON_ID", igrDAY_LEAVE.GetCellValue(i, vIDX_PERSON_ID));

                    idcDATA_CLOSE_CANCEL.ExecuteNonQuery();
                    mStatus = idcDATA_CLOSE_CANCEL.GetCommandParamValue("O_STATUS").ToString();
                    mMessage = iConv.ISNull(idcDATA_CLOSE_CANCEL.GetCommandParamValue("O_MESSAGE"));
                    Application.DoEvents();

                    if (idcDATA_CLOSE_CANCEL.ExcuteError || mStatus == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();
                        if (mMessage != string.Empty)
                        {
                            MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // refill.
            Search_DB();
        }

        private void SET_EXCEL_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0313_UPLOAD vHRMF0313_UPLOAD = new HRMF0313_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface, CORP_ID_0.EditValue);
            vdlgResult = vHRMF0313_UPLOAD.ShowDialog();
            vHRMF0313_UPLOAD.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void ibtnUP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(1);
        }

        private void ibtnDOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(-1);
        }

        private void igrDAY_LEAVE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {            
            if (e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("OPEN_TIME") || e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("OPEN_TIME1")
                || e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("CLOSE_TIME") || e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("CLOSE_TIME1"))
            {
                object mHoly_Type = igrDAY_LEAVE.GetCellValue("HOLY_TYPE");
                object mWork_Date = igrDAY_LEAVE.GetCellValue("WORK_DATE");
                object mWork_DateTime = e.NewValue;
                object mIO_Flag = "-";
                if (e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("OPEN_TIME") || e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("OPEN_TIME1"))
                {
                    mIO_Flag = "IN";
                }
                else if (e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("CLOSE_TIME") || e.ColIndex == igrDAY_LEAVE.GetColumnToIndex("CLOSE_TIME1"))
                {
                    mIO_Flag = "OUT";
                }
                if (Check_Work_Date_time(mHoly_Type, mIO_Flag, mWork_Date, mWork_DateTime) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
        }

        private void igrDAY_LEAVE_CellDoubleClick(object pSender)
        {
            string vCOL_NAME = null;
            if (igrDAY_LEAVE.GetColumnToIndex("OPEN_TIME") == igrDAY_LEAVE.ColIndex)
            {
                vCOL_NAME = "OPEN_TIME";
            }
            else if(igrDAY_LEAVE.GetColumnToIndex("CLOSE_TIME") == igrDAY_LEAVE.ColIndex)
            {
                vCOL_NAME = "CLOSE_TIME";
            }

            if (igrDAY_LEAVE.GetColumnToIndex(vCOL_NAME) == igrDAY_LEAVE.ColIndex)
            {
                if (iConv.ISNull(igrDAY_LEAVE.GetCellValue(vCOL_NAME)) == string.Empty)
                {
                    idcWORK_IO_TIME_P.SetCommandParamValue("W_WORK_TYPE", igrDAY_LEAVE.GetCellValue("WORK_TYPE_GROUP"));
                    idcWORK_IO_TIME_P.SetCommandParamValue("W_HOLY_TYPE", igrDAY_LEAVE.GetCellValue("HOLY_TYPE"));
                    idcWORK_IO_TIME_P.SetCommandParamValue("W_WORK_DATE", igrDAY_LEAVE.GetCellValue("WORK_DATE"));
                    idcWORK_IO_TIME_P.ExecuteNonQuery();
                    if (vCOL_NAME == "OPEN_TIME")
                    {//출근
                        igrDAY_LEAVE.SetCellValue("HOLY_TYPE", idcWORK_IO_TIME_P.GetCommandParamValue("O_HOLY_TYPE"));
                        igrDAY_LEAVE.SetCellValue("HOLY_TYPE_NAME", idcWORK_IO_TIME_P.GetCommandParamValue("O_HOLY_TYPE_NAME"));
                        igrDAY_LEAVE.SetCellValue(vCOL_NAME, idcWORK_IO_TIME_P.GetCommandParamValue("O_OPEN_TIME"));
                    }
                    else if (vCOL_NAME == "CLOSE_TIME")
                    {//퇴근
                        igrDAY_LEAVE.SetCellValue(vCOL_NAME, idcWORK_IO_TIME_P.GetCommandParamValue("O_CLOSE_TIME"));
                    }
                }
            }
        }
       
        #endregion  

        #region ----- Adapter Event -----

        private void idaDAY_LEAVE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Info(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WORK_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CORP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporate Name(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DUTY_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty Name(근태)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["HOLY_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Holy type(근무)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (Check_Holy_Type(e.Row["WORK_DATE"], e.Row["DUTY_ID"], e.Row["HOLY_TYPE"], e.Row["OPEN_TIME"], e.Row["CLOSE_TIME"]) == false)
            {
                e.Cancel = true;
                return;
            }

            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], "IN", e.Row["WORK_DATE"], e.Row["OPEN_TIME"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], "IN", e.Row["WORK_DATE"], e.Row["OPEN_TIME1"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], "OUT", e.Row["WORK_DATE"], e.Row["CLOSE_TIME"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], "OUT", e.Row["WORK_DATE"], e.Row["CLOSE_TIME1"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDAY_LEAVE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(igrDAY_LEAVE.GetCellValue("PERSON_ID"), igrDAY_LEAVE.GetCellValue("WORK_DATE"));
        }

        #endregion

        #region ----- LookUp Event -----

        private void ilaOPERATING_UNIT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", CORP_ID_0.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ildHOLY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE",END_DATE_0.EditValue);
        }

        private void ilaDUTY_MODIFY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaJOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaHOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_BREAKFAST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_LUNCH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_DINNER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_MIDNIGHT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "Y");
        }

        #endregion

        private void icb_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            for (int r = 0; r < igrDAY_LEAVE.RowCount; r++)
            {
                igrDAY_LEAVE.SetCellValue(r, igrDAY_LEAVE.GetColumnToIndex("SELECT_YN"), icb_SELECT_YN.CheckBoxString);
            }
            igrDAY_LEAVE.LastConfirmChanges();
            idaDAY_LEAVE.OraSelectData.AcceptChanges();
            idaDAY_LEAVE.Refillable = true;
        }
    }
}