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

namespace HRMF0346
{
    public partial class HRMF0346 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDateTime = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0346(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime( END_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }

            idaHOLY_TYPE.Fill();
            igrHOLY_TYPE.Focus();
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pStart_Date, Object pEnd_Date)
        {            
            idaWORK_CALENDAR.SetSelectParamValue("W_PERSON_ID", pPerson_ID);
            idaWORK_CALENDAR.SetSelectParamValue("W_START_DATE", pStart_Date);
            idaWORK_CALENDAR.SetSelectParamValue("W_END_DATE", pEnd_Date);
            idaWORK_CALENDAR.Fill();

            if (iString.ISNull(pStart_Date) == string.Empty)
            {
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_START_YEAR", iDateTime.ISYear(DateTime.Today));                
            }
            else
            {
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_START_YEAR", iDateTime.ISYear(Convert.ToDateTime(pStart_Date)));
            }
            if (iString.ISNull(pEnd_Date) == string.Empty)
            {
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_END_YEAR", iDateTime.ISYear(DateTime.Today));
            }
            else
            {                
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_END_YEAR", iDateTime.ISYear(Convert.ToDateTime(pEnd_Date)));
            }
            idaHOLIDAY_MANAGEMENT.Fill();
        }

        private bool isAdd_DB_Check()
        {// 데이터 추가시 검증.
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return false;
            }
            return true;
        }

        #endregion;

        //-- Report 관련 코드
        #region ----- XL Export Methods ----

        private void ExportXL(ISDataAdapter pAdapter)
        {
            int vCountRow = pAdapter.OraSelectData.Rows.Count; // (1)
            if (vCountRow < 1)
            {
                return;
            }

            string vsMessage = string.Empty;
            string vsSheetName = "Slip_Line";

            saveFileDialog1.Title = "Excel_Save";
            saveFileDialog1.FileName = "XL_00";
            saveFileDialog1.DefaultExt = "xls";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string vsSaveExcelFileName = saveFileDialog1.FileName;
                XL.XLPrint xlExport = new XL.XLPrint();
                bool vXLSaveOK = xlExport.XLExport(pAdapter.OraSelectData, vsSaveExcelFileName, vsSheetName); //
                if (vXLSaveOK == true)
                {
                    vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                else
                {
                    vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                xlExport.XLClose();
            }
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        #endregion;

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1()
        {
            string vMessageText = string.Empty;

            XLPrinting xlPrinting = new XLPrinting();

            try
            {
                //-------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0346_001.xls";
                xlPrinting.XLFileOpen();

                int vTerritory = GetTerritory(igrHOLY_TYPE.TerritoryLanguage);
                string vPeriodFrom = START_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);
                string vPeriodTo = END_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);

                string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);

                int viCutStart = this.Text.LastIndexOf("]") + 1;
                string vCaption = this.Text.Substring(0, viCutStart);
                int vPageNumber = xlPrinting.XLWirte(igrHOLY_TYPE, vTerritory, vPeriodFrom, vPeriodTo, vUserName, vCaption);

                xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                xlPrinting.Printing(3, 4);


                xlPrinting.Save("Cashier_"); //저장 파일명

                xlPrinting.PreView();

                xlPrinting.Dispose();
                //-------------------------------------------------------------------------

                vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
        }

        #endregion;

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
                    if(idaHOLY_TYPE.IsFocused)
                    {
                        if (isAdd_DB_Check() == false)
                        {
                            return;
                        }

                        idaHOLY_TYPE.AddOver();

                        igrHOLY_TYPE.SetCellValue("START_DATE", DateTime.Today.Date);
                        igrHOLY_TYPE.SetCellValue("END_DATE", DateTime.Today.Date);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {
                        if (isAdd_DB_Check() == false)
                        {
                            return;
                        }
                        idaHOLY_TYPE.AddUnder();

                        igrHOLY_TYPE.SetCellValue("START_DATE", DateTime.Today.Date);
                        igrHOLY_TYPE.SetCellValue("END_DATE", DateTime.Today.Date);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {
                        idaHOLY_TYPE.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {
                        idaHOLY_TYPE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {
                        idaHOLY_TYPE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print) //인쇄버튼
                {
                    XLPrinting1();
                }
                /*else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export) //엑셀파일 버튼
                {
                    if (idaDUTY_PERIOD.IsFocused == true) // 어뎁터가 하나 이상일 경우 else if문으로 사용
                    {
                        ExportXL(idaDUTY_PERIOD);
                    }
                }*/
            }
        }
        #endregion;

        #region ----- Form Event -----
        private void HRMF0346_Load(object sender, EventArgs e)
        {
            idaHOLY_TYPE.FillSchema();
            START_DATE_0.EditValue = DateTime.Today.AddDays(-7);
            END_DATE_0.EditValue = DateTime.Today.AddDays(7);

            DefaultCorporation();

            //APPROVE STATUS LOOKUP SETTING
            ildAPPROVE_STATUS.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            ildAPPROVE_STATUS.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
            
            // LOOKUP DEFAULT VALUE SETTING - APPROVE STATUS
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            APPROVE_STATUS_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            APPROVE_STATUS_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }

        #endregion  

        #region ----- Adapter Event -----
        private void idaHOLY_TYPE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(igrHOLY_TYPE.GetCellValue("PERSON_ID"), igrHOLY_TYPE.GetCellValue("START_DATE"), igrHOLY_TYPE.GetCellValue("END_DATE"));
        }

        private void idaHOLY_TYPE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if(e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=사원 정보"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["START_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=시작일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["END_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=종료일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Convert.ToDateTime(e.Row["START_DATE"]) > Convert.ToDateTime(e.Row["END_DATE"]))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaHOLY_TYPE_PreDelete(ISPreDeleteEventArgs e)
        {            
        }

        #endregion

        #region ----- LookUp Event -----
        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ildHOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_SelectedRowData(object pSender)
        {
            isSearch_WorkCalendar(igrHOLY_TYPE.GetCellValue("PERSON_ID"), igrHOLY_TYPE.GetCellValue("START_DATE"), igrHOLY_TYPE.GetCellValue("END_DATE"));
        }
        
        #endregion

        private void igrHOLY_TYPE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            object mStart_Date = igrHOLY_TYPE.GetCellValue("START_DATE");
            object mEnd_Date = igrHOLY_TYPE.GetCellValue("END_DATE");
            if (e.ColIndex == igrHOLY_TYPE.GetColumnToIndex("START_DATE") || e.ColIndex == igrHOLY_TYPE.GetColumnToIndex("END_DATE"))
            {
                if (e.ColIndex == igrHOLY_TYPE.GetColumnToIndex("START_DATE"))
                {
                    mStart_Date = e.NewValue;
                    mEnd_Date = mStart_Date;
                    igrHOLY_TYPE.SetCellValue("END_DATE", mEnd_Date);
                }
                if (e.ColIndex == igrHOLY_TYPE.GetColumnToIndex("END_DATE"))
                {
                    mEnd_Date = e.NewValue;
                }
                isSearch_WorkCalendar(igrHOLY_TYPE.GetCellValue("PERSON_ID"), mStart_Date, mEnd_Date);
            }
        }

        private void btnAPPR_REQUEST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaHOLY_TYPE.Update();

            int mRowCount = igrHOLY_TYPE.RowCount;
            for (int R = 0; R < mRowCount; R++)
            {
                if (iString.ISNull(igrHOLY_TYPE.GetCellValue(R, igrHOLY_TYPE.GetColumnToIndex("APPROVE_STATUS"))) == "N".ToString())
                {// 승인미요청 건에 대해서 승인 처리.
                    idcAPPROVAL_REQUEST.SetCommandParamValue("W_HOLY_TYPE_ID", igrHOLY_TYPE.GetCellValue(R, igrHOLY_TYPE.GetColumnToIndex("HOLY_TYPE_ID")));
                    idcAPPROVAL_REQUEST.ExecuteNonQuery();

                    object mValue;
                    mValue = idcAPPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS");
                    igrHOLY_TYPE.SetCellValue(R, igrHOLY_TYPE.GetColumnToIndex("APPROVE_STATUS"), mValue);
                    mValue = idcAPPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                    igrHOLY_TYPE.SetCellValue(R, igrHOLY_TYPE.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                }
            }

            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", "A");
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "HOLY");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", CORP_ID_0.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", DateTime.Today);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery();

            idaHOLY_TYPE.OraSelectData.AcceptChanges();
            idaHOLY_TYPE.Refillable = true;

            Search_DB();
        }
    }
}