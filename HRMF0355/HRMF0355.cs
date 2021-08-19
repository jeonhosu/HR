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

namespace HRMF0355
{
    public partial class HRMF0355 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0355()
        {
            InitializeComponent();
        }

        public HRMF0355(Form pMainForm, ISAppInterface pAppInterface)
        {
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
         
        #region ----- Private Methods ----
        
        private void DefaultCorporation()
        {
            ildCORP_0.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
            G_CORP_GROUP.BringToFront();
            //CORP TYPE :: 전체이면 그룹박스 표시, 
            if (iConv.ISNull(G_CORP_TYPE.EditValue, "1") == "1")
            {
                G_CORP_GROUP.Visible = false; //.Show();
                V_RB_OWNER.CheckedState = ISUtil.Enum.CheckedState.Checked;
                G_CORP_TYPE.EditValue = V_RB_OWNER.RadioCheckedString;
            }
            else
            {
                G_CORP_GROUP.Visible = true; //.Show();
                if (iConv.ISNull(G_CORP_TYPE.EditValue) == "ALL")
                {
                    V_RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ALL.RadioCheckedString;
                }
                else
                {
                    V_RB_ETC.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ETC.RadioCheckedString;
                }
            }
        }

        private void SEARCH_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(WORK_DATE_0.EditValue.ToString()))
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            idaDAY_INTERFACE_PRINT.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            idaDAY_INTERFACE_PRINT.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            idaDAY_INTERFACE_PRINT.Fill();
            igrDAY_INTERFACE_PRINT.Focus();
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = -1;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 0;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 5;
                    break;
            }

            return vTerritory;
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOutChoice, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;
            object vConv_Work_Date = null;  // 근무일자(년월일)
            object vWeek_Desc = null;       // 요일(요일).

            int vCountRow = pGrid.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSavePath = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("WorkData_{0}", WORK_DATE_0.DateTimeValue.ToShortDateString());

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xlsx";

                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    if (vFileName.Exists)
                    {
                        try
                        {
                            vFileName.Delete();
                        }
                        catch (Exception ex)
                        {
                            MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            vMessageText = string.Format(" Writing Start...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0355_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    // 헤더 인쇄.
                    idcGET_WORK_DATE.SetCommandParamValue("P_FORMAT", "MMDD");
                    idcGET_WORK_DATE.ExecuteNonQuery();
                    vConv_Work_Date = idcGET_WORK_DATE.GetCommandParamValue("O_CONV_DATE");
                    vWeek_Desc = idcGET_WORK_DATE.GetCommandParamValue("O_WEEK_DESC");

                    if (iConv.ISNull(vConv_Work_Date) == string.Empty || iConv.ISNull(vWeek_Desc) == string.Empty)
                    {
                        vMessageText = string.Format(" Writing End...Data Not Found...");
                        isAppInterfaceAdv1.OnAppMessage(vMessageText);
                        System.Windows.Forms.Application.DoEvents();
                        return;
                    }

                    // 실제 인쇄
                    vPageNumber = xlPrinting.LineWrite(pGrid, vConv_Work_Date, vWeek_Desc);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Events -----

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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaDAY_INTERFACE_PRINT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    idaDAY_INTERFACE_PRINT.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT", igrDAY_INTERFACE_PRINT);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE", igrDAY_INTERFACE_PRINT);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0355_Load(object sender, EventArgs e)
        {
            WORK_DATE_0.EditValue = DateTime.Today;
            DefaultCorporation();

            idaDAY_INTERFACE_PRINT.FillSchema();            
        }

        private void HRMF0355_Shown(object sender, EventArgs e)
        {
           
        }

        #endregion

        #region ----- Lookup Event ------

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_OPERATING_UNIT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaJOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e) //직구분
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
         
        #endregion

        #region ----- Adapter Event -----


        #endregion

    }
}