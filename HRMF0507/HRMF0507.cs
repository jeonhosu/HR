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
using Syncfusion.XlsIO;

namespace HRMF0507
{
    public partial class HRMF0507 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private string mCompany = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public HRMF0507()
        {
            InitializeComponent();
        }

        public HRMF0507(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void DefaultCorporation()
        {
            try
            {
                // Lookup SETTING
                ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
                idcDEFAULT_CORP.ExecuteNonQuery();
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

                CORP_NAME_0.BringToFront();
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }
             
            if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_DETAIL.TabIndex)
            {
                string vPERSON_NUM = string.Empty;
                if (IGR_PAYMENT_VIEW_DTL.RowIndex < 0)
                {
                    vPERSON_NUM = string.Empty;
                }
                else
                {
                    vPERSON_NUM = iString.ISNull(IGR_PAYMENT_VIEW_DTL.GetCellValue("PERSON_NUM"));
                }

                IDA_PAYMENT_VIEW_DTL.SetSelectParamValue("P_SOB_ID", -1);
                IDA_PAYMENT_VIEW_DTL.Fill();
                INIT_COLUMN_0();

                IDA_PAYMENT_VIEW_DTL.SetSelectParamValue("P_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
                IDA_PAYMENT_VIEW_DTL.Fill(); 
                if (vPERSON_NUM == string.Empty)
                {
                    IGR_PAYMENT_VIEW_DTL.Focus();
                }
                else
                {
                    int vIDX_PERSON_NUM = IGR_PAYMENT_VIEW_DTL.GetColumnToIndex("PERSON_NUM");
                    for (int r = 0; r < IGR_PAYMENT_VIEW_DTL.RowCount; r++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_PAYMENT_VIEW_DTL.GetCellValue(r, vIDX_PERSON_NUM)))
                        {
                            IGR_PAYMENT_VIEW_DTL.CurrentCellMoveTo(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_VIEW_DTL.CurrentCellActivate(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_VIEW_DTL.Focus();
                            return;
                        }
                    }
                }
            }             
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM.TabIndex)
            {
                string vPERSON_NUM = string.Empty;
                if (IGR_PAYMENT_VIEW_SUM.RowIndex < 0)
                {
                    vPERSON_NUM = string.Empty;
                }
                else
                {
                    vPERSON_NUM = iString.ISNull(IGR_PAYMENT_VIEW_SUM.GetCellValue("PERSON_NUM"));
                }

                IDA_PAYMENT_VIEW_SUM.SetSelectParamValue("P_SOB_ID", -1);
                IDA_PAYMENT_VIEW_SUM.Fill();
                INIT_COLUMN_1();

                IDA_PAYMENT_VIEW_SUM.SetSelectParamValue("P_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
                IDA_PAYMENT_VIEW_SUM.Fill();  
                if (vPERSON_NUM == string.Empty)
                {
                    IGR_PAYMENT_VIEW_SUM.Focus();
                }
                else
                {
                    int vIDX_PERSON_NUM = IGR_PAYMENT_VIEW_SUM.GetColumnToIndex("PERSON_NUM");
                    for (int r = 0; r < IGR_PAYMENT_VIEW_SUM.RowCount; r++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_PAYMENT_VIEW_SUM.GetCellValue(r, vIDX_PERSON_NUM)))
                        {
                            IGR_PAYMENT_VIEW_SUM.CurrentCellMoveTo(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_VIEW_SUM.CurrentCellActivate(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_VIEW_SUM.Focus();
                            return;
                        }
                    }
                }
            } 
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_DEPT.TabIndex)
            {
                IDA_PAYMENT_VIEW_DEPT.SetSelectParamValue("P_SOB_ID", -1);
                IDA_PAYMENT_VIEW_DEPT.Fill();
                INIT_COLUMN_2();

                IDA_PAYMENT_VIEW_DEPT.SetSelectParamValue("P_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
                IDA_PAYMENT_VIEW_DEPT.Fill(); 
            }
        }

        private void INIT_COLUMN_0()
        {
            int mGRID_START_COL = 19;   // 그리드 시작 COLUMN.
            int mIDX_HEADER_ROW = 0;
            int mMax_Column = 53 + 42 + 54;       // 종료 COLUMN.(항목수) 

            //보이는 컬럼 초기화.
            for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            { 
                IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0; 
            }

            //지급//
            IDA_PROMT_VIEW_DTL_A.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_A.Fill(); 
            if (IDA_PROMT_VIEW_DTL_A.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            mGRID_START_COL = 19;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_A.OraSelectData.Rows.Count - 1); 
            mMax_Column = 53;       // 종료 COLUMN.(항목수) 
           
            string mCOLUMN_DESC;        // 헤더 프롬프트.

            foreach(DataRow vRow in IDA_PROMT_VIEW_DTL_A.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if(IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;                        
                    }
                    IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                    IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                }
                mIDX_HEADER_ROW--;
            }

            //공제 
            IDA_PROMT_VIEW_DTL_D.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_D.Fill();
            if (IDA_PROMT_VIEW_DTL_D.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            mGRID_START_COL = 72;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_D.OraSelectData.Rows.Count - 1); 
            mMax_Column = 42;       // 종료 COLUMN.(항목수)  
            mCOLUMN_DESC = "";        // 헤더 프롬프트.

            foreach (DataRow vRow in IDA_PROMT_VIEW_DTL_D.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if (IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    }
                    IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                    IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                }
                mIDX_HEADER_ROW--;
            }

            //추가// 
            IDA_PROMT_VIEW_DTL_W.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_W.Fill();
            if (IDA_PROMT_VIEW_DTL_W.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            mGRID_START_COL = 113;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_W.OraSelectData.Rows.Count - 1);
            mMax_Column = 54;       // 종료 COLUMN.(항목수)  
            mCOLUMN_DESC = "";        // 헤더 프롬프트.

            foreach (DataRow vRow in IDA_PROMT_VIEW_DTL_W.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if (IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    }
                    IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                    IGR_PAYMENT_VIEW_DTL.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                }
                mIDX_HEADER_ROW--;
            }

            //그리드 헤더 merge//
            int rowcnt = IGR_PAYMENT_VIEW_DTL.ColHeaderCount;
            int colcnt = IGR_PAYMENT_VIEW_DTL.ColCount;

            string value;
            IGR_PAYMENT_VIEW_DTL.ColHeaderMergeElement.Clear();
            for (int j = 0; j < rowcnt; j++)
            {
                for (int i = 0; i < colcnt; i++)
                {
                    value = IGR_PAYMENT_VIEW_DTL.BaseGrid[j, i].CellValue?.ToString();
                    int bottom = j;
                    int right = i;

                    if (string.IsNullOrEmpty(value) == false)
                    {
                        for (int x = i + 1; x < colcnt; x++)
                        {
                            value = IGR_PAYMENT_VIEW_DTL.BaseGrid[j, x].CellValue?.ToString();

                            // already?
                            foreach (ISGridAdvExRangeElement ele in IGR_PAYMENT_VIEW_DTL.ColHeaderMergeElement)
                            {
                                if (ele.Top <= j && j <= ele.Bottom &&
                                    ele.Left <= x && x <= ele.Right)
                                {
                                    value = "-1";
                                    break;
                                }
                            }

                            if (string.IsNullOrEmpty(value))
                                right = x;
                            else
                                break;
                        }

                        for (int y = j + 1; y < rowcnt; y++)
                        {
                            value = IGR_PAYMENT_VIEW_DTL.BaseGrid[y, i].CellValue?.ToString();

                            // already?
                            foreach (ISGridAdvExRangeElement ele in IGR_PAYMENT_VIEW_DTL.ColHeaderMergeElement)
                            {
                                if (ele.Top <= y && y <= ele.Bottom &&
                                    ele.Left <= i && i <= ele.Right)
                                {
                                    value = "-1";
                                    break;
                                }
                            }

                            if (string.IsNullOrEmpty(value))
                                bottom = y;
                            else
                                break;
                        }

                        IGR_PAYMENT_VIEW_DTL.ColHeaderMergeElement.Add(new ISGridAdvExRangeElement() { Top = j, Left = i, Bottom = bottom, Right = right });
                    }
                }
            }
             
            IGR_PAYMENT_VIEW_DTL.ResetDraw = true;
            IGR_PAYMENT_VIEW_DTL.Refresh();
        }

        private void INIT_COLUMN_1()
        {
            int mGRID_START_COL = 26;   // 그리드 시작 COLUMN.
            int mIDX_HEADER_ROW = 0;
            int mMax_Column = 41 + 54;       // 종료 COLUMN.(항목수) 
            string mCOLUMN_DESC;            // 헤더 프롬프트.

            //보이는 컬럼 초기화.
            for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
            }
             
            //공제 
            IDA_PROMT_VIEW_DTL_D1.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_D1.Fill();
            if (IDA_PROMT_VIEW_DTL_D1.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            mGRID_START_COL = 26;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_D1.OraSelectData.Rows.Count - 2);
            mMax_Column = 41;       // 종료 COLUMN.(항목수)  
            mCOLUMN_DESC = "";        // 헤더 프롬프트.

            foreach (DataRow vRow in IDA_PROMT_VIEW_DTL_D1.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if (IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    }
                    if (mIDX_HEADER_ROW >= 0)
                    {
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                    }
                }
                mIDX_HEADER_ROW--;
            }

            //추가// 
            IDA_PROMT_VIEW_DTL_W1.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_W1.Fill();
            if (IDA_PROMT_VIEW_DTL_W1.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            mGRID_START_COL = 67;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_W1.OraSelectData.Rows.Count - 2);
            mMax_Column = 54;       // 종료 COLUMN.(항목수)  
            mCOLUMN_DESC = "";        // 헤더 프롬프트.

            foreach (DataRow vRow in IDA_PROMT_VIEW_DTL_W1.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if (IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    }
                    if (mIDX_HEADER_ROW >= 0)
                    {
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                        IGR_PAYMENT_VIEW_SUM.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                    }
                }
                mIDX_HEADER_ROW--;
            }

            IGR_PAYMENT_VIEW_SUM.ResetDraw = true;
            IGR_PAYMENT_VIEW_SUM.Refresh();
        }

        private void INIT_COLUMN_2()
        {
            int mGRID_START_COL = 10;   // 그리드 시작 COLUMN.
            int mIDX_HEADER_ROW = 0;
            int mMax_Column = 53 + 41 + 44;       // 종료 COLUMN.(항목수) 

            //보이는 컬럼 초기화.
            for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
            }
            
            //지급//
            IDA_PROMT_VIEW_DTL_A2.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_A2.Fill();
            if (IDA_PROMT_VIEW_DTL_A2.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            mGRID_START_COL = 9;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_A2.OraSelectData.Rows.Count - 1);
            mMax_Column = 53;       // 종료 COLUMN.(항목수) 

            string mCOLUMN_DESC;        // 헤더 프롬프트.

            foreach (DataRow vRow in IDA_PROMT_VIEW_DTL_A2.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if (IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    }
                    IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                    IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                }
                mIDX_HEADER_ROW--;
            }

            //공제 
            IDA_PROMT_VIEW_DTL_D2.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_D2.Fill();
            if (IDA_PROMT_VIEW_DTL_D2.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            mGRID_START_COL = 62;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_D2.OraSelectData.Rows.Count - 1);
            mMax_Column = 41;       // 종료 COLUMN.(항목수)  
            mCOLUMN_DESC = "";        // 헤더 프롬프트.

            foreach (DataRow vRow in IDA_PROMT_VIEW_DTL_D2.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if (IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    }
                    IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                    IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                }
                mIDX_HEADER_ROW--;
            }

            //추가// 
            IDA_PROMT_VIEW_DTL_W2.SetSelectParamValue("W_STD_YYYYMM", PAY_YYYYMM_0.EditValue);
            IDA_PROMT_VIEW_DTL_W2.Fill();
            if (IDA_PROMT_VIEW_DTL_W2.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            mGRID_START_COL = 103;   // 그리드 시작 COLUMN.
            mIDX_HEADER_ROW = (IDA_PROMT_VIEW_DTL_W2.OraSelectData.Rows.Count - 1);
            mMax_Column = 44;       // 종료 COLUMN.(항목수)  
            mCOLUMN_DESC = "";        // 헤더 프롬프트.

            foreach (DataRow vRow in IDA_PROMT_VIEW_DTL_W2.CurrentRows)
            {
                for (int mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
                {
                    mCOLUMN_DESC = iString.ISNull(vRow[mIDX_Column]);
                    if (IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible.ToString() == "1")
                    {
                        //
                    }
                    else if (mCOLUMN_DESC == string.Empty)
                    {
                        IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                    }
                    else
                    {
                        IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    }
                    IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].Default = mCOLUMN_DESC;
                    IGR_PAYMENT_VIEW_DEPT.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[mIDX_HEADER_ROW].TL1_KR = mCOLUMN_DESC;
                }
                mIDX_HEADER_ROW--;
            }

            //그리드 헤더 merge//
            //int RowCnt = IGR_PAYMENT_SPREAD_DTL.ColHeaderCount;
            //int ColCnt = IGR_PAYMENT_SPREAD_DTL.ColCount;
            //string value;
            //IGR_PAYMENT_SPREAD_DTL.ColHeaderMergeElement.Clear();
            //for (int j = RowCnt - 1; j >= 0; j--)
            //{
            //    for (int i = 0; i < ColCnt; i++)
            //    {
            //        value = IGR_PAYMENT_SPREAD_DTL.BaseGrid[j, i].CellValue?.ToString();
            //        int bottom = j;
            //        int right = i;

            //        if (string.IsNullOrEmpty(value) == false)
            //        {
            //            for (int x = i + 1; x < ColCnt; x++)
            //            {
            //                value = IGR_PAYMENT_SPREAD_DTL.BaseGrid[j, x].CellValue?.ToString();

            //                // already?
            //                foreach (ISGridAdvExRangeElement ele in IGR_PAYMENT_SPREAD_DTL.ColHeaderMergeElement)
            //                {
            //                    if (ele.Top <= j && j <= ele.Bottom &&
            //                        ele.Left <= x && x <= ele.Right)
            //                    {
            //                        value = "-1";
            //                        break;
            //                    }
            //                }

            //                if (string.IsNullOrEmpty(value))
            //                    right = x;
            //                else
            //                    break;
            //            }

            //            //for (int y = j - 1; y >= 0; y--)
            //            //{
            //            //    value = IGR_PAYMENT_SPREAD_DTL.BaseGrid[y, i].CellValue?.ToString();

            //            //    // already?
            //            //    foreach (ISGridAdvExRangeElement ele in IGR_PAYMENT_SPREAD_DTL.ColHeaderMergeElement)
            //            //    {
            //            //        if (ele.Top <= y && y <= ele.Bottom &&
            //            //            ele.Left <= i && i <= ele.Right)
            //            //        {
            //            //            value = "-1";
            //            //            break;
            //            //        }
            //            //    }

            //            //    if (string.IsNullOrEmpty(value))
            //            //        bottom = y;
            //            //    else
            //            //        break;
            //            //}

            //            IGR_PAYMENT_SPREAD_DTL.ColHeaderMergeElement.Add(new ISGridAdvExRangeElement() { Top = j, Left = i, Bottom = bottom, Right = right });
            //        }
            //    }
            //} 

            IGR_PAYMENT_VIEW_DEPT.ResetDraw = true;
            IGR_PAYMENT_VIEW_DEPT.Refresh();
        }

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_Flag_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_Flag_YN);
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

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        //private void XLPrinting_1(string pOutChoice, ISDataAdapter pAdapter)
        //{// pOutChoice : 출력구분.
        //    string vMessageText = string.Empty;
        //    string vSaveFileName = string.Empty;

        //    object vToday = DateTime.Today.ToShortDateString();

        //    Application.UseWaitCursor = false;
        //    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        //    Application.DoEvents();

        //    //출력구분이 파일인 경우 처리.
        //    if (pOutChoice == "FILE")
        //    {
        //        System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        //        vSaveFileName = string.Format("Accounts_{0}", vToday);

        //        saveFileDialog1.Title = "Excel Save";
        //        saveFileDialog1.FileName = vSaveFileName;
        //        saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
        //        saveFileDialog1.DefaultExt = "xls";
        //        if (saveFileDialog1.ShowDialog() != DialogResult.OK)
        //        {
        //            return;
        //        }
        //        else
        //        {
        //            vSaveFileName = saveFileDialog1.FileName;
        //            System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
        //            try
        //            {
        //                if (vFileName.Exists)
        //                {
        //                    vFileName.Delete();
        //                }
        //            }
        //            catch (Exception EX)
        //            {
        //                MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                return;
        //            }
        //        }
        //        vMessageText = string.Format(" Writing Starting...");
        //    }
        //    else
        //    {
        //        vMessageText = string.Format(" Printing Starting...");
        //    }
        //    isAppInterfaceAdv1.OnAppMessage(vMessageText);
        //    Application.UseWaitCursor = true;
        //    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
        //    Application.DoEvents();

        //    int vPageNumber = 0;
        //    //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);
        //    XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

        //    try
        //    {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.

        //        // open해야 할 파일명 지정.
        //        //-------------------------------------------------------------------------------------
        //        xlPrinting.OpenFileNameExcel = "HRMF0516_001.xls";
        //        //-------------------------------------------------------------------------------------
        //        // 파일 오픈.
        //        //-------------------------------------------------------------------------------------
        //        bool isOpen = xlPrinting.XLFileOpen();
        //        //-------------------------------------------------------------------------------------

        //        //-------------------------------------------------------------------------------------
        //        if (isOpen == true)
        //        {
        //            // 헤더 부분 인쇄.
        //            //xlPrinting.HeaderWrite(vAccountBook, vToday);

        //            // 라인 인쇄
        //            vPageNumber = xlPrinting.LineWrite(IGR_PAYMENT_ITEM_SUM);

        //            //출력구분에 따른 선택(인쇄 or file 저장)
        //            if (pOutChoice == "PRINT")
        //            {
        //                xlPrinting.Printing(1, vPageNumber);
        //            }
        //            else if (pOutChoice == "FILE")
        //            {
        //                xlPrinting.SAVE(vSaveFileName);
        //            }

        //            //-------------------------------------------------------------------------------------
        //            xlPrinting.Dispose();
        //            //-------------------------------------------------------------------------------------

        //            vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        else
        //        {
        //            vMessageText = "Excel File Open Error";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        //-------------------------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        xlPrinting.Dispose();

        //        vMessageText = ex.Message;
        //        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = false;
        //    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        //    System.Windows.Forms.Application.DoEvents();
        //}

        #endregion;
        
        #region ----- XL Print 1 Methods ----

        private void XLPrinting_Main()
        {
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", iDate.ISMonth_Last(PAY_YYYYMM_0.EditValue));
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0507");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));

            //print type 설정
            DialogResult vdlgResult;
            HRMF0507_PRINT_TYPE vHRMF0507_PRINT_TYPE = new HRMF0507_PRINT_TYPE(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0507_PRINT_TYPE.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            string vPRINT_TYPE = iString.ISNull(vHRMF0507_PRINT_TYPE.Get_Printer_Type);
            if (vPRINT_TYPE == string.Empty)
            {
                return;
            }
            vHRMF0507_PRINT_TYPE.Dispose();

            //급상여대장인쇄.
            if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_DETAIL.TabIndex)
            {
                XLPrinting_DTL(vPRINT_TYPE);
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        } 

        //급상여 상세//
        private void XLPrinting_DTL(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vTitle = string.Empty;
            string vSaveFileName = string.Empty;

            int vPageNumber = 0;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            IDA_PROMPT_PRINT_DTL.Fill();
            IDA_PAYMENT_PRINT_DTL.Fill();
            int vRowCount = IDA_PAYMENT_PRINT_DTL.CurrentRows.Count;

            if (vRowCount < 1)
            {
                isAppInterfaceAdv1.OnAppMessage("Print Data is not found. Check Please");
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            } 

            if (pOutput_Type == "EXCEL")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "Excel file(*.xls)|*.xls|(*.xlsx)|*.xlsx";
                vSaveFileDialog.DefaultExt = "xlsx";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                    return;
            }
            else if (pOutput_Type == "PDF")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "Pdf file(*.pdf)|*.pdf";
                vSaveFileDialog.DefaultExt = "pdf";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                    return;
            }

            vMessageText = string.Format(" Printing Starting...");

            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents(); 

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0507_001.xlsx";
                //-------------------------------------------------------------------------------------

                bool IsOpen = xlPrinting.XLFileOpen();
                if (IsOpen == true)
                {
                    isAppInterfaceAdv1.OnAppMessage("Printing Start...");

                    System.Windows.Forms.Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();

                    string vUserName = isAppInterfaceAdv1.AppInterface.LoginDescription;

                    string vCORP_NAME = CORP_NAME_0.EditValue as string;
                    string vYYYYMM = PAY_YYYYMM_0.EditValue as string;
                    string vWageTypeName = WAGE_TYPE_NAME_0.EditValue as string;
                    string vDepartment_NAME = DEPT_NAME_0.EditValue as string;

                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    //엑셀양식 헤더 인쇄//


                    vPageNumber = xlPrinting.XLWirteMain(IDA_PROMPT_PRINT_DTL, IDA_PAYMENT_PRINT_DTL, vLOCAL_DATE, vUserName, vCORP_NAME, vYYYYMM, vWageTypeName, vDepartment_NAME);
                    
                    if (pOutput_Type == "PDF")
                    { 
                        xlPrinting.PDF_Save(vSaveFileName);
                    }
                    else if(pOutput_Type == "EXCEL")
                    {
                        xlPrinting.Save(vSaveFileName);
                    }
                    else if (pOutput_Type == "PREVIEW")
                    {
                        xlPrinting.PreviewPrinting(1, vPageNumber);
                    }
                    else
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    } 
                    xlPrinting.Dispose();
                }
                else
                {
                    xlPrinting.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                try
                {
                    xlPrinting.Dispose();
                }
                catch
                {

                }
            }


            vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);

            System.Windows.Forms.Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
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
                application.DefaultVersion = ExcelVersion.Excel2007;
                IWorkbook workBook = ExcelUtils.CreateWorkbook(1);
                workBook.Version = ExcelVersion.Excel2007;
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

        #region ----- Events -----

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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    //ExportXL(igrMONTH_PAYMENT);
                    //XLPrinting("FILE");
                    if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_DETAIL.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_VIEW_DTL);
                    }                     
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_VIEW_SUM);
                    } 
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_DEPT.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_VIEW_DEPT);
                    }
                }
            }
        }

        #endregion;
        
        #region ----- Form Event -----

        private void HRMF0507_Load(object sender, EventArgs e)
        {
            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);

            DefaultCorporation();              //Default Corp. 
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaPAY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("PAY_TYPE", "Y");
        }

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("FLOOR", "Y");
        }
        
        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        #endregion


    }
}