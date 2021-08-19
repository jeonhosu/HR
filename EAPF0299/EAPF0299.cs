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
using System.Net;
using System.Xml; 

namespace EAPF0299
{
    public partial class EAPF0299 : Office2007Form
    {
        #region ----- Variables -----

        string mCountPerPage = "10000";      //1페이지당 출력 갯수.
        string mConfirmKey = "U01TX0FVVEgyMDE5MTEwNjE1MzY0OTEwOTE3MjM=";
        string mKeyWord = string.Empty;
        string mApiUrl = string.Empty;

        ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public object Get_Zip_Code
        {
            get
            {
                return IGR_ADDRESS.GetCellValue("ZIP_CODE");
            }
        }

        public object Get_Address
        {
            get
            {
                return IGR_ADDRESS.GetCellValue("ADDRESS");
            }
        }

        #endregion;

        #region ----- Constructor -----

        public EAPF0299()
        {
            InitializeComponent();
        }

        public EAPF0299(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm; //항상 최상위 폼 유지 위해.
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public EAPF0299(Form pMainForm, ISAppInterface pAppInterface, object pZIP_CODE, object pADDRESS)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm; //항상 최상위 폼 유지 위해.
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_ADDRESS.EditValue = pADDRESS;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB(int pCurrPage)
        {
            if (iConvert.ISNull(W_ADDRESS.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10297"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ADDRESS.Focus();
                return;
            }
            mKeyWord = iConvert.ISNull(W_ADDRESS.EditValue);

            if (iConvert.ISNull(W_STRUCTURE_DESC.EditValue) != string.Empty)
            {
                mKeyWord = string.Format("{0} {1}", mKeyWord, W_STRUCTURE_DESC.EditValue);
            }
             

            int vCurrentRow = 0;
            decimal vTotalPage = 1;

            WebClient vWC = new WebClient();

            mApiUrl = "http://www.juso.go.kr/addrlink/addrLinkApi.do?currentPage=" + pCurrPage + "&countPerPage=" + mCountPerPage + "&keyword=" + mKeyWord + "&confmKey=" + mConfirmKey;
            XmlReader vRead = new XmlTextReader(vWC.OpenRead(mApiUrl));
            DataSet vDS = new DataSet();
            vDS.ReadXml(vRead);

            DataRow[] vRow = vDS.Tables[0].Select();
            if (vRow[0]["totalcount"].ToString() != "0")
            {
                decimal vCountPerPage = iConvert.ISDecimaltoZero(vRow[0]["countPerPage"]);
                decimal vTotalCount = iConvert.ISDecimaltoZero(vRow[0]["totalCount"]);
                vTotalPage = Math.Ceiling(vTotalCount / vCountPerPage);

                V_CURRENT.EditValue = pCurrPage;
                V_TOTAL.EditValue = vTotalPage;
             
                foreach (DataRow mRow in vDS.Tables[1].Rows)
                {
                    IGR_ADDRESS.RowCount = vCurrentRow + 1;
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("ZIP_CODE"), mRow["zipNo"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("LAND_ADDR_DESC"), mRow["jibunAddr"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("ROAD_ADDR_DESC"), mRow["roadAddr"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("STRUCTURE_DESC"), mRow["bdNm"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("ADDRESS"), mRow["roadAddr"]);
                    vCurrentRow++;
                }
            }
        }

        private void SEARCH_DB_ADDR(int pPage)
        {
            if (iConvert.ISNull(W_ADDRESS.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10297"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ADDRESS.Focus();
                return;
            }
            mKeyWord = iConvert.ISNull(W_ADDRESS.EditValue);

            if (iConvert.ISNull(W_STRUCTURE_DESC.EditValue) != string.Empty)
            {
                mKeyWord = string.Format("{0} {1}", mKeyWord, W_STRUCTURE_DESC.EditValue);
            }

            int vCurrentRow = 0;
            decimal vTotalPage = 1;

            WebClient vWC = new WebClient();

            mApiUrl = "http://www.juso.go.kr/addrlink/addrLinkApi.do?currentPage=" + pPage + "&countPerPage=" + mCountPerPage + "&keyword=" + mKeyWord + "&confmKey=" + mConfirmKey;
            XmlReader vRead = new XmlTextReader(vWC.OpenRead(mApiUrl));
            DataSet vDS = new DataSet();
            vDS.ReadXml(vRead);

            DataRow[] vRow = vDS.Tables[0].Select();
            if (vRow[0]["totalcount"].ToString() != "0")
            {
                decimal vCountPerPage = iConvert.ISDecimaltoZero(vRow[0]["countPerPage"]);
                decimal vTotalCount = iConvert.ISDecimaltoZero(vRow[0]["totalCount"]);
                vTotalPage = Math.Ceiling(vTotalCount / vCountPerPage);
            }

            for (int r = 1; r <= vTotalPage; r++)
            {
                mApiUrl = "http://www.juso.go.kr/addrlink/addrLinkApi.do?currentPage=" + r + "&countPerPage=" + mCountPerPage + "&keyword=" + mKeyWord + "&confmKey=" + mConfirmKey;
                vRead = new XmlTextReader(vWC.OpenRead(mApiUrl));
                vDS = new DataSet();
                vDS.ReadXml(vRead);
                foreach (DataRow mRow in vDS.Tables[1].Rows)
                {
                    IGR_ADDRESS.RowCount = vCurrentRow + 1;
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("ZIP_CODE"), mRow["zipNo"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("LAND_ADDR_DESC"), mRow["jibunAddr"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("ROAD_ADDR_DESC"), mRow["roadAddr"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("STRUCTURE_DESC"), mRow["bdNm"]);
                    IGR_ADDRESS.SetCellValue(vCurrentRow, IGR_ADDRESS.GetColumnToIndex("ADDRESS"), mRow["roadAddr"]);
                    vCurrentRow++;
                }
            }
        }

        private void ADDRESS_CHOOSE()
        {
            if (IGR_ADDRESS.RowIndex < 0)
            {
                this.DialogResult = DialogResult.Cancel;
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    
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
                    if (IDA_ADDRESS.IsFocused)
                    {
                        IDA_ADDRESS.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ADDRESS.IsFocused)
                    {
                        IDA_ADDRESS.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void EAPF0299_Load(object sender, EventArgs e)
        {

        }

        private void EAPF0299_Shown(object sender, EventArgs e)
        {
            RB_ROAD.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_ADDRESS_TYPE.EditValue = RB_ROAD.RadioCheckedString;

            W_ADDRESS.Focus();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void RB_LAND_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB = sender as ISRadioButtonAdv;

            W_ADDRESS_TYPE.EditValue = RB.RadioCheckedString;
            W_ADDRESS.Focus();
        }

        private void ADDRESS_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB(1);
            }
            else if(e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void BTN_FIND_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB(1);
        }

        private void W_STRUCTURE_NUM_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB(1);
            }
            else if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void W_STRUCTURE_NAME_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB(1);
            }
            else if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void BTN_PRE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int vCurrPage = iConvert.ISNumtoZero(V_CURRENT.EditValue, 0);
            if(vCurrPage <= 1)
            {
                return;
            }
            V_CURRENT.EditValue = vCurrPage - 1; 
            SEARCH_DB(iConvert.ISNumtoZero(V_CURRENT.EditValue));
        }

        private void BTN_NEXT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int vCurrPage = iConvert.ISNumtoZero(V_CURRENT.EditValue, 0);
            if (iConvert.ISNumtoZero(V_TOTAL.EditValue,0) <= vCurrPage)
            {
                return;
            }
            V_CURRENT.EditValue = vCurrPage + 1;
            SEARCH_DB(iConvert.ISNumtoZero(V_CURRENT.EditValue));
        }

        private void BTN_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            ADDRESS_CHOOSE();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void IGR_ADDRESS_CellDoubleClick(object pSender)
        {
            ADDRESS_CHOOSE();
        }

        private void IGR_ADDRESS_CellKeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ADDRESS_CHOOSE();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        #endregion

    }
}