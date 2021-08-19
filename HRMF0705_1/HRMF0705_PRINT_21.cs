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

namespace HRMF0705
{
    public partial class HRMF0705_PRINT_21 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCORP_ID;
        object mPERSON_ID;
        object mPRINT_YEAR;
        object mJOB_CATEGORY_ID;
        object mFLOOR_ID;
        object mCB_EMPLOYE_3_YN;
        object mCB_PRINT_SAVING_YN;
        object mCB_PRINT_HOUSE_YN; 
        object mEARNER_YN;
        object mADDRESSOR1_YN;
        object mADDRESSOR2_YN;
        string mOutChoice;

        #endregion;

        #region ----- Constructor ----- 

        public HRMF0705_PRINT_21(ISAppInterface pAppInterface
                                , object pCorp_ID
                                , object pPERSON_ID
                                , object pPRINT_YEAR
                                , object pJOB_CATEGORY_ID
                                , object pFLOOR_ID
                                , object pCB_EMPLOYE_3_YN
                                , object pCB_PRINT_SAVING_YN
                                , object pCB_PRINT_HOUSE_YN
                                , object pPRINT_DATE
                                , object pEARNER_YN
                                , object pADDRESSOR1_YN
                                , object pADDRESSOR2_YN
                                , string pOutChoice)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mCORP_ID = pCorp_ID;
            mPERSON_ID = pPERSON_ID;
            mPRINT_YEAR = pPRINT_YEAR;
            mJOB_CATEGORY_ID = pJOB_CATEGORY_ID;
            mFLOOR_ID = pFLOOR_ID;
            mCB_EMPLOYE_3_YN = pCB_EMPLOYE_3_YN;
            mCB_PRINT_SAVING_YN = pCB_PRINT_SAVING_YN;
            mCB_PRINT_HOUSE_YN = pCB_PRINT_HOUSE_YN;
            mEARNER_YN = pEARNER_YN;
            mADDRESSOR1_YN = pADDRESSOR1_YN;
            mADDRESSOR2_YN = pADDRESSOR2_YN;
            mOutChoice = pOutChoice;

            PRINT_DATE.EditValue = pPRINT_DATE;            
        }
 
        #endregion;

        #region ----- Private Methods ----


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

        #region ----- XL Print 1 Methods 22 -----
        //HRMF0705_001.xls
        private void XLPrinting1(string pPrint_Year, string pOutChoice)
        {
            System.DateTime vStartTime = DateTime.Now;
            string vMessageText = string.Empty;
            string vPrint_Year = pPrint_Year;

            if (vPrint_Year == "2011" || vPrint_Year == "2012")
            {
                int vCountRow = gridWITHHOLDING_TAX_13.RowCount; //gridWITHHOLDING_TAX 그리드의 총 행수
                if (vCountRow < 1)
                {
                    vMessageText = string.Format("Without Data");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
            }
            else //2013년
            {
                int vCountRow = gridWITHHOLDING_TAX_13.RowCount; //gridWITHHOLDING_TAX 그리드의 총 행수
                if (vCountRow < 1)
                {
                    vMessageText = string.Format("Without Data");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
            }

            int vPageNumber = 3;
            string vPRINT_SAVING_YN = iString.ISNull(CB_PRINT_SAVING_YN.CheckBoxValue);
            //if (vPRINT_SAVING_YN == "Y")
            //{
            //    vPageNumber = 4;
            //}

            string vPRINT_HOUSE_YN = iString.ISNull(CB_PRINT_HOUSE_YN.CheckBoxValue);
            //if (vPRINT_HOUSE_YN == "Y")
            //{
            //    vPageNumber = 5;
            //}



            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            PRINT_DATE.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();


            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                vMessageText = string.Format(" XL Opening...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                //xlPrinting.OpenFileNameExcel = "HRMF0705_001.xls";
                //-------------------------------------------------------------------------------------

                object vPrintDate = PRINT_DATE.DateTimeValue.ToString("yyyy 년  MM 월  dd 일", null);
                //---------------------------------------------------------------------
                // 출력 용도 구분
                //---------------------------------------------------------------------

                string vPrint_Type = null;
                object vPrint_Type_Desc = null;
                if (mEARNER_YN.ToString() == "Y")
                {
                    vPrint_Type = "1";
                    vPrint_Type_Desc = "소득자 보관용";
                    vPageNumber = xlPrinting.WriteMain(gridWITHHOLDING_TAX, gridWITHHOLDING_TAX_13, gridSUPPORT_FAMILY, IGR_SAVING_INFO_2, IGR_SAVING_INFO_3, IGR_SAVING_INFO_4, IGR_SAVING_INFO_5, IGR_SAVING_INFO_6, vPrintDate, vPrint_Type, vPrint_Type_Desc, pOutChoice, vPageNumber, vPRINT_SAVING_YN, vPrint_Year, vPRINT_HOUSE_YN, idaHOUSE_LEASE_INFO_10, idaHOUSE_LEASE_INFO_20);
                }
                if (mADDRESSOR1_YN.ToString() == "Y")
                {
                    vPrint_Type = "2";
                    vPrint_Type_Desc = "발행자 보관용";
                    vPageNumber = xlPrinting.WriteMain(gridWITHHOLDING_TAX, gridWITHHOLDING_TAX_13, gridSUPPORT_FAMILY, IGR_SAVING_INFO_2, IGR_SAVING_INFO_3, IGR_SAVING_INFO_4, IGR_SAVING_INFO_5, IGR_SAVING_INFO_6, vPrintDate, vPrint_Type, vPrint_Type_Desc, pOutChoice, vPageNumber, vPRINT_SAVING_YN, vPrint_Year, vPRINT_HOUSE_YN, idaHOUSE_LEASE_INFO_10, idaHOUSE_LEASE_INFO_20);
                }
                if (mADDRESSOR2_YN.ToString() == "Y")
                {
                    vPrint_Type = "3";
                    vPrint_Type_Desc = "발행자 보고용";
                    vPageNumber = xlPrinting.WriteMain(gridWITHHOLDING_TAX, gridWITHHOLDING_TAX_13, gridSUPPORT_FAMILY, IGR_SAVING_INFO_2, IGR_SAVING_INFO_3, IGR_SAVING_INFO_4, IGR_SAVING_INFO_5, IGR_SAVING_INFO_6, vPrintDate, vPrint_Type, vPrint_Type_Desc, pOutChoice, vPageNumber, vPRINT_SAVING_YN, vPrint_Year, vPRINT_HOUSE_YN, idaHOUSE_LEASE_INFO_10, idaHOUSE_LEASE_INFO_20);
                }
            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            //-------------------------------------------------------------------------------------
            xlPrinting.Dispose();
            //-------------------------------------------------------------------------------------

            System.DateTime vEndTime = DateTime.Now;
            System.TimeSpan vTimeSpan = vEndTime - vStartTime;

            vMessageText = string.Format("Printing End [Total : {0}] ---> {1}", vPageNumber, vTimeSpan.ToString());
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void ibtCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion              

        private void HRMF0705_PRINT_21_Load(object sender, EventArgs e)
        {

            CORP_ID.EditValue = mCORP_ID;
            PERSON_ID.EditValue = mPERSON_ID;
            PRINT_YEAR.EditValue = mPRINT_YEAR;
            JOB_CATEGORY_ID.EditValue = mJOB_CATEGORY_ID;
            FLOOR_ID.EditValue = mFLOOR_ID;
            CB_EMPLOYE_3_YN.CheckBoxValue = mCB_EMPLOYE_3_YN;
            CB_PRINT_SAVING_YN.CheckBoxValue = mCB_PRINT_SAVING_YN;
            CB_PRINT_HOUSE_YN.CheckBoxValue = mCB_PRINT_HOUSE_YN; 

            ida_PRINT_SAVING_INFO_2.SetSelectParamValue("P_SAVING_GROUP", "1"); //퇴직연금 공제                
            ida_PRINT_SAVING_INFO_3.SetSelectParamValue("P_SAVING_GROUP", "2"); //연금저촉 공제                
            ida_PRINT_SAVING_INFO_4.SetSelectParamValue("P_SAVING_GROUP", "3"); //주택마련저축 공제                
            ida_PRINT_SAVING_INFO_5.SetSelectParamValue("P_SAVING_GROUP", "4"); //장기주식형저축 공제 
            ida_PRINT_SAVING_INFO_6.SetSelectParamValue("P_SAVING_GROUP", "5"); //장기집합투자증권 저축 

            idaWITHHOLDING_TAX.Fill();
            idaWITHHOLDING_TAX_13.Fill();
            //부양가족은 IDAWITHHOLDING_TAX_13 아답터와 FILTER관계로 엮어서 따로 FILL 하지 않음;;
            //idaHOUSE_LEASE_INFO_10.Fill();
            //idaHOUSE_LEASE_INFO_20.Fill();

            string vPrint_Year = iString.ISNull(PRINT_YEAR.EditValue);
            XLPrinting1(vPrint_Year, mOutChoice); // 출력 함수 호출

            this.Close();
        }

    }
}