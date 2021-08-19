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
    public partial class HRMF0705 : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0705()
        {
            InitializeComponent();
        }

        public HRMF0705(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Default Method ----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        #endregion;

        #region ----- Printing Method ----

        private void isOnPrinting(DateTime pPrint_Date, string pPrint_num, string pOutChoice)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return; 
            }

            if (STD_DATE_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_DATE_0.Focus();
                return;                            
            }
            
            DialogResult vdlgResult;

            if (pPrint_num != null)
            {
                Form vHRMF0705_PRINT = new HRMF0705_PRINT(isAppInterfaceAdv1.AppInterface
                                                          , Convert.ToInt32(CORP_ID_0.EditValue)                            //업체ID
                                                          , pPrint_num                                                      //발급번호
                                                          , pPrint_Date                                                     //발급일자
                                                          , igrCERTIFICATE.GetCellValue("PRINT_YEAR").ToString()            //징수년도                                                      
                                                          , igrCERTIFICATE.GetCellValue("CERT_TYPE_NAME").ToString()        //증명서구분
                                                          , igrCERTIFICATE.GetCellValue("CERT_TYPE_ID")                     //증명서 구분ID
                                                          , igrCERTIFICATE.GetCellValue("CERT_TYPE_CODE")                   //증명서 코드
                                                          , igrCERTIFICATE.GetCellValue("PERSON_NAME").ToString()           //성명
                                                          , igrCERTIFICATE.GetCellValue("PERSON_ID")                        //사원ID
                                                          , Convert.ToDateTime(igrCERTIFICATE.GetCellValue("JOIN_DATE"))    //입사일자
                                                          , Convert.ToDateTime(igrCERTIFICATE.GetCellValue("RETIRE_DATE"))  //퇴사일자
                                                          , igrCERTIFICATE.GetCellValue("DESCRIPTION").ToString()           //용도
                                                          , igrCERTIFICATE.GetCellValue("SEND_ORG").ToString()              //제출처
                                                          , Convert.ToInt32(igrCERTIFICATE.GetCellValue("PRINT_COUNT"))     //매수
                                                          , pOutChoice                                                      //출력방향[프린터/파일]
                                                          );
                vdlgResult = vHRMF0705_PRINT.ShowDialog();
                if (vdlgResult == DialogResult.OK)
                {}
                vHRMF0705_PRINT.Dispose();
            }
            else {
                DateTime Join_Date = new DateTime(1,1,1);
                DateTime Retire_Date = new DateTime(1,1,1);

                Form vHRMF0705_PRINT = new HRMF0705_PRINT(isAppInterfaceAdv1.AppInterface
                                                          , Convert.ToInt32(CORP_ID_0.EditValue) //업체ID
                                                          , pPrint_num                  //발급번호
                                                          , pPrint_Date                 //발급일자
                                                          , pPrint_Date.Year.ToString() //징수년도                                                      
                                                          , null                        //증명서구분
                                                          , null                        //증명서 구분ID
                                                          , null                        //증명서 코드
                                                          , null                        //성명
                                                          , null                        //사원ID
                                                          , Join_Date                   //입사일자
                                                          , Retire_Date                 //퇴사일자
                                                          , null                        //용도
                                                          , null                        //제출처
                                                          , 1                           //매수
                                                          , pOutChoice                  //출력방향[프린터/파일]
                                                          );                

                vdlgResult = vHRMF0705_PRINT.ShowDialog();
                if (vdlgResult == DialogResult.OK)
                {}
                vHRMF0705_PRINT.Dispose();
            }            
        }

        private void SEARCH_DB()
        {
            idaCERTIFICATE.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            idaCERTIFICATE.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            idaCERTIFICATE.Fill();
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    DateTime dPrint_Date = DateTime.Today;
                    string sPrint_num = null;

                    isOnPrinting(dPrint_Date, sPrint_num, "PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    DateTime dPrint_Date = DateTime.Today;
                    string sPrint_num = null;

                    isOnPrinting(dPrint_Date, sPrint_num, "FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Load Event -----

        private void HRMF0705_Load(object sender, EventArgs e)
        {
            STD_DATE_0.EditValue = DateTime.Today;

            DefaultCorporation();
        }

        #endregion;

        #region ----- Lookup Event -----

        private void ilaCERT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "CERT_TYPE");
            ildCOMMON.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 20 ");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }        

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", STD_DATE_0.EditValue);
        }

        //정산관련 인쇄관리 Grid에서 데이터 선택 시, 인쇄 폼으로 해당 데이터를 넘겨준다.
        /*
        private void igrCERTIFICATE_DoubleClick(object sender, EventArgs e) 
        {
            DateTime dPrint_Date = DateTime.Today;
            string sPrint_num = null;

            if (igrCERTIFICATE.RowCount > 0)
            {
                dPrint_Date = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("PRINT_DATE"));
                sPrint_num = igrCERTIFICATE.GetCellValue("PRINT_NUM").ToString();
            }
            isOnPrinting(dPrint_Date, sPrint_num);
        }
        */

        #endregion

        #region ----- Double Click Event -----

        //정산관련 인쇄관리 Grid에서 데이터 더블 클릭 시, 인쇄 폼으로 해당 데이터 인쇄 폼으로 넘겨준다.
        private void igrCERTIFICATE_CellDoubleClick(object pSender)
        {
            DateTime dPrint_Date = DateTime.Today;
            string sPrint_num = null;

            if (igrCERTIFICATE.RowCount > 0)
            {
                dPrint_Date = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("PRINT_DATE"));
                sPrint_num = igrCERTIFICATE.GetCellValue("PRINT_NUM").ToString();
            }
            isOnPrinting(dPrint_Date, sPrint_num, "PRINT");
        }

        #endregion;
    }
}