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

using System.Net; 
using System.IO;
using System.Xml;

using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0301
{
    public partial class HRMF0301 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0301(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;            
        }

        #endregion;

        #region ----- Private Methods ----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultSetAllCheck()
        {
            igrHOLIDAY.SetCellValue("ALL_CHECK", "Y");
        }

        private void isSEARCH_DB()
        {// 데이터 조회
            if (string.IsNullOrEmpty(W_WORK_YYYY.EditValue.ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                W_WORK_YYYY.Focus();
                return;
            }
            IDA_HOLIDAY_CALENDAR.Fill();
            igrHOLIDAY.Focus();
        }

        private bool isData_Add()
        {// 데이터 추가전 검증.
            if (string.IsNullOrEmpty(W_WORK_YYYY.EditValue.ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        #endregion;
        
        #region ----- Convert String Method ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    isSEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {                    
                    IDA_HOLIDAY_CALENDAR.AddOver(); 
                    DefaultSetAllCheck();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IDA_HOLIDAY_CALENDAR.AddUnder();
                    DefaultSetAllCheck();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_HOLIDAY_CALENDAR.Update();    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_HOLIDAY_CALENDAR.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {                    
                    IDA_HOLIDAY_CALENDAR.Delete();
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0301_Load(object sender, EventArgs e)
        {
            IDA_HOLIDAY_CALENDAR.FillSchema();

            W_WORK_YYYY.EditValue = System.DateTime.Today.Year.ToString();
        }

        private void BTN_INIT_DATE_ButtonClick(object pSender, EventArgs pEventArgs)        
        {
            if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            IDC_GET_API_AUTH_KEY_P.ExecuteNonQuery();
            string vAuth_Key = iConv.ISNull(IDC_GET_API_AUTH_KEY_P.GetCommandParamValue("O_KEY_VALUE"));

            string url = "http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo"; // URL
            url += "?ServiceKey=" + vAuth_Key;  // "HWRwrRtZlqXfu6yLCnW6A6oY5l5Bg8OblhOQIrWIURHAF7YNN6ZPN%2BCFJgilf%2FIch243dXcXgUUvd2FVjtyf7A%3D%3D"; // Service Key
            url += "&pageNo=1";
            url += "&numOfRows=900";
            url += "&solYear=" + iConv.ISNull(W_WORK_YYYY.EditValue);
            //url += "&solMonth=";

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";

            string results = string.Empty;
            HttpWebResponse response;
            using (response = request.GetResponse() as HttpWebResponse)
            {
                StreamReader reader = new StreamReader(response.GetResponseStream());
                results = reader.ReadToEnd();
            }
            if (string.IsNullOrEmpty(results))
            {
                MessageBoxAdv.Show(string.Format("Calendar Copy :: {0}", isMessageAdapter1.ReturnText("EAPP_10061")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                isSEARCH_DB();
                return;
            }

            //xml 파싱//
            XmlDocument vXml = new XmlDocument();
            vXml.LoadXml(results);
            XmlNodeList xnList = vXml.SelectNodes("/response/body/items/item"); //접근할 노드 
            foreach (XmlNode xn in xnList)
            {
                string vHoliday_Name = xn["dateName"].InnerText; //예술반 불러오기
                string vWork_Date = xn["locdate"].InnerText; //예술반 code_name 불러오기 
                  
                IDC_SAVE_HOLIDAY.SetCommandParamValue("P_WORK_DATE", vWork_Date);
                IDC_SAVE_HOLIDAY.SetCommandParamValue("P_HOLIDAY_NAME", vHoliday_Name);
                IDC_SAVE_HOLIDAY.ExecuteNonQuery();

                string vStatus = iConv.ISNull(IDC_SAVE_HOLIDAY.GetCommandParamValue("O_STATUS"));
                string vMessage = iConv.ISNull(IDC_SAVE_HOLIDAY.GetCommandParamValue("O_MESSAGE"));
                if (IDC_SAVE_HOLIDAY.ExcuteError)
                {
                    MessageBoxAdv.Show(IDC_SAVE_HOLIDAY.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if(vStatus.Equals("F"))
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }  
            }

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10038", "&&VALUE:=Calendar Copy"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            isSEARCH_DB();
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_HOLIDAY_CALENDAR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {// 저장전 검증           
            if (iConv.ISNull(e.Row["WORK_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Date(일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //조회 년도와 입력 년도 검증//
            if (iConv.ISNumtoZero(W_WORK_YYYY.EditValue,0) != iDate.ISGetDate(e.Row["WORK_DATE"]).Year)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10581"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(e.Row["HOLIDAY_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Hliday Name(휴일명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_HOLIDAY_CALENDAR_PreDelete(ISPreDeleteEventArgs e)
        {// 삭제 검증.
             
        }

        #endregion

        #region ------ Lookup Event -----

        private void ilaYEAR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYEAR.SetLookupParamValue("W_START_YEAR", iDate.ISYear(iDate.ISGetDate(), -10));
            ildYEAR.SetLookupParamValue("W_END_YEAR", iDate.ISYear(iDate.ISGetDate(), 2));
        }

        private void ilaYEAR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYEAR.SetLookupParamValue("W_START_YEAR", W_WORK_YYYY.EditValue.ToString());
            ildYEAR.SetLookupParamValue("W_END_YEAR", W_WORK_YYYY.EditValue.ToString());
        }

        private void ilaYEAR_1_SelectedRowData(object pSender)
        {
            System.Windows.Forms.SendKeys.Send("{TAB}");
        }

        private void ILA_HOLIDAY_CAL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_HOLIDAY_CAL_TYPE.SetLookupParamValue("W_GROUP_CODE", "HOLIDAY_CAL_TYPE");
            ILD_HOLIDAY_CAL_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ------ Button Event -----

        private void DATA_COPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vStatus = "F";
            string vMessage = string.Empty;

            DialogResult vdlgResult;

            //[FCM_10422]복사 하시겠습니까?
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10422"), "Holiday Copy", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();
            try
            {
                
                
                IDC_HOLIDAY_COPY.ExecuteNonQuery();
                vStatus =  iConv.ISNull(IDC_HOLIDAY_COPY.GetCommandParamValue("O_STATUS"));
                vMessage = iConv.ISNull(IDC_HOLIDAY_COPY.GetCommandParamValue("O_MESSAGE"));
                
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (IDC_HOLIDAY_COPY.ExcuteError || vStatus == "F")
                {
                    if (vMessage != string.Empty)
                    {
                        MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                        
                    }
                    return;
                }

                //[SDM_10027]복사 완료 되었습니다.
                //[FCM_10423]재검색후, 석가탄신일, 설, 추석 연휴 일자를 수정 하세요!
                vMessage = string.Format("{0}\n\n{1}", isMessageAdapter1.ReturnText("SDM_10027"), isMessageAdapter1.ReturnText("FCM_10423"));
                MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);            
            }
            catch (System.Exception ex)
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion

    }
}