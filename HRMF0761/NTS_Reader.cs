using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0761
{
    public partial class NTS_Reader : Form
    {
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISConvert iDate = new ISFunction.ISConvert();

        public NTS_Reader()
        {
            InitializeComponent();
        }

        public NTS_Reader(ISAppInterface pAppinterface, object pPERIOD_NAME, object pNAME, object pPERSON_NUM)
        {
            InitializeComponent();

            isAppInterfaceAdv1.AppInterface = pAppinterface;

            V_YYYYMM.EditValue = pPERIOD_NAME;
            V_NAME.EditValue = pNAME;
            V_PERSON_NUM.EditValue = pPERIOD_NAME;
        }

        #region ---- Class -----

        public object PDF_Filename
        {
            get
            {
                return V_FILE_NAME.EditValue;
            }
        }

        public string PDF_StrBuf
        {
            get
            {
                return txtUtf8.Text;
            }
        }

        public object PDF_PWD
        {
            get
            {
                return V_FILE_PWD.EditValue;
            }
        }
        
        #endregion

        #region ---- Property -----

        private bool PDF_Verify()
        {
            bool vVerify = false;

            long result = 0;
            string vFilePath = iConv.ISNull(V_FILE_NAME.EditValue);
            string strMsg = string.Empty;

            byte[] baGenTime = new byte[1024];
            byte[] baHashAlg = new byte[1024];
            byte[] baHashVal = new byte[1024];
            byte[] baCertDN = new byte[1024];

            result = TstUtil.DSTSPdfSigVerifyF(vFilePath, baGenTime, baHashAlg, baHashVal, baCertDN);


            String sGenTimeTemp = Encoding.Unicode.GetString(baGenTime);
            String sHashAlgTemp = Encoding.Unicode.GetString(baHashAlg);
            String sHashValTemp = Encoding.Unicode.GetString(baHashVal);
            String sCertDNTemp = Encoding.Unicode.GetString(baCertDN);

            String sGenTime = sGenTimeTemp.Replace('\0', ' ').Trim();
            String sHashAlg = sHashAlgTemp.Replace('\0', ' ').Trim();
            String sHashVal = sHashValTemp.Replace('\0', ' ').Trim();
            String sCertDN = sCertDNTemp.Replace('\0', ' ').Trim();

            switch (result)
            {
                case 0:
                    vVerify = true;
                    strMsg = String.Format("원본 파일입니다. \n\nTS시각: {0} \n해쉬알고리즘: {1} \n해쉬값: {2} \nTSA인증서: {3}", sGenTime, sHashAlg, sHashVal, sCertDN);
                    break;
                case 2101:
                    vVerify = true;
                    strMsg = String.Format("문서가 변조되었습니다.");
                    break;
                case 2102:
                    vVerify = true;
                    strMsg = String.Format("TSA 서명 검증에 실패하였습니다.");
                    break;
                case 2103:
                    vVerify = true;
                    strMsg = String.Format("지원하지 않는 해쉬알고리즘 입니다.");
                    break;
                case 2104:
                    vVerify = true;
                    strMsg = String.Format("해당 파일을 읽을 수 없습니다.");
                    break;
                case 2105:
                    vVerify = true;
                    strMsg = String.Format("서명검증을 위한 API 처리 오류입니다.");
                    break;
                case 2106:
                    vVerify = true;
                    strMsg = String.Format("타임스탬프 토큰 데이터 파싱 오류입니다.");
                    break;
                case 2107:
                    vVerify = true;
                    strMsg = String.Format("TSA 인증서 처리 오류입니다.");
                    break;
                case 2108:
                    vVerify = false;
                    strMsg = String.Format("타임스탬프가 적용되지 않은 파일입니다.");
                    break;
                case 2109:
                    vVerify = false;
                    strMsg = String.Format("인증서 검증에 실패 하였습니다.");
                    break;
                default:
                    vVerify = false;
                    strMsg = String.Format("에러코드 미정의 error - [%d]", result);
                    break;
            }
            if(vVerify == false)
            {                   
                MessageBox.Show(strMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return vVerify;
        }

        private bool PDF_Import()
        {
            bool vImport_Flag = false;
            string vFilePath = iConv.ISNull(V_FILE_NAME.EditValue);
            string vPassword = iConv.ISNull(V_FILE_PWD.EditValue);
            string strXML = "XML";

            if (string.IsNullOrEmpty(vFilePath))
            {
                return vImport_Flag;
            }

            if (PDF_Verify() == false)
            {
                return vImport_Flag;
            }


            int fileSize = EXPFile.NTS_GetFileSize(vFilePath, vPassword, strXML, 0);

            if (fileSize > 0)
            {
                byte[] buf = new byte[fileSize];
                fileSize = EXPFile.NTS_GetFileBuf(vFilePath, vPassword, strXML, buf, 0);
                if (fileSize > 0)
                {
                    string strBuf = Encoding.UTF8.GetString(buf);
                    strBuf.Replace("\n", "\r\n");
                    txtUtf8.Text = strBuf;
                }
            }

            if (fileSize == -10)
            {
                MessageBox.Show("파일이 없거나 손상된 PDF 파일입니다.");
                vImport_Flag = false;
                return vImport_Flag;
            }
            else if (fileSize == -11)
            {
                MessageBox.Show("국세청에서 발급된 전자문서가 아닙니다.");
                vImport_Flag = false;
                return vImport_Flag;
            }
            else if (fileSize == -13)
            {
                MessageBox.Show("추출용 버퍼가 유효하지 않습니다.");
                vImport_Flag = false;
                return vImport_Flag;
            }
            else if (fileSize == -200)
            {
                MessageBox.Show("비밀번호가 틀립니다.");
                vImport_Flag = false;
                return vImport_Flag;
            }
            vImport_Flag = true;
            return vImport_Flag;
        }

        #endregion
         
        #region ----- Form Event -----

        private void NTS_Reader_Load(object sender, EventArgs e)
        {

        }

        private void BTN_FILE_FIND_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            OpenFileDialog fileDlg = new OpenFileDialog();
            fileDlg.RestoreDirectory = true;

            fileDlg.Filter = "PDF Files(*.PDF)|*.PDF";
            fileDlg.Multiselect = false;

            if (fileDlg.ShowDialog() == DialogResult.OK)
            {
                V_FILE_NAME.EditValue = fileDlg.FileName;
                V_FILE_PWD.EditValue = string.Empty; 
                txtUtf8.Text = string.Empty;
            }
        }

        private void BTN_PDF_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (PDF_Import() == true)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion
         
    }
}
