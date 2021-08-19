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


namespace FX_Main
{
    public partial class APPF0020 : Office2007Form
    {
        #region ----- Enums -----



        #endregion;

        #region ----- Variables -----

        private ISAppInterface mAppInterface;
        private APPF0030 mAPPF0030 = null;

        private string mUserType = string.Empty;
        private string mUserAuthorityType = string.Empty;

        private System.IO.DirectoryInfo mRootInstallationDirectory = null;
        private string mBaseWorkingDirectory = "Flex_ERP\\Kor";
        private string mInstallationDirectory;
        private string mReportDirectory = "Report";

        private System.DateTime mLoginStartingDateTime = System.DateTime.Now;
        
        #endregion;

        #region ----- Constructor -----

        public APPF0020(string pOraHost, string pOraPort, string pOraServiceName, string pOraUserId, string pOraPassword,
                        string pAppHost, string pAppPort, string pAppUserId, string pAppPassword,
                        string pSOBID, string pORGID,
                        string pLoginId, string pLoginDescription, string pUserDisplayName, string pLoginDate, string pLoginTime,
                        string pTerritoryLanguage, string pUserType, string pUserAuthorityType, string pLoginNo,
                        string pPersonID, string pPersonNumber, string pDepartmentID, string pDepartmentName,
                        string pBaseWorkingDirectory)
        {
            InitializeComponent();

            string vDateTime = string.Format("{0} {1}", pLoginDate, pLoginTime);
            DateTime vLoginDateTime = Convert.ToDateTime(vDateTime);



            //----------------------------------------------------------------------------------------------
            //[Login Time]
            //----------------------------------------------------------------------------------------------
            vDateTime = string.Format("{0} {1}", pLoginDate, pLoginTime);
            mLoginStartingDateTime = Convert.ToDateTime(vDateTime);
            string vMessage = string.Format("Login Time : {0} {1}", pLoginDate, pLoginTime);
            MessageListBox(vMessage);
            //----------------------------------------------------------------------------------------------


            int vSOBID = Convert.ToInt32(pSOBID);
            int vORGID = Convert.ToInt32(pORGID);

            ISUtil.Enum.TerritoryLanguage vTerritoryLanguage = ISUtil.Enum.TerritoryLanguage.Default;
            if (pTerritoryLanguage == "ENG")
            {
                vTerritoryLanguage = ISUtil.Enum.TerritoryLanguage.Default;
            }
            else if (pTerritoryLanguage == "KOR")
            {
                vTerritoryLanguage = ISUtil.Enum.TerritoryLanguage.TL1_KR;
            }

            mUserType = pUserType;
            mUserAuthorityType = pUserAuthorityType;

            ISDataUtil.AppHostInfo vAppHostInfo = new ISDataUtil.AppHostInfo(pAppHost, pAppPort, pAppUserId, pAppPassword);
            //ISDataUtil.OraConnectionInfo vOraConnectionInfo = new ISDataUtil.OraConnectionInfo(pOraHost, pOraPort, pOraServiceName, pOraUserId, pOraPassword, ISUtil.Enum.TerritoryLanguage.Design);
            ISDataUtil.OraConnectionInfo vOraConnectionInfo = new ISDataUtil.OraConnectionInfo(pOraHost, pOraPort, pOraServiceName, pOraUserId, pOraPassword, vTerritoryLanguage);
            mAppInterface = new ISAppInterface(vAppHostInfo, vOraConnectionInfo, Convert.ToInt32(pLoginId), pLoginDescription, vLoginDateTime, vSOBID, vORGID);
            mAppInterface.OnAppMessage += Application_OnAppMessage;

            //[2010-11-02[화]
            mAppInterface.DisplayName = pUserDisplayName;
            mAppInterface.LoginNo = pLoginNo;
            mAppInterface.PersonId = int.Parse(pPersonID);
            mAppInterface.DeptId = int.Parse(pDepartmentID);
            mAppInterface.DeptName = pDepartmentName;

            isAppInterfaceAdv1.AppInterface = mAppInterface;

            mBaseWorkingDirectory = pBaseWorkingDirectory;
            mRootInstallationDirectory = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles));
            mInstallationDirectory = string.Format("{0}\\{1}", mRootInstallationDirectory.FullName, mBaseWorkingDirectory);

            //dir ..\Ch\Report
            //dir ..\Kor\Report
            int vCutStart = mBaseWorkingDirectory.LastIndexOf("\\") + 1;
            int vCutLength = mBaseWorkingDirectory.Length - vCutStart;
            string vDirectoryTerritory = mBaseWorkingDirectory.Substring(vCutStart, vCutLength);
            mReportDirectory = string.Format("..\\{0}\\{1}", vDirectoryTerritory, mReportDirectory);
            this.Tag = mReportDirectory;
        }

        #endregion;

        #region ----- Property -----

        //public string PathReport
        //{
        //    get
        //    {
        //        return mReportDirectory;
        //    }
        //}

        #endregion;

        #region ----- MessageListBox -----

        private void MessageListBox(string pMessage)
        {
            ListBox_Message.Items.Add(pMessage);
            ListBox_Message.SelectedIndex = ListBox_Message.Items.Count - 1;
        }

        #endregion;

        #region ----- MDi Background -----

        private void MDi_BackGround(string pFileName)
        {
            string vCurrentDirectory = System.IO.Directory.GetCurrentDirectory();
            string vPathImage = string.Format("{0}\\{1}", vCurrentDirectory, pFileName);

            System.IO.FileStream vFileStream;

            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;

            if (System.IO.File.Exists(vPathImage) == true)
            {
                vFileStream = System.IO.File.OpenRead(vPathImage);
                vFileStream.Position = 0;
                foreach (System.Windows.Forms.Control control in this.Controls)
                {
                    System.Windows.Forms.MdiClient client = control as System.Windows.Forms.MdiClient;
                    if (client != null)
                    {
                        //client.BackColor = System.Drawing.Color.White;
                        client.BackgroundImage = System.Drawing.Image.FromStream(vFileStream);
                        break;
                    }
                }
            }
        }

        #endregion;

        #region ----- Private Methods ----

        private void Navigator()
        {
            string vMessage = string.Empty;

            System.DateTime vStartTime = DateTime.Now;
            vMessage = string.Format("Start Time : {0}", vStartTime.ToString("yyyy-MM-dd HH:mm:ss.fffffff"));
            MessageListBox(vMessage);


            mAPPF0030 = new APPF0030(this, mAppInterface);
            mAPPF0030.Show();


            System.DateTime vEndTime = DateTime.Now;
            vMessage = string.Format("End   Time : {0}", vEndTime.ToString("yyyy-MM-dd HH:mm:ss.fffffff"));
            MessageListBox(vMessage);

            System.TimeSpan vTimeSpan = vEndTime - vStartTime;
            vMessage = string.Format("Span  Time : {0}", vTimeSpan.ToString());
            MessageListBox(vMessage);

            System.DateTime vLoginEndDateTime = DateTime.Now;
            System.TimeSpan vLoginSpan = vLoginEndDateTime - mLoginStartingDateTime;
            vMessage = string.Format("Login Span : {0}", vLoginSpan.ToString());
            MessageListBox(vMessage);
        }

        private void Application_OnAppMessage(string pMessageText)
        {
            appStatusText.Text = pMessageText; //Status Bar
        }

        private void Application_OnAppProgress(int pValue)
        {
            appProgressBar.Minimum = 0;
            appProgressBar.Maximum = 100;
            appProgressBar.Step = 2;
            appProgressBar.Style = ProgressBarStyle.Continuous;
            appProgressBar.Value = pValue;
        }

        private void f_Screen_Center_Location()
        {
            //작업표시 1줄 Pixel Size = 34
            //작업표시 2줄 Pixel Size = 64

            int v_Scrren_Count = System.Windows.Forms.Screen.AllScreens.Length;

            int v_Screen_Width = 0;
            int v_Screen_Height = 0;
            int v_Task_Bar = 0;
            int v_this_Location_X = 0;
            int v_this_Location_Y = 0;

            if (v_Scrren_Count > 1)
            {
                v_Screen_Width = System.Windows.Forms.Screen.AllScreens[1].Bounds.Width;
                v_Screen_Height = System.Windows.Forms.Screen.AllScreens[1].Bounds.Height;

                this.Width = v_Screen_Width - 300;
                this.Height = v_Screen_Height - 130;
            }
            else
            {
                v_Task_Bar = 64;

                v_Screen_Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
                v_Screen_Height = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;
            }

            int v_PrimaryScreen_Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;

            int v_Task_Bar_Cut_Screen_Height = v_Screen_Height - v_Task_Bar;

            int v_this_Form_Width = this.Width;
            int v_this_Form_Height = this.Height;

            if (v_Scrren_Count > 1)
            {
                v_this_Location_X = v_Screen_Width - v_this_Form_Width + v_PrimaryScreen_Width;
            }
            else
            {
                v_this_Location_X = v_Screen_Width - v_this_Form_Width;
            }
            v_this_Location_Y = v_Task_Bar_Cut_Screen_Height - v_this_Form_Height;

            if (v_this_Location_X < 0)
            {
                v_this_Location_X = 0;
            }
            if (v_this_Location_Y < 0)
            {
                v_this_Location_Y = 0;
            }

            this.Location = new System.Drawing.Point(v_this_Location_X, v_this_Location_Y);
        }

        #endregion;

        #region ----- Open Form Close -----

        private void OpenFormClose()
        {
            if (this.ActiveMdiChild != null)
            {
                string vActiveChildFormName = this.ActiveMdiChild.Name;
                if (vActiveChildFormName != "APPF0030")
                {
                    Form OpenForm = this.ActiveMdiChild;
                    OpenForm.Close();
                }
            }
        }

        #endregion;

        #region ----- FormLoading Methods -----

        private bool FormLoading1(string psPathTemp, string psLoadFile)
        {
            try
            {
                string vDllFilePathString = string.Format("{0}\\{1}\\bin\\Debug\\{2}.dll", psPathTemp, psLoadFile, psLoadFile);
                string vNameSpaceClass = string.Format("{0}.{1}", psLoadFile, psLoadFile);

                System.Reflection.Assembly vAssembly = System.Reflection.Assembly.LoadFrom(vDllFilePathString);
                System.Type vType = vAssembly.GetType(vNameSpaceClass);

                if (vType != null)
                {
                    object[] voParam = new object[2];
                    voParam[0] = this;
                    voParam[1] = mAppInterface;

                    object vObject = Activator.CreateInstance(vType, voParam);
                    System.Windows.Forms.Form vForm = vObject as System.Windows.Forms.Form;
                    vForm.Show();

                    return true;
                }
                else
                {
                    MessageBoxAdv.Show("Longing Assembly is Null");
                    return false;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        #endregion;

        #region ----- Events -----

        private void barItem3_Click(object sender, EventArgs e)
        {
            //어플리케이션 종료
            System.Windows.Forms.Application.Exit();
        }

        private void APPF0020_FormClosing(object sender, FormClosingEventArgs e)
        {
            //DialogResult ChoiceValue;

            //ChoiceValue = MessageBox.Show("프로그램_종료...?", "FLEX ERP", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

            //if (ChoiceValue != DialogResult.Yes)
            //{
            //    e.Cancel = true;
            //}
        }

        private void APPF0020_FormClosed(object sender, FormClosedEventArgs e)
        {
            foreach (System.Windows.Forms.Form OpenForm in this.MdiChildren)
            {
                if (OpenForm != null)
                {
                    OpenForm.Close();
                }
            }
        }

        private void MdiBackColor()
        {
            foreach (Control control in this.Controls)
            {
                MdiClient client = control as MdiClient;
                if (client != null)
                {
                    client.BackColor = System.Drawing.Color.White;
                }
            }
        }

        private void APPF0020_Load(object sender, EventArgs e)
        {
            //int vCount = this.Controls.Count - 1;
            //this.Controls[this.Controls.Count - vCount].BackColor = System.Drawing.Color.White;

            //MdiBackColor();

            //f_Screen_Center_Location();

            this.Size = new System.Drawing.Size(1400, 787);

            Navigator();
        }

        private void APPF0020_Shown(object sender, EventArgs e)
        {
            //MDi_BackGround("FPCB_0001.png");

            string vDBiP = isAppInterfaceAdv1.AppInterface.OraConnectionInfo.Host;
            this.Text = string.Format("{0} - {1} - {2}", this.Text, mAppInterface.SOB_ID, vDBiP);



            this.Tag = appProgressBar;
        }

        private void barItem22_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.Search);
        }

        private void barItem1_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.AddOver);
        }

        private void barItem16_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.AddUnder);
        }

        private void barItem6_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.Update);
        }

        private void barItem10_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.Cancel);
        }

        private void barItem14_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.Delete);
        }

        private void barItem4_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.Print);
        }

        private void barItem5_Click(object sender, EventArgs e)
        {
            mAppInterface.MainButtonEvent(ISUtil.Enum.AppMainButtonType.Export);
        }

        private void barItem7_Click(object sender, EventArgs e)
        {
            //계단식
            this.LayoutMdi(MdiLayout.Cascade);
        }

        private void barItem8_Click(object sender, EventArgs e)
        {
            //수평
            this.LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void barItem9_Click(object sender, EventArgs e)
        {
            //수직
            this.LayoutMdi(MdiLayout.TileVertical);
        }

        private void barItem11_Click(object sender, EventArgs e)
        {
            mAPPF0030.Activate();
        }

        private void barItem18_Click(object sender, EventArgs e)
        {
            string vPath = string.Format("{0}", @"..\..\..");
            FormLoading1(vPath, "AppAssemblyEntry");
        }

        private void barItem21_Click(object sender, EventArgs e)
        {
            string vPath = string.Format("{0}", @"..\..\..");
            FormLoading1(vPath, "AppMenuEntry");
        }

        private void barItem2_Click(object sender, EventArgs e)
        {
            //Form_Close
            OpenFormClose();
        }

        #endregion;
    }
}