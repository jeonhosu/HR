namespace HRMF0607
{
    partial class HRMF0607_FILE
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo oraConnectionInfo1 = new InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.CHK_ENCRYPT_PWD = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.ENCRYPT_PWD = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.btnCLOSE = new InfoSummit.Win.ControlAdv.ISButton();
            this.igbCONFIRM_INFOMATION = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.btnCANCEL = new InfoSummit.Win.ControlAdv.ISButton();
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.igbCONFIRM_INFOMATION)).BeginInit();
            this.igbCONFIRM_INFOMATION.SuspendLayout();
            // 
            // CHK_ENCRYPT_PWD
            // 
            this.CHK_ENCRYPT_PWD.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CHK_ENCRYPT_PWD.ComboBoxValue = "";
            this.CHK_ENCRYPT_PWD.ComboData = null;
            this.CHK_ENCRYPT_PWD.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.CHK_ENCRYPT_PWD.DataAdapter = null;
            this.CHK_ENCRYPT_PWD.DataColumn = "";
            this.CHK_ENCRYPT_PWD.DateTimeValue = new System.DateTime(2010, 3, 17, 19, 7, 59, 703);
            this.CHK_ENCRYPT_PWD.DoubleValue = 0;
            this.CHK_ENCRYPT_PWD.EditValue = "";
            // 
            // HRMF0607_FILE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(306, 137);
            this.ControlBox = false;
            this.Controls.Add(this.igbCONFIRM_INFOMATION);
            this.Controls.Add(this.btnCANCEL);
            this.Controls.Add(this.btnCLOSE);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HRMF0607_FILE";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "전산매체 암호화 암호 입력";
            this.Load += new System.EventHandler(this.HRMF0607_FILE_Load);
            this.Shown += new System.EventHandler(this.HRMF0607_FILE_Shown);
            this.CHK_ENCRYPT_PWD.Location = new System.Drawing.Point(8, 42);
            this.CHK_ENCRYPT_PWD.LookupAdapter = null;
            this.CHK_ENCRYPT_PWD.Name = "CHK_ENCRYPT_PWD";
            this.CHK_ENCRYPT_PWD.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.CHK_ENCRYPT_PWD.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.CHK_ENCRYPT_PWD.PromptText = "(확인)암호 비밀번호";
            isLanguageElement5.Default = "Check Password";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "(확인)암호 비밀번호";
            isLanguageElement5.TL2_CN = "";
            isLanguageElement5.TL3_VN = "";
            isLanguageElement5.TL4_JP = "";
            isLanguageElement5.TL5_XAA = "";
            this.CHK_ENCRYPT_PWD.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.CHK_ENCRYPT_PWD.PromptWidth = 130;
            this.CHK_ENCRYPT_PWD.Size = new System.Drawing.Size(265, 21);
            this.CHK_ENCRYPT_PWD.TabIndex = 1;
            this.CHK_ENCRYPT_PWD.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.CHK_ENCRYPT_PWD.TextValue = "";
            this.CHK_ENCRYPT_PWD.UseSystemPassword = true;
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
            // 
            // ENCRYPT_PWD
            // 
            this.ENCRYPT_PWD.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.ENCRYPT_PWD.ComboBoxValue = "";
            this.ENCRYPT_PWD.ComboData = null;
            this.ENCRYPT_PWD.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.ENCRYPT_PWD.DataAdapter = null;
            this.ENCRYPT_PWD.DataColumn = "";
            this.ENCRYPT_PWD.DateTimeValue = new System.DateTime(2012, 4, 6, 0, 0, 0, 0);
            this.ENCRYPT_PWD.DoubleValue = 0;
            this.ENCRYPT_PWD.EditValue = "";
            this.ENCRYPT_PWD.Location = new System.Drawing.Point(8, 15);
            this.ENCRYPT_PWD.LookupAdapter = null;
            this.ENCRYPT_PWD.Name = "ENCRYPT_PWD";
            this.ENCRYPT_PWD.Nullable = true;
            this.ENCRYPT_PWD.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.ENCRYPT_PWD.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.ENCRYPT_PWD.PromptText = "암호화 비밀번호";
            isLanguageElement1.Default = "Encrypt Password";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "암호화 비밀번호";
            isLanguageElement1.TL2_CN = "";
            isLanguageElement1.TL3_VN = "";
            isLanguageElement1.TL4_JP = "";
            isLanguageElement1.TL5_XAA = "";
            this.ENCRYPT_PWD.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.ENCRYPT_PWD.PromptWidth = 130;
            this.ENCRYPT_PWD.Size = new System.Drawing.Size(265, 21);
            this.ENCRYPT_PWD.TabIndex = 0;
            this.ENCRYPT_PWD.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.ENCRYPT_PWD.TextValue = "";
            this.ENCRYPT_PWD.UseSystemPassword = true;
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "211.168.59.26";
            this.isOraConnection1.OraPassword = "infoflex";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "FXCDB";
            this.isOraConnection1.OraUserId = "APPS";
            // 
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
            // 
            // btnCLOSE
            // 
            this.btnCLOSE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.btnCLOSE.ButtonText = "OK";
            isLanguageElement4.Default = "OK";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "확인";
            isLanguageElement4.TL2_CN = "";
            isLanguageElement4.TL3_VN = "";
            isLanguageElement4.TL4_JP = "";
            isLanguageElement4.TL5_XAA = "";
            this.btnCLOSE.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.btnCLOSE.Location = new System.Drawing.Point(121, 99);
            this.btnCLOSE.Name = "btnCLOSE";
            this.btnCLOSE.Size = new System.Drawing.Size(75, 25);
            this.btnCLOSE.TabIndex = 1;
            this.btnCLOSE.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.btnOK_ButtonClick);
            // 
            // igbCONFIRM_INFOMATION
            // 
            this.igbCONFIRM_INFOMATION.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.igbCONFIRM_INFOMATION.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.igbCONFIRM_INFOMATION.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.igbCONFIRM_INFOMATION.Controls.Add(this.ENCRYPT_PWD);
            this.igbCONFIRM_INFOMATION.Controls.Add(this.CHK_ENCRYPT_PWD);
            this.igbCONFIRM_INFOMATION.Location = new System.Drawing.Point(8, 8);
            this.igbCONFIRM_INFOMATION.Name = "igbCONFIRM_INFOMATION";
            this.igbCONFIRM_INFOMATION.PromptText = "Confirm Infomation";
            isLanguageElement2.Default = "Confirm Infomation";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "승인 정보";
            isLanguageElement2.TL2_CN = "";
            isLanguageElement2.TL3_VN = "";
            isLanguageElement2.TL4_JP = "";
            isLanguageElement2.TL5_XAA = "";
            this.igbCONFIRM_INFOMATION.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.igbCONFIRM_INFOMATION.PromptVisible = false;
            this.igbCONFIRM_INFOMATION.Size = new System.Drawing.Size(290, 76);
            this.igbCONFIRM_INFOMATION.TabIndex = 0;
            // 
            // btnCANCEL
            // 
            this.btnCANCEL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.btnCANCEL.ButtonText = "Cancel";
            isLanguageElement3.Default = "Cancel";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "취소";
            isLanguageElement3.TL2_CN = "";
            isLanguageElement3.TL3_VN = "";
            isLanguageElement3.TL4_JP = "";
            isLanguageElement3.TL5_XAA = "";
            this.btnCANCEL.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.btnCANCEL.Location = new System.Drawing.Point(202, 99);
            this.btnCANCEL.Name = "btnCANCEL";
            this.btnCANCEL.Size = new System.Drawing.Size(75, 25);
            this.btnCANCEL.TabIndex = 2;
            this.btnCANCEL.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.btnCANCEL_ButtonClick);
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.igbCONFIRM_INFOMATION)).EndInit();
            this.igbCONFIRM_INFOMATION.ResumeLayout(false);

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISButton btnCLOSE;
        private InfoSummit.Win.ControlAdv.ISGroupBox igbCONFIRM_INFOMATION;
        private InfoSummit.Win.ControlAdv.ISEditAdv CHK_ENCRYPT_PWD;
        private InfoSummit.Win.ControlAdv.ISEditAdv ENCRYPT_PWD;
        private InfoSummit.Win.ControlAdv.ISButton btnCANCEL;
    }
}

