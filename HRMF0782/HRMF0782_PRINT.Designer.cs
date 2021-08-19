namespace HRMF0782
{
    partial class HRMF0782_PRINT
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
            InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo oraConnectionInfo1 = new InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.ibtCANCEL = new InfoSummit.Win.ControlAdv.ISButton();
            this.ibtPRINTING = new InfoSummit.Win.ControlAdv.ISButton();
            this.isGroupBox1 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.CB_PRINT_1 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.CB_PRINT_2 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox1)).BeginInit();
            this.isGroupBox1.SuspendLayout();
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "59.16.125.10";
            this.isOraConnection1.OraPassword = "erp0201";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "MESORA";
            this.isOraConnection1.OraUserId = "APPS";
            // 
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
            // 
            // ibtCANCEL
            // 
            this.ibtCANCEL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.ibtCANCEL.ButtonText = "취소";
            isLanguageElement1.Default = "Cancel";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "취소";
            isLanguageElement1.TL2_CN = "";
            isLanguageElement1.TL3_VN = "";
            isLanguageElement1.TL4_JP = "";
            isLanguageElement1.TL5_XAA = "";
            this.ibtCANCEL.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.ibtCANCEL.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.ibtCANCEL.ForeColor = System.Drawing.Color.Blue;
            // 
            // HRMF0782_PRINT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(293, 171);
            this.ControlBox = false;
            this.Controls.Add(this.isGroupBox1);
            this.Controls.Add(this.ibtCANCEL);
            this.Controls.Add(this.ibtPRINTING);
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HRMF0782_PRINT";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "인쇄 선택";
            this.Load += new System.EventHandler(this.HRMF0782_PRINT_Load);
            this.Shown += new System.EventHandler(this.HRMF0782_PRINT_Shown);
            this.ibtCANCEL.Location = new System.Drawing.Point(183, 122);
            this.ibtCANCEL.Name = "ibtCANCEL";
            this.ibtCANCEL.Size = new System.Drawing.Size(98, 27);
            this.ibtCANCEL.TabIndex = 16;
            this.ibtCANCEL.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.ibtCANCEL.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.ibtCANCEL_ButtonClick);
            // 
            // ibtPRINTING
            // 
            this.ibtPRINTING.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.ibtPRINTING.ButtonText = "인쇄";
            isLanguageElement5.Default = "Printing";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "인쇄";
            isLanguageElement5.TL2_CN = "";
            isLanguageElement5.TL3_VN = "";
            isLanguageElement5.TL4_JP = "";
            isLanguageElement5.TL5_XAA = "";
            this.ibtPRINTING.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.ibtPRINTING.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.ibtPRINTING.ForeColor = System.Drawing.Color.Blue;
            this.ibtPRINTING.Location = new System.Drawing.Point(79, 122);
            this.ibtPRINTING.Name = "ibtPRINTING";
            this.ibtPRINTING.Size = new System.Drawing.Size(98, 27);
            this.ibtPRINTING.TabIndex = 15;
            this.ibtPRINTING.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.ibtPRINTING.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.ibtPRINTING_ButtonClick);
            // 
            // isGroupBox1
            // 
            this.isGroupBox1.AppInterfaceAdv = null;
            this.isGroupBox1.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.isGroupBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isGroupBox1.Controls.Add(this.CB_PRINT_1);
            this.isGroupBox1.Controls.Add(this.CB_PRINT_2);
            this.isGroupBox1.Location = new System.Drawing.Point(8, 8);
            this.isGroupBox1.Name = "isGroupBox1";
            this.isGroupBox1.PromptText = "isGroupBox1";
            isLanguageElement4.Default = "isGroupBox1";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = null;
            isLanguageElement4.TL2_CN = null;
            isLanguageElement4.TL3_VN = null;
            isLanguageElement4.TL4_JP = null;
            isLanguageElement4.TL5_XAA = null;
            this.isGroupBox1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.isGroupBox1.PromptVisible = false;
            this.isGroupBox1.Size = new System.Drawing.Size(277, 92);
            this.isGroupBox1.TabIndex = 18;
            // 
            // CB_PRINT_1
            // 
            this.CB_PRINT_1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CB_PRINT_1.CheckBoxValue = "N";
            this.CB_PRINT_1.CheckedString = "Y";
            this.CB_PRINT_1.DataAdapter = null;
            this.CB_PRINT_1.DataColumn = null;
            this.CB_PRINT_1.Location = new System.Drawing.Point(33, 23);
            this.CB_PRINT_1.Name = "CB_PRINT_1";
            this.CB_PRINT_1.PromptText = "지방소득세특별징수 계산서";
            isLanguageElement2.Default = "지방소득세특별징수 계산서";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "지방소득세특별징수 계산서";
            isLanguageElement2.TL2_CN = null;
            isLanguageElement2.TL3_VN = null;
            isLanguageElement2.TL4_JP = null;
            isLanguageElement2.TL5_XAA = null;
            this.CB_PRINT_1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.CB_PRINT_1.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_1.TabIndex = 189;
            this.CB_PRINT_1.UncheckedString = "N";
            // 
            // CB_PRINT_2
            // 
            this.CB_PRINT_2.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CB_PRINT_2.CheckBoxValue = "N";
            this.CB_PRINT_2.CheckedString = "Y";
            this.CB_PRINT_2.DataAdapter = null;
            this.CB_PRINT_2.DataColumn = null;
            this.CB_PRINT_2.Location = new System.Drawing.Point(33, 50);
            this.CB_PRINT_2.Name = "CB_PRINT_2";
            this.CB_PRINT_2.PromptText = "지방소득세특별징수 납부서";
            isLanguageElement3.Default = "지방소득세특별징수 납부서";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "지방소득세특별징수 납부서";
            isLanguageElement3.TL2_CN = null;
            isLanguageElement3.TL3_VN = null;
            isLanguageElement3.TL4_JP = null;
            isLanguageElement3.TL5_XAA = null;
            this.CB_PRINT_2.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.CB_PRINT_2.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_2.TabIndex = 189;
            this.CB_PRINT_2.UncheckedString = "N";
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox1)).EndInit();
            this.isGroupBox1.ResumeLayout(false);

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISButton ibtCANCEL;
        private InfoSummit.Win.ControlAdv.ISButton ibtPRINTING;
        private InfoSummit.Win.ControlAdv.ISGroupBox isGroupBox1;
        private InfoSummit.Win.ControlAdv.ISCheckBoxAdv CB_PRINT_1;
        private InfoSummit.Win.ControlAdv.ISCheckBoxAdv CB_PRINT_2;
    }
}

