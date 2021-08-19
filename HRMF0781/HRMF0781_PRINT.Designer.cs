namespace HRMF0781
{
    partial class HRMF0781_PRINT
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
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement9 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement8 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement7 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.ibtCANCEL = new InfoSummit.Win.ControlAdv.ISButton();
            this.ibtPRINTING = new InfoSummit.Win.ControlAdv.ISButton();
            this.isGroupBox1 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.CB_PRINT_5 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.CB_PRINT_1 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.CB_PRINT_2 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.CB_PRINT_3 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.CB_PRINT_4 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.CB_PRINT_6 = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
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
            // HRMF0781_PRINT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(293, 254);
            this.ControlBox = false;
            this.Controls.Add(this.isGroupBox1);
            this.Controls.Add(this.ibtCANCEL);
            this.Controls.Add(this.ibtPRINTING);
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HRMF0781_PRINT";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "인쇄 선택";
            this.Load += new System.EventHandler(this.HRMF0781_PRINT_Load);
            this.Shown += new System.EventHandler(this.HRMF0781_PRINT_Shown);
            this.ibtCANCEL.Location = new System.Drawing.Point(177, 209);
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
            isLanguageElement9.Default = "Printing";
            isLanguageElement9.SiteName = null;
            isLanguageElement9.TL1_KR = "인쇄";
            isLanguageElement9.TL2_CN = "";
            isLanguageElement9.TL3_VN = "";
            isLanguageElement9.TL4_JP = "";
            isLanguageElement9.TL5_XAA = "";
            this.ibtPRINTING.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement9});
            this.ibtPRINTING.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.ibtPRINTING.ForeColor = System.Drawing.Color.Blue;
            this.ibtPRINTING.Location = new System.Drawing.Point(73, 209);
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
            this.isGroupBox1.Controls.Add(this.CB_PRINT_6);
            this.isGroupBox1.Controls.Add(this.CB_PRINT_5);
            this.isGroupBox1.Controls.Add(this.CB_PRINT_1);
            this.isGroupBox1.Controls.Add(this.CB_PRINT_2);
            this.isGroupBox1.Controls.Add(this.CB_PRINT_3);
            this.isGroupBox1.Controls.Add(this.CB_PRINT_4);
            this.isGroupBox1.Location = new System.Drawing.Point(8, 8);
            this.isGroupBox1.Name = "isGroupBox1";
            this.isGroupBox1.PromptText = "isGroupBox1";
            isLanguageElement8.Default = "isGroupBox1";
            isLanguageElement8.SiteName = null;
            isLanguageElement8.TL1_KR = null;
            isLanguageElement8.TL2_CN = null;
            isLanguageElement8.TL3_VN = null;
            isLanguageElement8.TL4_JP = null;
            isLanguageElement8.TL5_XAA = null;
            this.isGroupBox1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement8});
            this.isGroupBox1.PromptVisible = false;
            this.isGroupBox1.Size = new System.Drawing.Size(277, 195);
            this.isGroupBox1.TabIndex = 18;
            // 
            // CB_PRINT_5
            // 
            this.CB_PRINT_5.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CB_PRINT_5.CheckBoxValue = "N";
            this.CB_PRINT_5.CheckedString = "Y";
            this.CB_PRINT_5.DataAdapter = null;
            this.CB_PRINT_5.DataColumn = null;
            this.CB_PRINT_5.Location = new System.Drawing.Point(33, 131);
            this.CB_PRINT_5.Name = "CB_PRINT_5";
            this.CB_PRINT_5.PromptText = "소득세납부서[기타소득]";
            isLanguageElement3.Default = "소득세납부서[기타소득]";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "소득세납부서[기타소득]";
            isLanguageElement3.TL2_CN = null;
            isLanguageElement3.TL3_VN = null;
            isLanguageElement3.TL4_JP = null;
            isLanguageElement3.TL5_XAA = null;
            this.CB_PRINT_5.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.CB_PRINT_5.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_5.TabIndex = 190;
            this.CB_PRINT_5.TabStop = false;
            this.CB_PRINT_5.UncheckedString = "N";
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
            this.CB_PRINT_1.PromptText = "원천징수이행상황신고서";
            isLanguageElement4.Default = "원천징수이행상황신고서";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "원천징수이행상황신고서";
            isLanguageElement4.TL2_CN = null;
            isLanguageElement4.TL3_VN = null;
            isLanguageElement4.TL4_JP = null;
            isLanguageElement4.TL5_XAA = null;
            this.CB_PRINT_1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.CB_PRINT_1.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_1.TabIndex = 189;
            this.CB_PRINT_1.TabStop = false;
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
            this.CB_PRINT_2.PromptText = "소득세납부서[근로소득]";
            isLanguageElement5.Default = "소득세납부서[근로소득]";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "소득세납부서[근로소득]";
            isLanguageElement5.TL2_CN = null;
            isLanguageElement5.TL3_VN = null;
            isLanguageElement5.TL4_JP = null;
            isLanguageElement5.TL5_XAA = null;
            this.CB_PRINT_2.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.CB_PRINT_2.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_2.TabIndex = 189;
            this.CB_PRINT_2.TabStop = false;
            this.CB_PRINT_2.UncheckedString = "N";
            // 
            // CB_PRINT_3
            // 
            this.CB_PRINT_3.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CB_PRINT_3.CheckBoxValue = "N";
            this.CB_PRINT_3.CheckedString = "Y";
            this.CB_PRINT_3.DataAdapter = null;
            this.CB_PRINT_3.DataColumn = null;
            this.CB_PRINT_3.Location = new System.Drawing.Point(33, 77);
            this.CB_PRINT_3.Name = "CB_PRINT_3";
            this.CB_PRINT_3.PromptText = "소득세납부서[사업소득]";
            isLanguageElement6.Default = "소득세납부서[사업소득]";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = "소득세납부서[사업소득]";
            isLanguageElement6.TL2_CN = null;
            isLanguageElement6.TL3_VN = null;
            isLanguageElement6.TL4_JP = null;
            isLanguageElement6.TL5_XAA = null;
            this.CB_PRINT_3.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.CB_PRINT_3.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_3.TabIndex = 189;
            this.CB_PRINT_3.TabStop = false;
            this.CB_PRINT_3.UncheckedString = "N";
            // 
            // CB_PRINT_4
            // 
            this.CB_PRINT_4.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CB_PRINT_4.CheckBoxValue = "N";
            this.CB_PRINT_4.CheckedString = "Y";
            this.CB_PRINT_4.DataAdapter = null;
            this.CB_PRINT_4.DataColumn = null;
            this.CB_PRINT_4.Location = new System.Drawing.Point(33, 104);
            this.CB_PRINT_4.Name = "CB_PRINT_4";
            this.CB_PRINT_4.PromptText = "소득세납부서[퇴직소득]";
            isLanguageElement7.Default = "소득세납부서[퇴직소득]";
            isLanguageElement7.SiteName = null;
            isLanguageElement7.TL1_KR = "소득세납부서[퇴직소득]";
            isLanguageElement7.TL2_CN = null;
            isLanguageElement7.TL3_VN = null;
            isLanguageElement7.TL4_JP = null;
            isLanguageElement7.TL5_XAA = null;
            this.CB_PRINT_4.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement7});
            this.CB_PRINT_4.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_4.TabIndex = 189;
            this.CB_PRINT_4.TabStop = false;
            this.CB_PRINT_4.UncheckedString = "N";
            // 
            // CB_PRINT_6
            // 
            this.CB_PRINT_6.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CB_PRINT_6.CheckBoxValue = "N";
            this.CB_PRINT_6.CheckedString = "Y";
            this.CB_PRINT_6.DataAdapter = null;
            this.CB_PRINT_6.DataColumn = null;
            this.CB_PRINT_6.Location = new System.Drawing.Point(33, 158);
            this.CB_PRINT_6.Name = "CB_PRINT_6";
            this.CB_PRINT_6.PromptText = "소득세납부서[이자소득]";
            isLanguageElement2.Default = "소득세납부서[이자소득]";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "소득세납부서[이자소득]";
            isLanguageElement2.TL2_CN = null;
            isLanguageElement2.TL3_VN = null;
            isLanguageElement2.TL4_JP = null;
            isLanguageElement2.TL5_XAA = null;
            this.CB_PRINT_6.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.CB_PRINT_6.Size = new System.Drawing.Size(202, 21);
            this.CB_PRINT_6.TabIndex = 191;
            this.CB_PRINT_6.TabStop = false;
            this.CB_PRINT_6.UncheckedString = "N";
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
        private InfoSummit.Win.ControlAdv.ISCheckBoxAdv CB_PRINT_3;
        private InfoSummit.Win.ControlAdv.ISCheckBoxAdv CB_PRINT_4;
        private InfoSummit.Win.ControlAdv.ISCheckBoxAdv CB_PRINT_5;
        private InfoSummit.Win.ControlAdv.ISCheckBoxAdv CB_PRINT_6;
    }
}

