
namespace TAXREG
{
    partial class TAXREG
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
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.isDataAdapter1 = new InfoSummit.Win.ControlAdv.ISDataAdapter(this.components);
            this.isButton1 = new InfoSummit.Win.ControlAdv.ISButton();
            this.isEditAdv1 = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.SuspendLayout();
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
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
            // isDataAdapter1
            // 
            this.isDataAdapter1.CancelMember.Cancel = false;
            this.isDataAdapter1.CancelMember.Member = null;
            this.isDataAdapter1.CancelMember.Prompt = null;
            this.isDataAdapter1.CancelMember.TabIndex = -1;
            this.isDataAdapter1.CancelMember.ValueItem = null;
            this.isDataAdapter1.CancelUpdateFilterString = null;
            this.isDataAdapter1.CancelUpdateRow = null;
            this.isDataAdapter1.DataTransaction = null;
            this.isDataAdapter1.FocusedControl = null;
            // 
            // TAXREG
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(954, 601);
            this.Controls.Add(this.isEditAdv1);
            this.Controls.Add(this.isButton1);
            this.Name = "TAXREG";
            this.Padding = new System.Windows.Forms.Padding(2);
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "970X640";
            this.isDataAdapter1.MasterAdapter = null;
            this.isDataAdapter1.OraConnection = this.isOraConnection1;
            this.isDataAdapter1.OraDelete = null;
            this.isDataAdapter1.OraInsert = null;
            this.isDataAdapter1.OraOwner = "APPS";
            this.isDataAdapter1.OraPackage = null;
            this.isDataAdapter1.OraSelect = null;
            this.isDataAdapter1.OraSelectData = null;
            this.isDataAdapter1.OraUpdate = null;
            this.isDataAdapter1.WizardOwner = null;
            this.isDataAdapter1.WizardProcedure = null;
            this.isDataAdapter1.WizardTableName = null;
            // 
            // isButton1
            // 
            this.isButton1.AppInterfaceAdv = null;
            this.isButton1.ButtonText = "isButton1";
            isLanguageElement1.Default = "isButton1";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = null;
            isLanguageElement1.TL2_CN = null;
            isLanguageElement1.TL3_VN = null;
            isLanguageElement1.TL4_JP = null;
            isLanguageElement1.TL5_XAA = null;
            this.isButton1.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.isButton1.Location = new System.Drawing.Point(208, 157);
            this.isButton1.Name = "isButton1";
            this.isButton1.Size = new System.Drawing.Size(487, 63);
            this.isButton1.TabIndex = 0;
            // 
            // isEditAdv1
            // 
            this.isEditAdv1.AppInterfaceAdv = null;
            this.isEditAdv1.ComboBoxValue = "";
            this.isEditAdv1.ComboData = null;
            this.isEditAdv1.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.isEditAdv1.DataAdapter = null;
            this.isEditAdv1.DataColumn = null;
            this.isEditAdv1.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.isEditAdv1.DoubleValue = 0D;
            this.isEditAdv1.EditValue = "";
            this.isEditAdv1.Location = new System.Drawing.Point(170, 113);
            this.isEditAdv1.LookupAdapter = null;
            this.isEditAdv1.Name = "isEditAdv1";
            this.isEditAdv1.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.isEditAdv1.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.isEditAdv1.PromptText = null;
            this.isEditAdv1.Size = new System.Drawing.Size(587, 21);
            this.isEditAdv1.TabIndex = 1;
            this.isEditAdv1.TextValue = "";
            this.ResumeLayout(false);

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISDataAdapter isDataAdapter1;
        private InfoSummit.Win.ControlAdv.ISEditAdv isEditAdv1;
        private InfoSummit.Win.ControlAdv.ISButton isButton1;
    }
}

