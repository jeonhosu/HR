﻿namespace HRMF0304
{
    partial class XLExport
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
            this.isProgressBar1 = new InfoSummit.Win.ControlAdv.ISProgressBar();
            this.SuspendLayout();
            // 
            // isProgressBar1
            // 
            this.isProgressBar1.BarFillPercent = 0F;
            this.isProgressBar1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.isProgressBar1.Location = new System.Drawing.Point(5, 5);
            this.isProgressBar1.Name = "isProgressBar1";
            this.isProgressBar1.Size = new System.Drawing.Size(290, 15);
            this.isProgressBar1.StepSize = 0F;
            this.isProgressBar1.TabIndex = 0;
            this.isProgressBar1.Text = "isProgressBar1";
            // 
            // XLExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(300, 25);
            this.Controls.Add(this.isProgressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "XLExport";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "XLExport";
            this.ResumeLayout(false);

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISProgressBar isProgressBar1;
    }
}

