
namespace Calculate_line_ratio
{
    partial class Form1
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
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnFileLoad = new System.Windows.Forms.Button();
            this.lblFileName = new System.Windows.Forms.Label();
            this.dgv1 = new System.Windows.Forms.DataGridView();
            this.btnCalLineRatio = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnFileLoad
            // 
            this.btnFileLoad.Location = new System.Drawing.Point(25, 51);
            this.btnFileLoad.Name = "btnFileLoad";
            this.btnFileLoad.Size = new System.Drawing.Size(118, 23);
            this.btnFileLoad.TabIndex = 0;
            this.btnFileLoad.Text = "파일 불러오기";
            this.btnFileLoad.UseVisualStyleBackColor = true;
            this.btnFileLoad.Click += new System.EventHandler(this.fBtnFileLoad_Click);
            // 
            // lblFileName
            // 
            this.lblFileName.AutoSize = true;
            this.lblFileName.Location = new System.Drawing.Point(158, 56);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(41, 12);
            this.lblFileName.TabIndex = 1;
            this.lblFileName.Text = "파일명";
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 100);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowTemplate.Height = 23;
            this.dgv1.Size = new System.Drawing.Size(776, 338);
            this.dgv1.TabIndex = 2;
            // 
            // btnCalLineRatio
            // 
            this.btnCalLineRatio.Location = new System.Drawing.Point(471, 51);
            this.btnCalLineRatio.Name = "btnCalLineRatio";
            this.btnCalLineRatio.Size = new System.Drawing.Size(145, 23);
            this.btnCalLineRatio.TabIndex = 3;
            this.btnCalLineRatio.Text = "Calculate lineRatio";
            this.btnCalLineRatio.UseVisualStyleBackColor = true;
            this.btnCalLineRatio.Click += new System.EventHandler(this.btnCalLineRatio_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnCalLineRatio);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.lblFileName);
            this.Controls.Add(this.btnFileLoad);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnFileLoad;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.DataGridView dgv1;
        private System.Windows.Forms.Button btnCalLineRatio;
    }
}

