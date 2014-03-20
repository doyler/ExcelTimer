namespace TimeLog
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnStop = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.myTimer = new System.Windows.Forms.Timer(this.components);
            this.theTime = new System.Windows.Forms.TextBox();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnLock = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(93, 41);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnStop
            // 
            this.btnStop.Enabled = false;
            this.btnStop.Location = new System.Drawing.Point(93, 70);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(75, 23);
            this.btnStop.TabIndex = 1;
            this.btnStop.Text = "Stop";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(93, 99);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // myTimer
            // 
            this.myTimer.Interval = 1000;
            this.myTimer.Tick += new System.EventHandler(this.myTimer_Tick);
            // 
            // theTime
            // 
            this.theTime.Enabled = false;
            this.theTime.Location = new System.Drawing.Point(93, 15);
            this.theTime.Name = "theTime";
            this.theTime.Size = new System.Drawing.Size(75, 20);
            this.theTime.TabIndex = 4;
            this.theTime.Text = "00:00:00";
            // 
            // btnEdit
            // 
            this.btnEdit.Enabled = false;
            this.btnEdit.Location = new System.Drawing.Point(12, 12);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 23);
            this.btnEdit.TabIndex = 5;
            this.btnEdit.Text = "Edit";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnLock
            // 
            this.btnLock.Enabled = false;
            this.btnLock.Location = new System.Drawing.Point(174, 12);
            this.btnLock.Name = "btnLock";
            this.btnLock.Size = new System.Drawing.Size(75, 23);
            this.btnLock.TabIndex = 6;
            this.btnLock.Text = "Lock";
            this.btnLock.UseVisualStyleBackColor = true;
            this.btnLock.Click += new System.EventHandler(this.btnLock_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(263, 130);
            this.Controls.Add(this.btnLock);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.theTime);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.btnStart);
            this.Name = "Form1";
            this.Text = "Excel Timer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Timer myTimer;
        private System.Windows.Forms.TextBox theTime;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Button btnLock;
    }
}

