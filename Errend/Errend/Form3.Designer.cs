namespace Errend
{
    partial class Form3
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.update = new System.Windows.Forms.Button();
            this.viewLines = new System.Windows.Forms.Button();
            this.viewSender = new System.Windows.Forms.Button();
            this.viewReceiver = new System.Windows.Forms.Button();
            this.viewVessels = new System.Windows.Forms.Button();
            this.viewPortT = new System.Windows.Forms.Button();
            this.viewCountryPort = new System.Windows.Forms.Button();
            this.viewCargo = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(960, 400);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CelEndEdite);
            // 
            // update
            // 
            this.update.Location = new System.Drawing.Point(900, 425);
            this.update.Name = "update";
            this.update.Size = new System.Drawing.Size(75, 25);
            this.update.TabIndex = 1;
            this.update.Text = "Обновить";
            this.update.UseVisualStyleBackColor = true;
            this.update.Click += new System.EventHandler(this.update_Click);
            // 
            // viewLines
            // 
            this.viewLines.Location = new System.Drawing.Point(10, 425);
            this.viewLines.Name = "viewLines";
            this.viewLines.Size = new System.Drawing.Size(75, 25);
            this.viewLines.TabIndex = 2;
            this.viewLines.Text = "Линии";
            this.viewLines.UseVisualStyleBackColor = true;
            this.viewLines.Click += new System.EventHandler(this.viewLines_Click);
            // 
            // viewSender
            // 
            this.viewSender.Location = new System.Drawing.Point(90, 425);
            this.viewSender.Name = "viewSender";
            this.viewSender.Size = new System.Drawing.Size(85, 25);
            this.viewSender.TabIndex = 3;
            this.viewSender.Text = "Отправители";
            this.viewSender.UseVisualStyleBackColor = true;
            this.viewSender.Click += new System.EventHandler(this.viewSender_Click);
            // 
            // viewReceiver
            // 
            this.viewReceiver.Location = new System.Drawing.Point(180, 425);
            this.viewReceiver.Name = "viewReceiver";
            this.viewReceiver.Size = new System.Drawing.Size(75, 25);
            this.viewReceiver.TabIndex = 4;
            this.viewReceiver.Text = "Получатели";
            this.viewReceiver.UseVisualStyleBackColor = true;
            this.viewReceiver.Click += new System.EventHandler(this.viewReceiver_Click);
            // 
            // viewVessels
            // 
            this.viewVessels.Location = new System.Drawing.Point(260, 425);
            this.viewVessels.Name = "viewVessels";
            this.viewVessels.Size = new System.Drawing.Size(75, 25);
            this.viewVessels.TabIndex = 5;
            this.viewVessels.Text = "Суда";
            this.viewVessels.UseVisualStyleBackColor = true;
            this.viewVessels.Click += new System.EventHandler(this.viewVessels_Click);
            // 
            // viewPortT
            // 
            this.viewPortT.Location = new System.Drawing.Point(340, 425);
            this.viewPortT.Name = "viewPortT";
            this.viewPortT.Size = new System.Drawing.Size(105, 25);
            this.viewPortT.TabIndex = 6;
            this.viewPortT.Text = "Порты перевалки";
            this.viewPortT.UseVisualStyleBackColor = true;
            this.viewPortT.Click += new System.EventHandler(this.viewPortT_Click);
            // 
            // viewCountryPort
            // 
            this.viewCountryPort.Location = new System.Drawing.Point(450, 425);
            this.viewCountryPort.Name = "viewCountryPort";
            this.viewCountryPort.Size = new System.Drawing.Size(105, 25);
            this.viewCountryPort.TabIndex = 7;
            this.viewCountryPort.Text = "Порты выгрузки";
            this.viewCountryPort.UseVisualStyleBackColor = true;
            this.viewCountryPort.Click += new System.EventHandler(this.viewCountryPort_Click);
            // 
            // viewCargo
            // 
            this.viewCargo.Location = new System.Drawing.Point(561, 425);
            this.viewCargo.Name = "viewCargo";
            this.viewCargo.Size = new System.Drawing.Size(75, 25);
            this.viewCargo.TabIndex = 8;
            this.viewCargo.Text = "Грузы";
            this.viewCargo.UseVisualStyleBackColor = true;
            this.viewCargo.Click += new System.EventHandler(this.viewCargo_Click);
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 461);
            this.Controls.Add(this.viewCargo);
            this.Controls.Add(this.viewCountryPort);
            this.Controls.Add(this.viewPortT);
            this.Controls.Add(this.viewVessels);
            this.Controls.Add(this.viewReceiver);
            this.Controls.Add(this.viewSender);
            this.Controls.Add(this.viewLines);
            this.Controls.Add(this.update);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form3";
            this.Text = "База данных";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form3_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button update;
        private System.Windows.Forms.Button viewLines;
        private System.Windows.Forms.Button viewSender;
        private System.Windows.Forms.Button viewReceiver;
        private System.Windows.Forms.Button viewVessels;
        private System.Windows.Forms.Button viewPortT;
        private System.Windows.Forms.Button viewCountryPort;
        private System.Windows.Forms.Button viewCargo;
    }
}