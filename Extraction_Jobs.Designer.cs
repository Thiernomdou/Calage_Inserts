namespace Calage_Inserts
{
    partial class Extraction_Jobs
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
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.JOB_NB = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CAVITY_JOB_NB = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LB_LOGI_SKU = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CREATED_BY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CREATION_DATE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LAST_UPDATE_BY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LAST_UPDATE_DATE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SUBSTRATE_CODE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CAVITY_DISPATCH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button3 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Cambria", 14.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(416, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(243, 22);
            this.label2.TabIndex = 16;
            this.label2.Text = "Affichage des données Combo";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Gainsboro;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(897, 13);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(110, 34);
            this.button1.TabIndex = 17;
            this.button1.Text = "Retour";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.RosyBrown;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button2.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(434, 61);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(110, 34);
            this.button2.TabIndex = 18;
            this.button2.Text = "Afficher";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.JOB_NB,
            this.CAVITY_JOB_NB,
            this.LB_LOGI_SKU,
            this.CREATED_BY,
            this.CREATION_DATE,
            this.LAST_UPDATE_BY,
            this.LAST_UPDATE_DATE,
            this.SUBSTRATE_CODE,
            this.CAVITY_DISPATCH});
            this.dataGridView1.Location = new System.Drawing.Point(78, 123);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1097, 397);
            this.dataGridView1.TabIndex = 19;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // JOB_NB
            // 
            this.JOB_NB.HeaderText = "JOB_NB";
            this.JOB_NB.Name = "JOB_NB";
            this.JOB_NB.Width = 73;
            // 
            // CAVITY_JOB_NB
            // 
            this.CAVITY_JOB_NB.HeaderText = "CAVITY_JOB_NB";
            this.CAVITY_JOB_NB.Name = "CAVITY_JOB_NB";
            this.CAVITY_JOB_NB.Width = 117;
            // 
            // LB_LOGI_SKU
            // 
            this.LB_LOGI_SKU.HeaderText = "LB_LOGI_SKU";
            this.LB_LOGI_SKU.Name = "LB_LOGI_SKU";
            this.LB_LOGI_SKU.Width = 104;
            // 
            // CREATED_BY
            // 
            this.CREATED_BY.HeaderText = "CREATED_BY";
            this.CREATED_BY.Name = "CREATED_BY";
            this.CREATED_BY.Width = 103;
            // 
            // CREATION_DATE
            // 
            this.CREATION_DATE.HeaderText = "CREATION_DATE";
            this.CREATION_DATE.Name = "CREATION_DATE";
            this.CREATION_DATE.Width = 122;
            // 
            // LAST_UPDATE_BY
            // 
            this.LAST_UPDATE_BY.HeaderText = "LAST_UPDATE_BY";
            this.LAST_UPDATE_BY.Name = "LAST_UPDATE_BY";
            this.LAST_UPDATE_BY.Width = 129;
            // 
            // LAST_UPDATE_DATE
            // 
            this.LAST_UPDATE_DATE.HeaderText = "LAST_UPDATE_DATE";
            this.LAST_UPDATE_DATE.Name = "LAST_UPDATE_DATE";
            this.LAST_UPDATE_DATE.Width = 144;
            // 
            // SUBSTRATE_CODE
            // 
            this.SUBSTRATE_CODE.HeaderText = "SUBSTRATE_CODE";
            this.SUBSTRATE_CODE.Name = "SUBSTRATE_CODE";
            this.SUBSTRATE_CODE.Width = 133;
            // 
            // CAVITY_DISPATCH
            // 
            this.CAVITY_DISPATCH.HeaderText = "CAVITY_DISPATCH";
            this.CAVITY_DISPATCH.Name = "CAVITY_DISPATCH";
            this.CAVITY_DISPATCH.Width = 130;
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.RosyBrown;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button3.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(576, 61);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(110, 34);
            this.button3.TabIndex = 20;
            this.button3.Text = "Extraire";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 19);
            this.label1.TabIndex = 21;
            this.label1.Text = "JOB_NB";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(79, 60);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(138, 33);
            this.textBox1.TabIndex = 22;
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.Color.RosyBrown;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button4.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.Location = new System.Drawing.Point(232, 60);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(110, 34);
            this.button4.TabIndex = 23;
            this.button4.Text = "Rechercher";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Extraction_Jobs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkGray;
            this.ClientSize = new System.Drawing.Size(1226, 532);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Name = "Extraction_Jobs";
            this.Text = "Extraction_Jobs";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Extraction_Jobs_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridViewTextBoxColumn JOB_NB;
        private System.Windows.Forms.DataGridViewTextBoxColumn CAVITY_JOB_NB;
        private System.Windows.Forms.DataGridViewTextBoxColumn LB_LOGI_SKU;
        private System.Windows.Forms.DataGridViewTextBoxColumn CREATED_BY;
        private System.Windows.Forms.DataGridViewTextBoxColumn CREATION_DATE;
        private System.Windows.Forms.DataGridViewTextBoxColumn LAST_UPDATE_BY;
        private System.Windows.Forms.DataGridViewTextBoxColumn LAST_UPDATE_DATE;
        private System.Windows.Forms.DataGridViewTextBoxColumn SUBSTRATE_CODE;
        private System.Windows.Forms.DataGridViewTextBoxColumn CAVITY_DISPATCH;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button4;
    }
}