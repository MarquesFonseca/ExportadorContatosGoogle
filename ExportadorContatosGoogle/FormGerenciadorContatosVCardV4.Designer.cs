namespace ExportadorContatosGoogle
{
    partial class FormGerenciadorContatosVCardV4
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormGerenciadorContatosVCardV4));
            this.Menu1 = new System.Windows.Forms.MenuStrip();
            this.testeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuArquivoExcelXLSX = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuArquivoExcelXLS = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuArquivoExcelCSV = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuArquivoExcelXML = new System.Windows.Forms.ToolStripMenuItem();
            this.AbrirTamplates = new System.Windows.Forms.OpenFileDialog();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.Menu1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Menu1
            // 
            this.Menu1.Dock = System.Windows.Forms.DockStyle.None;
            this.Menu1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.testeToolStripMenuItem});
            this.Menu1.Location = new System.Drawing.Point(13, 7);
            this.Menu1.MdiWindowListItem = this.testeToolStripMenuItem;
            this.Menu1.Name = "Menu1";
            this.Menu1.Size = new System.Drawing.Size(288, 24);
            this.Menu1.TabIndex = 10;
            this.Menu1.Text = "menuStrip1";
            // 
            // testeToolStripMenuItem
            // 
            this.testeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuArquivoExcelXLSX,
            this.MenuArquivoExcelXLS,
            this.MenuArquivoExcelCSV,
            this.MenuArquivoExcelXML});
            this.testeToolStripMenuItem.Image = global::ExportadorContatosGoogle.Properties.Resources.import_24;
            this.testeToolStripMenuItem.Name = "testeToolStripMenuItem";
            this.testeToolStripMenuItem.Size = new System.Drawing.Size(188, 20);
            this.testeToolStripMenuItem.Text = "Buscar arquivo para Importar";
            // 
            // MenuArquivoExcelXLSX
            // 
            this.MenuArquivoExcelXLSX.Image = global::ExportadorContatosGoogle.Properties.Resources.file_xls;
            this.MenuArquivoExcelXLSX.Name = "MenuArquivoExcelXLSX";
            this.MenuArquivoExcelXLSX.Size = new System.Drawing.Size(188, 22);
            this.MenuArquivoExcelXLSX.Text = "Arquivo Excel (*.xlsx)";
            this.MenuArquivoExcelXLSX.Click += new System.EventHandler(this.MenuArquivoExcelXLSX_Click);
            // 
            // MenuArquivoExcelXLS
            // 
            this.MenuArquivoExcelXLS.Image = global::ExportadorContatosGoogle.Properties.Resources.file_xls;
            this.MenuArquivoExcelXLS.Name = "MenuArquivoExcelXLS";
            this.MenuArquivoExcelXLS.Size = new System.Drawing.Size(188, 22);
            this.MenuArquivoExcelXLS.Text = "Arquivo Excel ( *.xls )";
            this.MenuArquivoExcelXLS.Click += new System.EventHandler(this.MenuArquivoExcelXLS_Click);
            // 
            // MenuArquivoExcelCSV
            // 
            this.MenuArquivoExcelCSV.Image = global::ExportadorContatosGoogle.Properties.Resources.doc_excel_csv;
            this.MenuArquivoExcelCSV.Name = "MenuArquivoExcelCSV";
            this.MenuArquivoExcelCSV.Size = new System.Drawing.Size(188, 22);
            this.MenuArquivoExcelCSV.Text = "Arquivo Excel ( *.csv )";
            this.MenuArquivoExcelCSV.Visible = false;
            this.MenuArquivoExcelCSV.Click += new System.EventHandler(this.MenuArquivoExcelCSV_Click);
            // 
            // MenuArquivoExcelXML
            // 
            this.MenuArquivoExcelXML.Image = global::ExportadorContatosGoogle.Properties.Resources.file_xml;
            this.MenuArquivoExcelXML.Name = "MenuArquivoExcelXML";
            this.MenuArquivoExcelXML.Size = new System.Drawing.Size(188, 22);
            this.MenuArquivoExcelXML.Text = "Arquivo XML ( *.xml )";
            this.MenuArquivoExcelXML.Visible = false;
            this.MenuArquivoExcelXML.Click += new System.EventHandler(this.MenuArquivoExcelXML_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 38);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(734, 451);
            this.dataGridView1.TabIndex = 11;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.Menu1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(734, 38);
            this.panel1.TabIndex = 12;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(521, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(203, 27);
            this.button1.TabIndex = 12;
            this.button1.Text = "&Exportar para google";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // FormGerenciadorContatosVCardV4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(734, 489);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormGerenciadorContatosVCardV4";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gerenciador de exportação de contatos V4.0";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FormGerenciadorContatosVCardV4_Load);
            this.Menu1.ResumeLayout(false);
            this.Menu1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.MenuStrip Menu1;
        private System.Windows.Forms.ToolStripMenuItem testeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem MenuArquivoExcelXLSX;
        private System.Windows.Forms.ToolStripMenuItem MenuArquivoExcelXLS;
        private System.Windows.Forms.ToolStripMenuItem MenuArquivoExcelCSV;
        private System.Windows.Forms.ToolStripMenuItem MenuArquivoExcelXML;
        private System.Windows.Forms.OpenFileDialog AbrirTamplates;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}