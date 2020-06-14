using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExportadorContatosGoogle
{
    public partial class FormPlan : Form
    //public partial class FormPlan : DevExpress.XtraEditors.XtraForm
    {
        public string retorno = "";
        public bool cancelado = false;
        private List<DataTable> ListaPlan = new List<DataTable>();
        private List<string> ListaNomePlan = new List<string>();

        public FormPlan()
        {
            InitializeComponent();
        }

        public FormPlan(List<DataTable> listaPlan, List<string> listaNomePlan, string nomeArquivoBuscado)
        {
            InitializeComponent();
            ListaPlan = listaPlan;
            ListaNomePlan = listaNomePlan;
            LbnTituloFormulario.Text = RetornaNomeArquivo(nomeArquivoBuscado);
            this.Text = LbnTituloFormulario.Text;

            //this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
            this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
        }

        private void FormPlan_Load(object sender, EventArgs e)
        {
            foreach (string item in ListaNomePlan)
            {
                listBox1.Items.Add(item.Replace("$", ""));
                tabControl1.TabPages.Add(item.Replace("$", ""));
            }
            listBox1.Focus();
            listBox1.SelectedIndex = 0;
            tabControl1.SelectedIndex = 0;
        }

        private string RetornaNomeArquivo(string nomeArquivoBuscado)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(nomeArquivoBuscado);
            //Mostra o nome do arquivo
            string fileName = fileInfo.Name;
            //Mostra a extensão do arquivo
            string fileExtension = fileInfo.Extension;
            //Mostra o caminho completo do arquivo junto com o nome
            string fileFullName = fileInfo.FullName;
            return string.Format("Selecione uma plan do arquivo: '{0}'", fileName);
        }

        private void BtnConfirmar_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count > 0)
                retorno = listBox1.SelectedItem.ToString();
            else retorno = "";
            this.Close();
        }

        private void BtnLimpar_Click(object sender, EventArgs e)
        {

        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            cancelado = true;
            this.Close();
        }

        private void FormPlan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                BtnConfirmar_Click(null, null);
            }
            if (e.KeyData == Keys.Escape)
            {
                this.Close();
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ListaPlan[listBox1.SelectedIndex];
            tabControl1.SelectedIndex = listBox1.SelectedIndex;
            listBox1.Focus();
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BtnConfirmar_Click(null, null);
        }

        private void tabControl1_MouseClick(object sender, MouseEventArgs e)
        {
            dataGridView1.DataSource = ListaPlan[tabControl1.SelectedIndex];
            listBox1.SelectedIndex = tabControl1.SelectedIndex;
        }
    }
}
