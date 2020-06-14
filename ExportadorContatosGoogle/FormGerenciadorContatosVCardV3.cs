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
    public partial class FormGerenciadorContatosVCardV3 : Form
    {
        private string DirArquivo { get; set; }
        private DataTable TblListaAtual { get; set; }

        //private List<string> CamposGoogleIngles = new List<string> { "Selecione...", "Given Name", "Name Prefix", "Nickname", "Birthday", "Gender", "Occupation", "Notes", "Group Membership", "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2- Value", "Phone 3 - Type", "Phone 3- Value", "Phone 4 - Type", "Phone 4- Value", "Phone 5 - Type", "Phone 5- Value", "Phone 6 - Type", "Phone 6- Value", "Address 1 - Type", "Address 1 - Formatted", "Address 1 - Street", "Address 1 - City", "Address 1 - PO Box", "Address 1 - Region", "Address 1 - Postal Code", "Address 1 - Country", "Address 1 - Extended Address", "Relation 1 - Type", "Relation 1 - Value", "Event 1 - Type", "Event 1 - Value" };
        //private List<string> CamposGooglePortugues = new List<string> { "Selecione...", "Nome completo", "Tratamento", "Apelido", "Data nascimento", "Sexo", "Telefone Principal", "Telefone 2", "Telefone 3", "Telefone 4", "Telefone 5", "Telefone 6", "Email Pessoal", "Email Comercial", "Endereço Res.", "Bairro Res.", "Cidade Res.", "Estado Res.", "CEP Res.", "País Res.", "Endereço Com.", "Bairro Com.", "Cidade Com.", "Estado Com.", "CEP Com.", "País Com.", "Empresa", "Cargo/Funcão", "Nota/Histórico", "Site Pessoal", "Site Comercial", "Nome Cônjuge", "Data aniv. Cônjuge", "Nome do pai", "Nome da mãe", "Grupos" };
        private List<string> CamposGooglePortugues = new List<string> { "Selecione...", "Nome completo", "Tratamento", "Apelido", "Data nascimento", "Sexo", "Telefone Principal", "Telefone 2", "Telefone 3", "Telefone 4", "Telefone 5", "Telefone 6", "Email Pessoal", "Email Comercial", "Endereço Res.", "Endereço Res. Núm.", "Bairro Res.", "Cidade Res.", "Estado Res.", "CEP Res.", "País Res.", "Endereço Com.", "Endereço Com. Núm.", "Bairro Com.", "Cidade Com.", "Estado Com.", "CEP Com.", "País Com.", "Empresa", "Cargo/Funcão", "Nota/Histórico", "Site Pessoal", "Site Comercial", "Nome Cônjuge", "Data aniv. Cônjuge", "Nome do pai", "Nome da mãe", "Grupos" };
        private DataTable ColunaComboBoxCamposSelecaoGoogle;
        private string NomeUltimoCampo = "Event 1 - Value";
        private StringBuilder gruposLinhaAtual = new StringBuilder();
        private bool grupoLinhaJaProcessado = false;
        public FormGerenciadorContatosVCardV3()
        {
            InitializeComponent();

            //this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
            this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;

            dataGridView1.CellValueChanged += new DataGridViewCellEventHandler(dataGridView1_CellValueChanged);

            dataGridView1.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(dataGridView1_EditingControlShowing);
            dataGridView1.SelectionChanged += new EventHandler(dataGridView1_SelectionChanged);
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void FormGerenciadorContatosVCardV3_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {

        }

        int indiceCo = 0;
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                indiceCo = e.ColumnIndex;
                DataGridViewComboBoxCell cb = (DataGridViewComboBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];

                if (cb.Value != null)
                {

                    //Mensagens.Informa(string.Format("Voce selecionou: {0}", cb.EditedFormattedValue));
                    foreach (DataGridViewComboBoxCell item in dataGridView1.Rows[0].Cells)
                    {
                        string CampoAtualSelecionado = dataGridView1.Rows[0].Cells[item.ColumnIndex].FormattedValue.ToString();

                        //percorre todas as colunas que não seja essa atual...
                        if (item.ColumnIndex == e.ColumnIndex) continue;
                        if (CampoAtualSelecionado == "Grupos") continue;
                        if (item.Value == cb.Value)
                        {
                            Mensagens.Informa(string.Format("Já existe um campo selecionado para \"{0}\".\nSelecione outra opção para continuar. ", cb.FormattedValue));
                            //cb.Value = 0;                            
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

        private void MenuArquivoExcelXLSX_Click(object sender, EventArgs e)
        {
            AbrirTamplates.Title = "Buscar Arquivo Excel";
            //AbrirTamplates.InitialDirectory = DirArquivo;
            //AbrirTamplates.FileName = string.Empty;
            AbrirTamplates.DefaultExt = ".xlsx";
            AbrirTamplates.Filter = "Arquivos Excel|*.xlsx";
            AbrirTamplates.RestoreDirectory = true;

            if (AbrirTamplates.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string NomePlan = RetornaNomePlanilhaSelecionadoXLS(AbrirTamplates.FileName);
                if (string.IsNullOrEmpty(NomePlan)) return;

                try
                {
                    using (DataTable dt = new ImportarArquivos().ImportarXLSXNovo(AbrirTamplates.FileName, string.Format("{0}$", NomePlan.Replace("$", "")), 0))
                    {
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            CarregaGridView(dt);
                            return;
                        }
                        else
                        {
                            Mensagens.Alerta("Não foi possível carregar nenhum registro apartir do .xlsx informado. Por favor selecione outro arquivo.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Mensagens.Erro(string.Format("Não foi possível carregar o arquivo: {0}", ex.Message));
                }
            }
        }

        private void MenuArquivoExcelXLS_Click(object sender, EventArgs e)
        {
            AbrirTamplates.Title = "Buscar Arquivo Excel";
            //AbrirTamplates.InitialDirectory = DirArquivo;
            //AbrirTamplates.FileName = string.Empty;
            AbrirTamplates.DefaultExt = ".xls";
            AbrirTamplates.Filter = "Arquivos Excel|*.xls*";
            AbrirTamplates.RestoreDirectory = true;

            if (AbrirTamplates.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string NomePlan = RetornaNomePlanilhaSelecionadoXLS(AbrirTamplates.FileName);
                if (string.IsNullOrEmpty(NomePlan)) return;

                try
                {
                    //using (DataTable dt = new ImportarArquivos().ImportarXLS(AbrirTamplates.FileName, NomePlan))
                    using (DataTable dt = new ImportarArquivos().ImportarXLSXNovo(AbrirTamplates.FileName, string.Format("{0}$", NomePlan.Replace("$", "")), 0))
                    {
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            CarregaGridView(dt);
                            return;
                        }
                        else
                        {
                            Mensagens.Alerta("Não foi possível carregar nenhum registro apartir do .xls informado. Por favor selecione outro arquivo.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Mensagens.Erro(string.Format("Não foi possível carregar o arquivo: {0}", ex.Message));
                }
            }
        }

        private void MenuArquivoExcelCSV_Click(object sender, EventArgs e)
        {
            AbrirTamplates.Title = "Buscar Arquivo Excel";
            //AbrirTamplates.InitialDirectory = DirArquivo;
            //AbrirTamplates.FileName = string.Empty;
            AbrirTamplates.DefaultExt = ".csv";
            AbrirTamplates.Filter = "Arquivos Excel|*.csv";
            AbrirTamplates.RestoreDirectory = true;

            if (AbrirTamplates.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //string NomePlan = RetornaNomePlanilhaSelecionado();
                //if (string.IsNullOrEmpty(NomePlan)) return;

                try
                {
                    ImportarArquivos csv = new ImportarArquivos();
                    using (DataTable dt = csv.ImportarSCV(AbrirTamplates.FileName))
                    {
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            CarregaGridView(dt);
                            return;
                        }
                        else
                        {
                            Mensagens.Alerta("Não foi possível carregar nenhum registro apartir do .csv informado. Por favor selecione outro arquivo.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Mensagens.Erro(string.Format("Não foi possível carregar o arquivo: {0}", ex.Message));
                }
            }
        }

        private void MenuArquivoExcelXML_Click(object sender, EventArgs e)
        {
            AbrirTamplates.Title = "Buscar Arquivo XML";
            AbrirTamplates.InitialDirectory = DirArquivo;
            AbrirTamplates.FileName = string.Empty;
            AbrirTamplates.DefaultExt = ".xml";
            AbrirTamplates.Filter = "Arquivos XML|*.xml|*.XML|";
            AbrirTamplates.RestoreDirectory = true;

            if (AbrirTamplates.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    using (DataSet Ds = new DataSet())
                    {
                        Ds.ReadXml(AbrirTamplates.FileName);
                        if (Ds != null && Ds.Tables[0].Rows.Count > 0)
                        {
                            CarregaGridView(Ds.Tables[0]);
                            return;
                        }
                        else
                        {
                            Mensagens.Alerta("Não foi possível carregar nenhum registro apartir do XML informado. Por favor selecione outro arquivo.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Mensagens.Erro(string.Format("Não foi possível carregar o arquivo XML. {0}", ex.Message));
                }
            }
        }

        private string RetornaNomePlanilhaSelecionadoXLS(string nomeArquivoBuscado)
        {
            List<DataTable> ListaDt = new List<DataTable>();
            int qtdLinhasDesejadas = 10;
            List<string> ListaNomePlan = new ImportarArquivos().ListSheetInExcel(String.Format(@"{0}", nomeArquivoBuscado));
            List<string> novaListaPlan = new List<string>();
            foreach (string item in ListaNomePlan)
            {
                string lllll = item.Replace("$_", "$");
                if (novaListaPlan.AsEnumerable().Any(m => m.Contains(lllll)) == false)
                {
                    novaListaPlan.Add(lllll);
                }
            }
            if (novaListaPlan.Count == 0)
            {
                return "";
            }
            if (novaListaPlan.Count == 1)
            {
                return novaListaPlan[0];
            }
            foreach (string itemNomePlan in novaListaPlan)
            {
                using (DataTable dt = new ImportarArquivos().ImportarXLSXNovo(nomeArquivoBuscado, itemNomePlan, qtdLinhasDesejadas))
                {
                    if (dt != null && dt.Rows.Count == 0)
                    {
                        DataTable data = new DataTable();
                        data.Columns.Add("  -");
                        data.Columns.Add("A");
                        data.Columns.Add("B");
                        data.Columns.Add("C");
                        data.Columns.Add("D");
                        data.Columns.Add("E");
                        data.Columns.Add("F");
                        data.Columns.Add("G");
                        data.Columns.Add("H");
                        data.Columns.Add("I");
                        data.Columns.Add("J");
                        data.Columns.Add("K");
                        data.Columns.Add("L");
                        data.Columns.Add("M");
                        data.Columns.Add("N");
                        data.Columns.Add("O");
                        data.Columns.Add("P");
                        data.Columns.Add("Q");

                        for (int i = 1; i <= qtdLinhasDesejadas; i++)
                        {
                            DataRow row = data.NewRow();
                            row["  -"] = i;
                            row["A"] = null;
                            row["B"] = "";
                            row["C"] = "";
                            row["D"] = "";
                            row["E"] = "";
                            row["F"] = "";
                            row["G"] = "";
                            row["H"] = "";
                            row["I"] = "";
                            row["J"] = "";
                            row["K"] = "";
                            row["L"] = "";
                            row["M"] = "";
                            row["N"] = "";
                            row["O"] = "";
                            row["P"] = "";
                            row["Q"] = "";
                            data.Rows.Add(row);
                        }
                        ListaDt.Add(data);
                    }
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        ListaDt.Add(dt);
                    }
                }
            }

            using (FormPlan plan = new FormPlan(ListaDt, novaListaPlan, nomeArquivoBuscado))
            {
                plan.ShowDialog(this);

                if (plan.cancelado == true)
                    return "";
                else
                    return plan.retorno;
            }
        }

        //private string RetornaNomePlanilhaSelecionadoXLSX(string nomeArquivoBuscado)
        //{
        //    List<DataTable> ListaDt = new List<DataTable>();
        //    int qtdLinhasDesejadas = 10;
        //    List<string> ListaNomePlan = new ImportarArquivos().ListSheetInExcel(String.Format(@"{0}", nomeArquivoBuscado));
        //    foreach (string itemNomePlan in ListaNomePlan)
        //    {
        //        using (DataTable dt = new ImportarArquivos().ImportarXLSXNovo(nomeArquivoBuscado, itemNomePlan, qtdLinhasDesejadas))
        //        {
        //            if (dt != null && dt.Rows.Count == 0)
        //            {
        //                DataTable data = new DataTable();
        //                data.Columns.Add("  -");
        //                data.Columns.Add("A");
        //                data.Columns.Add("B");
        //                data.Columns.Add("C");
        //                data.Columns.Add("D");
        //                data.Columns.Add("E");
        //                data.Columns.Add("F");
        //                data.Columns.Add("G");
        //                data.Columns.Add("H");
        //                data.Columns.Add("I");
        //                data.Columns.Add("J");
        //                data.Columns.Add("K");
        //                data.Columns.Add("L");
        //                data.Columns.Add("M");
        //                data.Columns.Add("N");
        //                data.Columns.Add("O");
        //                data.Columns.Add("P");
        //                data.Columns.Add("Q");

        //                for (int i = 1; i <= qtdLinhasDesejadas; i++)
        //                {
        //                    DataRow row = data.NewRow();
        //                    row["  -"] = i;
        //                    row["A"] = null;
        //                    row["B"] = "";
        //                    row["C"] = "";
        //                    row["D"] = "";
        //                    row["E"] = "";
        //                    row["F"] = "";
        //                    row["G"] = "";
        //                    row["H"] = "";
        //                    row["I"] = "";
        //                    row["J"] = "";
        //                    row["K"] = "";
        //                    row["L"] = "";
        //                    row["M"] = "";
        //                    row["N"] = "";
        //                    row["O"] = "";
        //                    row["P"] = "";
        //                    row["Q"] = "";
        //                    data.Rows.Add(row);
        //                }
        //                ListaDt.Add(data);
        //            }
        //            if (dt != null && dt.Rows.Count > 0)
        //            {
        //                ListaDt.Add(dt);
        //            }
        //        }
        //    }

        //    using (FormPlan plan = new FormPlan(ListaDt, ListaNomePlan, nomeArquivoBuscado))
        //    {
        //        plan.ShowDialog(this);

        //        if (plan.cancelado == true)
        //            return "";
        //        else
        //            return plan.retorno;
        //    }
        //}


        private void CarregaGridView(DataTable dt)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView1.ColumnCount = dt.Columns.Count;
            dataGridView1.ColumnHeadersVisible = true;

            // Set the column header style.
            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.DarkGray;
            columnHeaderStyle.ForeColor = Color.Black;
            columnHeaderStyle.Font = new Font("Verdana", 12, FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            int i = 0;
            foreach (var item in dt.Columns)
            {
                dataGridView1.Columns[i].Name = dt.Columns[i].ColumnName;
                dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                i++;
            }

            CarregaItensCombo();
            SetaItensDataSourceComboBox(dt);
            //SetaItensComboBox(dt);

            i = 0;
            foreach (var L in dt.Rows)
            {
                int j = 0;
                int qtdColunas = dt.Columns.Count;
                string[] itemValor = new string[qtdColunas];
                foreach (var C in dt.Columns)
                {
                    string valor = dt.Rows[i][j].ToString();
                    itemValor[j] = valor;
                    j++;
                }
                this.dataGridView1.Rows.Insert(i + 1, itemValor);
                dataGridView1.Rows[i + 1].ReadOnly = true;
                i++;

                //if (i > 0 && i.VerificaSePar())// muda a cor somente em linhas pares...
                //    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
            }

            if (dataGridView1.Rows.Count >= 1)
                dataGridView1.Rows[1].Frozen = true;
        }

        private string[] list;
        private void SetaItensDataSourceComboBox(DataTable Dt)
        {
            DataGridViewRow LinhaCombos = new DataGridViewRow();

            for (int i = 0; i < Dt.Columns.Count; i++)
            {
                DataGridViewComboBoxCell Cellcombo = new DataGridViewComboBoxCell();
                Cellcombo.DataSource = ColunaComboBoxCamposSelecaoGoogle;
                Cellcombo.Value = ColunaComboBoxCamposSelecaoGoogle.Rows[0][0]; // default value for the ComboBox
                Cellcombo.ValueMember = "ID";
                Cellcombo.DisplayMember = "CamposGooglePortugues";
                LinhaCombos.Cells.Add(Cellcombo);
            }
            dataGridView1.Rows.Add(LinhaCombos);
            //dataGridView1.Rows[1].Frozen = true;
        }

        private void SetaItensComboBox(DataTable Dt)
        {
            DataGridViewRow LinhaCombos = new DataGridViewRow();

            for (int i = 0; i < Dt.Columns.Count; i++)
            {
                DataGridViewComboBoxCell Cellcombo = new DataGridViewComboBoxCell();
                int indice = 0;
                foreach (string item in CamposGooglePortugues)
                {
                    Cellcombo.Items.Add(new { Text = item, Value = indice });
                    indice++;
                }

                //Cellcombo.Items = ColunaComboBoxCamposSelecaoGoogle[0];
                Cellcombo.Value = 0; // default value for the ComboBox
                Cellcombo.DisplayMember = "Text";
                Cellcombo.ValueMember = "Value";
                LinhaCombos.Cells.Add(Cellcombo);
            }

            dataGridView1.Rows.Add(LinhaCombos);
            //dataGridView1.Rows[1].Frozen = true;
        }

        private void CarregaItensCombo()
        {
            ColunaComboBoxCamposSelecaoGoogle = new DataTable();
            ColunaComboBoxCamposSelecaoGoogle.Columns.Add("ID");
            ColunaComboBoxCamposSelecaoGoogle.Columns.Add("CamposGooglePortugues");
            //ColunaComboBoxCamposSelecaoGoogle.Columns.Add("CamposGoogleIngles");
            for (int i = 0; i < CamposGooglePortugues.Count; i++)
            {
                DataRow linha = ColunaComboBoxCamposSelecaoGoogle.NewRow();
                linha["ID"] = i;
                linha["CamposGooglePortugues"] = CamposGooglePortugues[i];
                //linha["CamposGoogleIngles"] = CamposGooglePortugues[i];
                ColunaComboBoxCamposSelecaoGoogle.Rows.Add(linha);
            }
        }

        private StringBuilder sbFinalTXT = new StringBuilder();
        StringBuilder sbContato = new StringBuilder();
        Dictionary<int, string> dicionarioSequenciaGrid = new Dictionary<int, string>();
        public static int valorIndiceLinhaGrid = 0;
        void Processando()
        {
            for (int IndiceLinhaGrid = 0; IndiceLinhaGrid < dataGridView1.Rows.Count; IndiceLinhaGrid++) //percorrendo linhas....
            {
                if (IndiceLinhaGrid == 0) continue;
                gruposLinhaAtual.Clear(); // = new StringBuilder();
                grupoLinhaJaProcessado = false;
                int qtdEncontrado = 0;

                sbContato.AppendLine("BEGIN:VCARD");//linha 1
                sbContato.AppendLine("VERSION:3.0");//linha 2 

                #region Nome completo
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Nome completo");
                if (qtdEncontrado == 1)
                {
                    string pegaCampoNomeCompleto = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Nome completo");

                    //verifica se contém coluna selecionada para tratamento
                    if (dicionarioSequenciaGrid.AsEnumerable().Any(pair => pair.Value == "Tratamento"))
                    {
                        string pegaCampoTratamento = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Tratamento");

                        string linha = string.Format("N:;{0};;{1};", pegaCampoNomeCompleto, pegaCampoTratamento);
                        sbContato.AppendLine(linha);
                        linha = string.Format("FN:{0} {1}", pegaCampoTratamento, pegaCampoNomeCompleto);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                    else
                    {
                        string linha = string.Format("N:;{0};;;", pegaCampoNomeCompleto);
                        sbContato.AppendLine(linha);
                        linha = string.Format("FN:{0}", pegaCampoNomeCompleto);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Apelido
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Apelido");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Apelido");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("NICKNAME:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Grupos
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Grupos");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Grupos");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("CATEGORIES:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                if (qtdEncontrado > 1)
                {
                    if (grupoLinhaJaProcessado == false)
                    {
                        gruposLinhaAtual.Clear();
                        string valorCelulaAtual = "";
                        var itensIguaisGrid = dicionarioSequenciaGrid.AsEnumerable().Where(pair => pair.Value == "Grupos");
                        foreach (KeyValuePair<int, string> item in itensIguaisGrid)
                        {
                            var key = item.Key;
                            var value = item.Value;
                            valorCelulaAtual = dataGridView1.Rows[IndiceLinhaGrid].Cells[key].FormattedValue.ToString();
                            if (string.IsNullOrEmpty(valorCelulaAtual)) continue;
                            if (gruposLinhaAtual.ToString() == "")
                                gruposLinhaAtual.AppendFormat(string.Format("{0}", valorCelulaAtual));
                            else
                            {
                                string novoValor = string.Format("{0},{1}", gruposLinhaAtual, valorCelulaAtual);
                                gruposLinhaAtual.Clear();
                                gruposLinhaAtual.AppendFormat(novoValor);
                            }
                        }
                        valorCelulaAtual = string.Format("{0},myContacts", gruposLinhaAtual.ToString());
                        if (!string.IsNullOrEmpty(valorCelulaAtual))
                        {
                            string linha = string.Format("CATEGORIES:{0}", valorCelulaAtual);
                            sbContato.AppendLine(linha);
                            string temp = sbContato.ToString();
                        }
                        grupoLinhaJaProcessado = true;
                    }
                }
                #endregion
                #region Data nascimento
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Data nascimento");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Data nascimento");
                    DateTime objDate;
                    if (DateTime.TryParse(pegaCampo, out objDate))
                    {
                        //Se é uma data válida
                        string linha = string.Format("BDAY:{0}", pegaCampo.ToDateTime().Date.ToString("dd/MM/yyyy"));
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Sexo
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Sexo");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Sexo");
                    if (pegaCampo.ToUpper() == "M" || pegaCampo.ToUpper() == "MASCULINO")
                    {
                        string linha = string.Format("GENDER:{0}", "M");
                        sbContato.AppendLine(linha);
                    }
                    if (pegaCampo.ToUpper() == "F" || pegaCampo.ToUpper() == "FEMININO")
                    {
                        string linha = string.Format("GENDER:{0}", "F");
                        sbContato.AppendLine(linha);
                    }

                    string temp = sbContato.ToString();
                }
                #endregion
                #region Empresa
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Empresa");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Empresa");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("ORG:{0};", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Cargo/Funcão
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Cargo/Funcão");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Cargo/Funcão");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("TITLE:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Nota/Histórico
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Nota/Histórico");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Nota/Histórico");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("NOTE:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Email Pessoal
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Email Pessoal");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Email Pessoal");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("EMAIL;TYPE=INTERNET;TYPE=Pessoal;type=pref:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Email Comercial
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Email Comercial");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Email Comercial");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("EMAIL;TYPE=INTERNET;TYPE=WORK:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Telefone Principal
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Telefone Principal");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Telefone Principal");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("TEL;type=Telefone1;type=pref:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Telefone 2
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Telefone 2");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Telefone 2");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("TEL;type=Telefone2:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Telefone 3
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Telefone 3");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Telefone 3");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("TEL;type=Telefone3:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Telefone 4
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Telefone 4");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Telefone 4");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("TEL;type=Telefone4:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Telefone 5
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Telefone 5");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Telefone 5");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("TEL;type=Telefone5:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                #region Telefone 6
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Telefone 6");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Telefone 6");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("TEL;type=Telefone6:{0}", pegaCampo);
                        sbContato.AppendLine(linha);
                        string temp = sbContato.ToString();
                    }
                }
                #endregion
                int i = 1;
                #region Trata Endereço Res.
                //qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Endereço Res.");
                //if (qtdEncontrado == 1)
                //{
                //    string pegaCampoEnderecoRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Endereço Res.");
                //    //string pegaCampoNumeroRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Endereço Res.");
                //    string pegaCampoBairroRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Bairro Res.");
                //    string pegaCampoCidadeRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Cidade Res.");
                //    string pegaCampoEstadoRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Estado Res.").TrocaEstadoPorSegla();
                //    string pegaCampoCEPRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "CEP Res.");
                //    string pegaCampoPaisRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "País Res.");
                //    if (pegaCampoPaisRes.ToUpper() == "Brasil".ToUpper()) pegaCampoPaisRes = "BRAZIL";

                //    string linha = string.Format("item{6}.ADR;type=HOME;type=pref:;;{0} {1};{2};{3};{4};{5}",
                //        pegaCampoEnderecoRes == "" ? "-" : pegaCampoEnderecoRes,
                //        pegaCampoBairroRes == "" ? "-" : pegaCampoBairroRes,
                //        pegaCampoCidadeRes == "" ? "-" : pegaCampoCidadeRes,
                //        pegaCampoEstadoRes == "" ? "-" : pegaCampoEstadoRes,
                //        pegaCampoCEPRes == "" ? "-" : pegaCampoCEPRes,
                //        pegaCampoPaisRes == "" ? "-" : pegaCampoPaisRes,
                //        i);

                //    sbContato.AppendLine(linha);
                //    sbContato.AppendLine(string.Format("item{0}.X-ABADR:BR", i));
                //    string temp = sbContato.ToString();
                //    i++;
                //}
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Endereço Res.");
                if (qtdEncontrado == 1)
                {
                    string pegaCampoEnderecoRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Endereço Res.");
                    string pegaCampoNumeroRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Endereço Res. Núm.");
                    string pegaCampoBairroRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Bairro Res.");
                    string pegaCampoCidadeRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Cidade Res.");
                    //string pegaCampoEstadoRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Estado Res.").TrocaEstadoPorSegla();
                    string pegaCampoEstadoRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Estado Res.");
                    string pegaCampoCEPRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "CEP Res.");
                    string pegaCampoPaisRes = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "País Res.");
                    if (pegaCampoPaisRes.ToUpper() == "".ToUpper() || pegaCampoPaisRes.ToUpper() == "Brasil".ToUpper()) pegaCampoPaisRes = "BRAZIL";

                    //item2.ADR:NUMERO0;BAIRRO1;RUA2;CIDADE3;Tocantins4;CEP5;Brazil6
                    //item2.X - ABLabel:
                    string linha = string.Format("item{7}.ADR:{0};{1};{2};{3};{4};{5};{6}",
                        pegaCampoNumeroRes == "" ? "" : pegaCampoNumeroRes,
                        pegaCampoBairroRes == "" ? "" : pegaCampoBairroRes,
                        pegaCampoEnderecoRes == "" ? "" : pegaCampoEnderecoRes,
                        pegaCampoCidadeRes == "" ? "" : pegaCampoCidadeRes,
                        pegaCampoEstadoRes == "" ? "" : pegaCampoEstadoRes,
                        pegaCampoCEPRes == "" ? "" : pegaCampoCEPRes,
                        pegaCampoPaisRes == "" ? "" : pegaCampoPaisRes,
                        i);

                    sbContato.AppendLine(linha);
                    sbContato.AppendLine(string.Format("item{0}.X-ABLabel", i));
                    string temp = sbContato.ToString();
                    i++;

                }
                #endregion
                #region Trata Endereço Com.
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Endereço Res.");
                if (qtdEncontrado == 1)
                {
                    string pegaCampoEnderecoCom = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Endereço Com.");
                    string pegaCampoNumeroCom = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Endereço Com. Núm.");
                    string pegaCampoBairroCom = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Bairro Com.");
                    string pegaCampoCidadeCom = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Cidade Com.");
                    string pegaCampoEstadoCom = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Estado Com.").TrocaEstadoPorSegla();
                    string pegaCampoCEPCom = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "CEP Com.");
                    string pegaCampoPaisCom = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "País Com.");
                    if (pegaCampoPaisCom.ToUpper() == "Brasil".ToUpper()) pegaCampoPaisCom = "BRAZIL";

                    string linha = string.Format("item{7}.ADR;type=WORK:{0};{1};{2};{3};{4};{5};{6}",
                        pegaCampoNumeroCom == "" ? "" : pegaCampoNumeroCom, 
                        pegaCampoBairroCom == "" ? "" : pegaCampoBairroCom,
                        pegaCampoEnderecoCom == "" ? "" : pegaCampoEnderecoCom,
                        pegaCampoCidadeCom == "" ? "" : pegaCampoCidadeCom,
                        pegaCampoEstadoCom == "" ? "" : pegaCampoEstadoCom,
                        pegaCampoCEPCom == "" ? "" : pegaCampoCEPCom,
                        pegaCampoPaisCom == "" ? "" : pegaCampoPaisCom, 
                        i);
                    sbContato.AppendLine(linha);
                    sbContato.AppendLine(string.Format("item{0}.X-ABLabel", i));
                    string temp = sbContato.ToString();
                    i++;
                }
                #endregion
                #region Site Pessoal
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Site Pessoal");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Site Pessoal");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("item{0}.URL;type=pref:{1}", i, pegaCampo);
                        sbContato.AppendLine(linha);
                        sbContato.AppendLine(string.Format("item{0}.X-ABLabel:_$!<HomePage>!$_", i));
                        string temp = sbContato.ToString();
                        i++;
                    }
                }
                #endregion
                #region Site Comercial
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Site Comercial");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Site Comercial");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("item{0}.URL:{1}", i, pegaCampo);
                        sbContato.AppendLine(linha);
                        sbContato.AppendLine(string.Format("item{0}.X-ABLabel:Comercial", i));
                        string temp = sbContato.ToString();
                        i++;
                    }
                }
                #endregion
                #region Data aniv. Cônjuge
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Data aniv. Cônjuge");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Data aniv. Cônjuge");
                    DateTime objDate;
                    if (DateTime.TryParse(pegaCampo, out objDate))
                    {
                        //Se é uma data válida
                        string linha = string.Format("item{1}.X-ABDATE:{0}", pegaCampo.ToDateTime().Date.ToString("yyyyMMdd"), i);
                        sbContato.AppendLine(linha);
                        sbContato.AppendLine(string.Format("item{0}.X-ABLabel:Aniversário da Esposa", i));
                        string temp = sbContato.ToString();
                        i++;
                    }
                }
                #endregion
                #region Nome Cônjuge
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Nome Cônjuge");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Nome Cônjuge");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("item{1}.X-ABRELATEDNAMES;type=pref:{0}", pegaCampo, i);
                        sbContato.AppendLine(linha);
                        sbContato.AppendLine(string.Format("item{0}.X-ABLabel:_$!<Spouse>!$_", i));
                        string temp = sbContato.ToString();
                        i++;
                    }
                }
                #endregion
                #region Nome do pai
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Nome do pai");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Nome do pai");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("item{1}.X-ABRELATEDNAMES:{0}", pegaCampo, i);
                        sbContato.AppendLine(linha);
                        sbContato.AppendLine(string.Format("item{0}.X-ABLabel:_$!<Father>!$_", i));
                        string temp = sbContato.ToString();
                        i++;
                    }
                }
                #endregion
                #region Nome da mãe
                qtdEncontrado = RetornaQtdColunasParaOCampoIndicado(dicionarioSequenciaGrid, "Nome da mãe");
                if (qtdEncontrado == 1)
                {
                    string pegaCampo = RetornaValorCelulaPelaColuna(IndiceLinhaGrid, dicionarioSequenciaGrid, "Nome da mãe");
                    if (!string.IsNullOrEmpty(pegaCampo))
                    {
                        string linha = string.Format("item{1}.X-ABRELATEDNAMES:{0}", pegaCampo, i);
                        sbContato.AppendLine(linha);
                        sbContato.AppendLine(string.Format("item{0}.X-ABLabel:_$!<Mother>!$_", i));
                        string temp = sbContato.ToString();
                        i++;
                    }
                }
                #endregion

                sbContato.AppendLine("END:VCARD");
            }
            string tempo = sbContato.ToString();


            sbFinalTXT.AppendLine(sbContato.ToString());
            string final = sbFinalTXT.ToString();//teste          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            valorIndiceLinhaGrid = 0;
            sbContato = new StringBuilder();
            sbFinalTXT = new StringBuilder();

            //int: indice da coluna Grid / string: Nome Combo Selecionado
            dicionarioSequenciaGrid = new Dictionary<int, string>();
            for (int indice = 0; indice < dataGridView1.Columns.Count; indice++)
            {
                string valorCombo = dataGridView1.Rows[0].Cells[indice].FormattedValue.ToString();
                if (dataGridView1.Rows[0].Cells[indice].Value.ToString() == "0") continue;
                dicionarioSequenciaGrid.Add(indice, valorCombo);
            }
            if (dataGridView1.Rows.Count == 0)
            {
                Mensagens.Informa("Nenhuma linha na seleção atual.\nBusque um fonte de dados e tente novamente.");
                return;
            }
            if (dicionarioSequenciaGrid.Count == 0)
            {
                Mensagens.Informa("Nenhuma coluna foi modificada.");
                return;
            }

            using (FormWaiting frm = new FormWaiting(Processando))
            {
                frm.ShowDialog(this);
            }


            //-----------------------------------------------------
            //define o titulo
            saveFileDialog1.Title = "Salvar Arquivo Texto";
            //Define as extensões permitidas
            //saveFileDialog1.Filter = "Arquivo vCard|.vcf";
            saveFileDialog1.Filter = "Arquivo vCard (*.vcf)|*.vcf";
            //define o indice do filtro
            saveFileDialog1.FilterIndex = 0;
            //Atribui um valor vazio ao nome do arquivo
            saveFileDialog1.FileName = "ContatosExportados_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            //Define a extensão padrão como .txt
            saveFileDialog1.DefaultExt = ".vcf";
            //define o diretório padrão
            //saveFileDialog1.InitialDirectory = @"c:\dados";
            //restaura o diretorio atual antes de fechar a janela
            saveFileDialog1.RestoreDirectory = true;

            //Abre a caixa de dialogo e determina qual botão foi pressionado
            DialogResult resultado = saveFileDialog1.ShowDialog();

            //Se o ousuário pressionar o botão Salvar
            if (resultado == DialogResult.OK)
            {
                //cria um stream usando o nome do arquivo
                System.IO.FileStream fs = new System.IO.FileStream(saveFileDialog1.FileName, System.IO.FileMode.Create);

                //cria um escrito que irá escrever no stream
                System.IO.StreamWriter writer = new System.IO.StreamWriter(fs);
                //escreve o conteúdo da caixa de texto no stream
                writer.Write(sbFinalTXT.ToString());
                //fecha o escrito e o stream
                writer.Close();


                //string caminhoArquivo = @"C:\Projetos\KEMUEL\ContatosExportados.vcf"; //caminho completo
                //System.IO.File.WriteAllText(saveFileDialog1.FileName, sbFinalTXT.ToString());

                Mensagens.Informa("Exportação concluída com sucesso!");
            }
            else
            {
                //exibe mensagem informando que a operação foi cancelada
                //MessageBox.Show("Operação cancelada");
                return;
            }
        }

        private static int RetornaQtdColunasParaOCampoIndicado(Dictionary<int, string> dicionarioSequenciaGrid, string nomeCampo)
        {
            int qtdEncontrado = dicionarioSequenciaGrid.AsEnumerable().Count(pair => pair.Value == nomeCampo).ToInt();
            return qtdEncontrado;
        }

        private string RetornaValorCelulaPelaColuna(int IndiceLinhaGrid, Dictionary<int, string> dicionarioSequenciaGrid, string CampoComboBox)
        {
            string pegaCampo = "";
            if (dicionarioSequenciaGrid.AsEnumerable().Any(pair => pair.Value == CampoComboBox))
            {
                var indiceCampo = dicionarioSequenciaGrid.AsEnumerable().First(pair => pair.Value == CampoComboBox);//pega qual coluna tem o campo
                pegaCampo = dataGridView1.Rows[IndiceLinhaGrid].Cells[indiceCampo.Key].FormattedValue.ToString();
            }
            return pegaCampo;
        }

        private string TrataCondicoesEspecificas(string valorCelulaAtual, string value)
        {
            string retorno = "";

            switch (value)
            {
                case "Nome completo":
                    retorno = string.Format("N:;{0};;[Tratamento];", valorCelulaAtual);
                    break;
                case "Tratamento":

                    break;
                case "Apelido":

                    break;
                case "Data nascimento":

                    break;
                case "Sexo":

                    break;
                case "Telefone Principal":

                    break;
                case "Telefone 2":

                    break;
                case "Telefone 3":

                    break;
                case "Telefone 4":

                    break;
                case "Telefone 5":

                    break;
                case "telefone 6":

                    break;
                case "Email Pessoal":

                    break;
                case "Email Comercial":

                    break;
                case "Endereço Res.":

                    break;
                case "Bairro Res.":

                    break;
                case "Cidade Res.":

                    break;
                case "Estado Res.":

                    break;
                case "CEP Res.":

                    break;
                case "País Res.":

                    break;
                case "Endereço Com.":

                    break;
                case "Bairro Com.":

                    break;
                case "Cidade Com.":

                    break;
                case "Estado Com.":

                    break;
                case "CEP Com.":

                    break;
                case "País Com.":

                    break;
                case "Empresa":

                    break;
                case "Cargo/Funcão":

                    break;
                case "Nota/Histórico":

                    break;
                case "Site Pessoal":

                    break;
                case "Site Comercial":

                    break;
                case "Nome Cônjuge":

                    break;
                case "Data aniv. Cônjuge":

                    break;
                case "Nome do pai":

                    break;
                case "Nome da mãe":

                    break;
                case "Grupos":

                    break;

                default:
                    retorno = valorCelulaAtual;
                    break;
            }

            return retorno;
        }

    }
}

