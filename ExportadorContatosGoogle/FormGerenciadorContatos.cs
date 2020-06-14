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
    public partial class FormGerenciadorContatos : Form
    {
        private string DirArquivo { get; set; }
        private DataTable TblListaAtual { get; set; }

        private List<string> CamposGoogleIngles = new List<string> { "Selecione...", "Given Name", "Name Prefix", "Nickname", "Birthday", "Gender", "Occupation", "Notes", "Group Membership", "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2- Value", "Phone 3 - Type", "Phone 3- Value", "Phone 4 - Type", "Phone 4- Value", "Phone 5 - Type", "Phone 5- Value", "Phone 6 - Type", "Phone 6- Value", "Address 1 - Type", "Address 1 - Formatted", "Address 1 - Street", "Address 1 - City", "Address 1 - PO Box", "Address 1 - Region", "Address 1 - Postal Code", "Address 1 - Country", "Address 1 - Extended Address", "Relation 1 - Type", "Relation 1 - Value", "Event 1 - Type", "Event 1 - Value" };
        private List<string> CamposGooglePortugues = new List<string> { "Selecione...", "Nome completo", "Tratamento", "Nome", "Data de aniversário", "Sexo", "Ocupação", "Notas", "Grupo", "Contato1", "Numero do telefone1", "Contato2", "Numero do telefone2", "Contato3", "Numero do telefone3", "Contato4", "Numero do telefone4", "Contato5", "Numero do telefone5", "Contato6", "Numero do telefone6", "Tipo endereço", "Endereço completo", "Logradouro", "Cidade", "Número", "Estado", "CEP", "País", "Bairro", "Tipo de relacionamento", "Relacionamento", "Data evento", "Descrição do evento" };
        private DataTable ColunaComboBoxCamposSelecaoGoogle;
        private string NomeUltimoCampo = "Event 1 - Value";
        private StringBuilder gruposLinhaAtual = new StringBuilder();

        public FormGerenciadorContatos()
        {
            InitializeComponent();

            dataGridView1.CellValueChanged += new DataGridViewCellEventHandler(dataGridView1_CellValueChanged);

            dataGridView1.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(dataGridView1_EditingControlShowing);
            dataGridView1.SelectionChanged += new EventHandler(dataGridView1_SelectionChanged);
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void FormGerenciadorContatos_Load(object sender, EventArgs e)
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
                        if (CampoAtualSelecionado == "Grupo") continue;
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
                string NomePlan = RetornaNomePlanilhaSelecionado();
                if (string.IsNullOrEmpty(NomePlan)) return;

                try
                {
                    ImportarArquivos xlsx = new ImportarArquivos();
                    //using (DataTable dt = xlsx.ImportarXLSX("C:\\Users\\MARQUES\\Desktop\\Pasta1.xlsx", "plan1"))
                    using (DataTable dt = xlsx.ImportarXLSX(AbrirTamplates.FileName, NomePlan))
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
                string NomePlan = RetornaNomePlanilhaSelecionado();
                if (string.IsNullOrEmpty(NomePlan)) return;

                try
                {
                    ImportarArquivos xls = new ImportarArquivos();
                    using (DataTable dt = xls.ImportarXLS(AbrirTamplates.FileName, NomePlan))
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
                string NomePlan = RetornaNomePlanilhaSelecionado();
                if (string.IsNullOrEmpty(NomePlan)) return;

                try
                {
                    ImportarArquivos csv = new ImportarArquivos();
                    using (DataTable dt = csv.ImportarSCV(AbrirTamplates.FileName, NomePlan))
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

        private string RetornaNomePlanilhaSelecionado()
        {
            using (FormPlan plan = new FormPlan())
            {
                plan.ShowDialog(this);

                if (plan.cancelado == true)
                    return "";
                else
                    return "";//plan.TxtPlanilha.Text;
            }
        }

        private void CarregaGridView(DataTable dt)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView1.ColumnCount = dt.Columns.Count;
            dataGridView1.ColumnHeadersVisible = true;

            // Set the column header style.
            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.DarkGray;
            columnHeaderStyle.ForeColor = Color.White;
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

                if (i > 0 && i.VerificaSePar())// muda a cor somente em linhas pares...
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
            }

            if (dataGridView1.Rows.Count >= 1)
                dataGridView1.Rows[1].Frozen = true;
        }



        private void AddComboBoxColumn(List<string> NomesColunasExcel)
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DataGridViewComboBoxColumn ColunaComboGrid = new DataGridViewComboBoxColumn()
                        {
                            DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox,
                            ReadOnly = false,
                            HeaderText = "Coluna Google",
                            Name = "ColunaGoogle",
                            DefaultCellStyle = dataGridViewCellStyle1,
                            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                            MinimumWidth = 400
                        };



            //System.Collections.ArrayList retornoArquivoExportadorPeloGoogle = Arquivos.LerArquivo(@"C:\Users\MARQUES\Desktop\terci.csv");
            //string[] itens = retornoArquivoExportadorPeloGoogle[0].ToString().Split(',');//pega a 1º linha
            //CamposGoogleIngles = new List<string>(itens.ToList());
            //CamposGooglePortuguesSort = new List<string>(CamposGooglePortugues);
            //CamposGooglePortuguesSort.Sort();

            ColunaComboBoxCamposSelecaoGoogle = new DataTable();
            //ColunaComboBoxCamposSelecaoGoogle.Columns.Add("ID");
            //ColunaComboBoxCamposSelecaoGoogle.Columns.Add("NomeIngles");
            ColunaComboBoxCamposSelecaoGoogle.Columns.Add("NomePortugues");
            for (int i = 0; i < CamposGoogleIngles.Count; i++)
            {
                DataRow linha = ColunaComboBoxCamposSelecaoGoogle.NewRow();
                linha["ID"] = i;
                //linha["NomeIngles"] = CamposGoogleIngles[i];
                linha["NomePortugues"] = CamposGooglePortugues[i];
                ColunaComboBoxCamposSelecaoGoogle.Rows.Add(linha);
            }

            ColunaComboGrid.ValueMember = "ID";
            ColunaComboGrid.DisplayMember = "NomePortugues";
            ColunaComboGrid.DataSource = ColunaComboBoxCamposSelecaoGoogle;
            for (int i = 0; i < NomesColunasExcel.Count; i++)
            {
                if (i == 10) break;
                //dataGridView1.Columns.Add();
            }

        }

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
            ColunaComboBoxCamposSelecaoGoogle.Columns.Add("CamposGoogleIngles");
            for (int i = 0; i < CamposGoogleIngles.Count; i++)
            {
                DataRow linha = ColunaComboBoxCamposSelecaoGoogle.NewRow();
                linha["ID"] = i;
                linha["CamposGooglePortugues"] = CamposGooglePortugues[i];
                linha["CamposGoogleIngles"] = CamposGoogleIngles[i];
                ColunaComboBoxCamposSelecaoGoogle.Rows.Add(linha);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sblinhaTXT = new StringBuilder();
            StringBuilder sbFinalTXT = new StringBuilder();
            for (int i = 0; i < CamposGoogleIngles.Count; i++)
            {
                if (i == 0) continue;
                string valor = CamposGoogleIngles[i].ToString();
                //verifica se for o ultimo não colocar ","
                if (valor == NomeUltimoCampo) // --> "Event 1 - Value"
                    sblinhaTXT.AppendFormat(string.Format("{0}", valor));
                else
                    sblinhaTXT.AppendFormat(string.Format("{0},", valor));
            }
            sbFinalTXT.AppendLine(sblinhaTXT.ToString());
            //sblinhaTXT.AppendLine();            

            //int: indice da coluna Grid / string: Nome Combo Selecionado
            Dictionary<int, string> dicionarioSequenciaGrid = new Dictionary<int, string>();
            for (int indice = 0; indice < dataGridView1.Columns.Count; indice++)
            {
                string valorCombo = dataGridView1.Rows[0].Cells[indice].FormattedValue.ToString();
                dicionarioSequenciaGrid.Add(indice, valorCombo);
            }

            for (int IndiceLinhaGrid = 0; IndiceLinhaGrid < dataGridView1.Rows.Count; IndiceLinhaGrid++) //percorrendo linhas....
            {
                sblinhaTXT.Clear();
                gruposLinhaAtual.Clear(); // = new StringBuilder();
                if (IndiceLinhaGrid == 0) continue;
                for (int indiceGoogle = 0; indiceGoogle < CamposGooglePortugues.Count; indiceGoogle++)
                {
                    if (indiceGoogle == 0) continue;
                    string campoGoogle = CamposGooglePortugues[indiceGoogle].ToString();
                    int indiceCampoGoogle = indiceGoogle;

                    int qtdEncontrado = dicionarioSequenciaGrid.AsEnumerable().Count(pair => pair.Value == campoGoogle).ToInt();
                    #region Se quantidade == 0
                    if (qtdEncontrado == 0)
                    {
                        //trata contatos pois não está no excel retornado...
                        if (campoGoogle == "Contato1" || campoGoogle == "Contato2" || campoGoogle == "Contato3" || campoGoogle == "Contato4" || campoGoogle == "Contato5" || campoGoogle == "Contato6")
                        {
                            if (indiceGoogle == (CamposGooglePortugues.Count - 1))//ultimo registro, tira a ultima virgula
                                sblinhaTXT.AppendFormat(string.Format("{0}", "Outro"));
                            else
                                sblinhaTXT.AppendFormat(string.Format("{0},", "Outro"));
                            continue;
                        }

                        if (indiceGoogle == (CamposGooglePortugues.Count - 1))//ultimo registro, tira a ultima virgula
                            sblinhaTXT.AppendFormat(string.Format("{0}", ""));
                        else
                            sblinhaTXT.AppendFormat(string.Format("{0},", ""));

                        continue;
                    }
                    #endregion
                    #region Se quantidade == 1
                    if (qtdEncontrado == 1)
                    {
                        var key = dicionarioSequenciaGrid.AsEnumerable().FirstOrDefault(pair => pair.Value == campoGoogle).Key;
                        var value = dicionarioSequenciaGrid.AsEnumerable().FirstOrDefault(pair => pair.Value == campoGoogle).Value;
                        string valorCelulaAtual = dataGridView1.Rows[IndiceLinhaGrid].Cells[key].FormattedValue.ToString();
                        valorCelulaAtual = TrataCondicoesEspecificas(valorCelulaAtual, value);
                        if (indiceGoogle == (CamposGooglePortugues.Count - 1))//ultimo registro, tira a ultima virgula
                        {
                            sblinhaTXT.AppendFormat(string.Format("{0}", valorCelulaAtual));
                        }
                        else
                        {
                            sblinhaTXT.AppendFormat(string.Format("{0},", valorCelulaAtual));
                        }
                        continue;
                    }
                    #endregion
                    #region Se quantidade > 1
                    if (qtdEncontrado > 1)
                    {
                        string valorCelulaAtual = "";
                        var itensIguaisGrid = dicionarioSequenciaGrid.AsEnumerable().Where(pair => pair.Value == campoGoogle);
                        foreach (KeyValuePair<int, string> item in itensIguaisGrid)
                        {
                            var key = item.Key;
                            var value = item.Value;
                            valorCelulaAtual = dataGridView1.Rows[IndiceLinhaGrid].Cells[key].FormattedValue.ToString();
                            valorCelulaAtual = TrataCondicoesEspecificas(valorCelulaAtual, value);

                        }
                        if (campoGoogle == "Grupo") valorCelulaAtual = string.Format("{0};myContacts", valorCelulaAtual);
                        if (indiceGoogle == (CamposGooglePortugues.Count - 1))//ultimo registro, tira a ultima virgula
                        {
                            sblinhaTXT.AppendFormat(string.Format("{0}", valorCelulaAtual));
                        }
                        else
                        {
                            sblinhaTXT.AppendFormat(string.Format("{0},", valorCelulaAtual));
                        }
                        continue;
                    }
                    #endregion
                }
                sbFinalTXT.AppendLine(sblinhaTXT.ToString());
            }
            string final = sbFinalTXT.ToString();

            string caminhoArquivo = @"C:\Projetos\KEMUEL\ArquivoExportado.csv"; //caminho completo

            //StringBuilder sConteudo = new StringBuilder();
            //sConteudo.AppendLine("Primeira linha do arquivo.");
            //sConteudo.AppendLine("Segunda linha do arquivo.");
            //sConteudo.AppendLine("Terceira e última linha do arquivo.");
            //invocando o método WriteAllText, informando o caminho e o conteúdo
            System.IO.File.WriteAllText(caminhoArquivo, sbFinalTXT.ToString());

            Mensagens.Informa("Exportação concluída com sucesso!");
        }

        private string TrataCondicoesEspecificas(string valorCelulaAtual, string value)
        {
            string retorno = "";

            switch (value)
            {
                case "Nome completo":
                    retorno = valorCelulaAtual;
                    break;
                case "Tratamento":
                    retorno = valorCelulaAtual;
                    break;
                case "Nome":
                    retorno = valorCelulaAtual;
                    break;
                case "Data de aniversário":
                    try
                    {
                        retorno = String.Format("{0:dd/MM/yyyy}", valorCelulaAtual);
                        retorno = retorno.ToDateTime().ToShortDateString();
                    }
                    catch (Exception)
                    {
                        retorno = valorCelulaAtual;
                    }

                    break;
                case "Sexo":
                    if (valorCelulaAtual.ToUpper() == "F".ToUpper())
                    {
                        retorno = "feme";
                    }
                    else if (valorCelulaAtual.ToUpper() == "M".ToUpper())
                    {
                        retorno = "male";
                    }
                    else retorno = valorCelulaAtual;
                    break;
                case "Ocupação":
                    retorno = valorCelulaAtual;
                    break;
                case "Notas":
                    if (string.IsNullOrEmpty(valorCelulaAtual))
                        retorno = valorCelulaAtual;
                    else
                        retorno = string.Format("\"{0}\"", valorCelulaAtual);
                    break;
                case "Grupo":
                    if (string.IsNullOrEmpty(valorCelulaAtual))
                    {
                        retorno = gruposLinhaAtual.ToString();
                    }
                    else
                    {
                        if (gruposLinhaAtual.ToString() == "")
                            gruposLinhaAtual.AppendFormat(string.Format("{0}", valorCelulaAtual));
                        else
                        {
                            string novoValor = string.Format("{0};{1}", gruposLinhaAtual, valorCelulaAtual);
                            gruposLinhaAtual.Clear();
                            gruposLinhaAtual.AppendFormat(novoValor);
                        }
                        retorno = gruposLinhaAtual.ToString();
                    }
                    break;
                case "Contato1":
                    retorno = "Outro";
                    break;
                case "Numero do telefone1":
                    retorno = valorCelulaAtual;
                    break;
                case "Contato2":
                    retorno = "Outro";
                    break;
                case "Numero do telefone2":
                    retorno = valorCelulaAtual;
                    break;
                case "Contato3":
                    retorno = "Outro";
                    break;
                case "Numero do telefone3":
                    retorno = valorCelulaAtual;
                    break;
                case "Contato4":
                    retorno = "Outro";
                    break;
                case "Numero do telefone4":
                    retorno = valorCelulaAtual;
                    break;
                case "Contato5":
                    retorno = "Outro";
                    break;
                case "Numero do telefone5":
                    retorno = valorCelulaAtual;
                    break;
                case "Contato6":
                    retorno = "Outro";
                    break;
                case "Numero do telefone6":
                    retorno = valorCelulaAtual;
                    break;
                case "Tipo endereço":
                    retorno = valorCelulaAtual;
                    break;
                case "Endereço completo":
                    if (string.IsNullOrEmpty(valorCelulaAtual))
                        retorno = valorCelulaAtual;
                    else
                        retorno = string.Format("\"{0}\"", valorCelulaAtual);
                    retorno = retorno.RemoveSimbolos().RemovePontuacao().RemoveSpecialChars();
                    break;
                case "Logradouro":
                    if (string.IsNullOrEmpty(valorCelulaAtual))
                        retorno = valorCelulaAtual;
                    else
                        retorno = string.Format("\"{0}\"", valorCelulaAtual);
                    retorno = retorno.RemoveSimbolos().RemovePontuacao().RemoveSpecialChars();
                    break;
                case "Cidade":
                    retorno = valorCelulaAtual;
                    break;
                case "Número":
                    retorno = valorCelulaAtual;
                    break;
                case "Estado":
                    retorno = valorCelulaAtual;
                    break;
                case "CEP":
                    retorno = valorCelulaAtual;
                    break;
                case "País":
                    retorno = valorCelulaAtual;
                    break;
                case "Bairro":
                    retorno = valorCelulaAtual;
                    break;
                case "Tipo de relacionamento":
                    retorno = valorCelulaAtual;
                    break;
                case "Relacionamento":
                    retorno = valorCelulaAtual;
                    break;
                case "Data evento":
                    retorno = valorCelulaAtual;
                    break;
                case "Descrição do evento":
                    retorno = valorCelulaAtual;
                    break;
                default:
                    retorno = valorCelulaAtual;
                    break;
            }

            return retorno;
        }
    }
}

