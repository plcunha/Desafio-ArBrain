using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Xml.Linq;

namespace ArBrain
{
    public partial class FormCadastroMetas : Form
    {
        private List<Meta> listaMetas = new List<Meta>();
        private Stack<Meta> historicoMetas = new Stack<Meta>();
        private List<string> historicoOperacoes = new List<string>();
        private string filtroAtual = "";
        private bool dadosNaoSalvos = false;

        private ComboBox comboBoxTipoMeta;
        private ComboBox comboBoxPeriodicidade;
        private ComboBox comboBoxProduto;
        private TextBox textBoxVendedor;
        private NumericUpDown numericValor;
        private DataGridView dataGridViewMetas;
        private Button btnSalvar;
        private Button btnExcluir;
        private Button btnDuplicar;
        private Button btnExportarExcel;
        private Button btnExportarPDF;
        private Button btnVoltar;
        private Button btnAdicionar;
        private Button btnBuscar;
        private Label labelLegenda;
        private PictureBox pictureBoxLupa;
        private DataGridViewImageColumn colunaAtivo;
        private TextBox textBoxBusca;

        private readonly Color CorPrimaria = Color.FromArgb(0, 123, 255); // Azul
        private readonly Color CorSecundaria = Color.FromArgb(255, 193, 7); // Amarelo
        private readonly Color CorErro = Color.FromArgb(220, 53, 69); // Vermelho
        private readonly Color CorSucesso = Color.FromArgb(40, 167, 69); // Verde

        public FormCadastroMetas()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // Configuração do Form
            this.Text = "Cadastro de Metas";
            this.Size = new System.Drawing.Size(800, 500);
            this.BackColor = Color.White;
            this.KeyPreview = true;
            this.KeyDown += FormCadastroMetas_KeyDown;
            this.FormClosing += FormCadastroMetas_FormClosing;
            this.Load += FormCadastroMetas_Load;

            // ComboBox Tipo de Meta
            this.comboBoxTipoMeta = new ComboBox();
            this.comboBoxTipoMeta.Items.AddRange(new string[] { "Venda", "Atendimento", "Outros" });
            this.comboBoxTipoMeta.Location = new System.Drawing.Point(20, 20);
            this.comboBoxTipoMeta.Size = new System.Drawing.Size(150, 20);
            this.comboBoxTipoMeta.SelectedIndex = 0;

            // ComboBox Periodicidade
            this.comboBoxPeriodicidade = new ComboBox();
            this.comboBoxPeriodicidade.Items.AddRange(new string[] { "Diária", "Semanal", "Mensal" });
            this.comboBoxPeriodicidade.Location = new System.Drawing.Point(20, 50);
            this.comboBoxPeriodicidade.Size = new System.Drawing.Size(150, 20);
            this.comboBoxPeriodicidade.SelectedIndex = 0;

            // ComboBox Produto
            this.comboBoxProduto = new ComboBox();
            this.comboBoxProduto.Items.AddRange(new string[] { "Produto A", "Produto B", "Produto C" });
            this.comboBoxProduto.Location = new System.Drawing.Point(20, 80);
            this.comboBoxProduto.Size = new System.Drawing.Size(150, 20);
            this.comboBoxProduto.SelectedIndex = 0;

            // TextBox Vendedor
            this.textBoxVendedor = new TextBox();
            this.textBoxVendedor.Location = new System.Drawing.Point(20, 110);
            this.textBoxVendedor.Size = new System.Drawing.Size(200, 20);
            this.textBoxVendedor.Text = "Nome do vendedor";
            this.textBoxVendedor.ForeColor = System.Drawing.Color.Gray;
            this.textBoxVendedor.GotFocus += TextBoxVendedor_GotFocus;
            this.textBoxVendedor.LostFocus += TextBoxVendedor_LostFocus;

            // Numeric Valor
            this.numericValor = new NumericUpDown();
            this.numericValor.Location = new System.Drawing.Point(20, 140);
            this.numericValor.Size = new System.Drawing.Size(100, 20);
            this.numericValor.Minimum = 1;
            this.numericValor.Maximum = 1000000;

            // DataGridView Metas
            this.dataGridViewMetas = new DataGridView();
            this.dataGridViewMetas.Location = new System.Drawing.Point(20, 170);
            this.dataGridViewMetas.Size = new System.Drawing.Size(750, 150);
            this.dataGridViewMetas.AutoGenerateColumns = true;

            // Botão Salvar
            this.btnSalvar = new Button();
            this.btnSalvar.Text = "✓ Salvar";
            this.btnSalvar.Location = new System.Drawing.Point(20, 330);
            this.btnSalvar.BackColor = CorPrimaria;
            this.btnSalvar.ForeColor = Color.White;
            this.btnSalvar.Click += btnSalvar_Click;
            this.btnSalvar.MouseHover += (sender, e) => btnSalvar.BackColor = Color.LightGreen;
            this.btnSalvar.MouseLeave += (sender, e) => btnSalvar.BackColor = CorPrimaria;
            this.btnSalvar.Cursor = Cursors.Hand;

            // Botão Excluir
            this.btnExcluir = new Button();
            this.btnExcluir.Text = "Excluir";
            this.btnExcluir.Location = new System.Drawing.Point(100, 330);
            this.btnExcluir.BackColor = Color.Yellow;
            this.btnExcluir.ForeColor = Color.Black;
            this.btnExcluir.FlatStyle = FlatStyle.Flat;
            this.btnExcluir.FlatAppearance.BorderColor = Color.Black;
            this.btnExcluir.Click += btnExcluir_Click;
            this.btnExcluir.MouseHover += (sender, e) => btnExcluir.BackColor = Color.LightCoral;
            this.btnExcluir.MouseLeave += (sender, e) => btnExcluir.BackColor = Color.Yellow;
            this.btnExcluir.Cursor = Cursors.Hand;

            // Botão Duplicar
            this.btnDuplicar = new Button();
            this.btnDuplicar.Text = "Duplicar";
            this.btnDuplicar.Location = new System.Drawing.Point(180, 330);
            this.btnDuplicar.BackColor = CorSecundaria;
            this.btnDuplicar.ForeColor = Color.White;
            this.btnDuplicar.Click += btnDuplicar_Click;
            this.btnDuplicar.MouseHover += (sender, e) => btnDuplicar.BackColor = Color.LightBlue;
            this.btnDuplicar.MouseLeave += (sender, e) => btnDuplicar.BackColor = CorSecundaria;
            this.btnDuplicar.Cursor = Cursors.Hand;

            // Botão Exportar Excel
            this.btnExportarExcel = new Button();
            this.btnExportarExcel.Text = "Exportar Excel";
            this.btnExportarExcel.Location = new System.Drawing.Point(260, 330);
            this.btnExportarExcel.BackColor = Color.DarkCyan;
            this.btnExportarExcel.ForeColor = Color.White;
            this.btnExportarExcel.Click += btnExportarExcel_Click;
            this.btnExportarExcel.Cursor = Cursors.Hand;

            // Botão Exportar PDF
            this.btnExportarPDF = new Button();
            this.btnExportarPDF.Text = "Exportar PDF";
            this.btnExportarPDF.Location = new System.Drawing.Point(360, 330);
            this.btnExportarPDF.BackColor = Color.DarkCyan;
            this.btnExportarPDF.ForeColor = Color.White;
            this.btnExportarPDF.Click += btnExportarPDF_Click;
            this.btnExportarPDF.Cursor = Cursors.Hand;

            // Botão Voltar
            this.btnVoltar = new Button();
            this.btnVoltar.Text = "Voltar";
            this.btnVoltar.Location = new System.Drawing.Point(440, 330);
            this.btnVoltar.BackColor = Color.Gray;
            this.btnVoltar.ForeColor = Color.White;
            this.btnVoltar.FlatStyle = FlatStyle.Flat;
            this.btnVoltar.FlatAppearance.BorderColor = Color.Black;
            this.btnVoltar.Click += btnVoltar_Click;
            this.btnVoltar.Cursor = Cursors.Hand;

            // Botão Adicionar
            this.btnAdicionar = new Button();
            this.btnAdicionar.Text = "+";
            this.btnAdicionar.Location = new System.Drawing.Point(520, 330);
            this.btnAdicionar.BackColor = Color.Green;
            this.btnAdicionar.ForeColor = Color.White;
            this.btnAdicionar.Click += btnAdicionar_Click;
            this.btnAdicionar.Cursor = Cursors.Hand;

            // Botão Buscar
            this.btnBuscar = new Button();
            this.btnBuscar.Text = "Buscar";
            this.btnBuscar.Location = new System.Drawing.Point(600, 330);
            this.btnBuscar.BackColor = Color.DarkCyan;
            this.btnBuscar.ForeColor = Color.White;
            this.btnBuscar.Click += btnBuscar_Click;
            this.btnBuscar.Cursor = Cursors.Hand;

            // TextBox Busca
            this.textBoxBusca = new TextBox();
            this.textBoxBusca.Location = new System.Drawing.Point(20, 390);
            this.textBoxBusca.Size = new System.Drawing.Size(200, 20);
            this.textBoxBusca.Text = "Digite para buscar";
            this.textBoxBusca.ForeColor = System.Drawing.Color.Gray;
            this.textBoxBusca.GotFocus += TextBoxBusca_GotFocus;
            this.textBoxBusca.LostFocus += TextBoxBusca_LostFocus;

            // Legenda
            this.labelLegenda = new Label();
            this.labelLegenda.Text = "Filtrar por Vendedor";
            this.labelLegenda.Location = new System.Drawing.Point(20, 360);
            this.labelLegenda.ForeColor = Color.Gray;

            // Adicionando controles ao Form
            this.Controls.Add(this.comboBoxTipoMeta);
            this.Controls.Add(this.comboBoxPeriodicidade);
            this.Controls.Add(this.comboBoxProduto);
            this.Controls.Add(this.textBoxVendedor);
            this.Controls.Add(this.numericValor);
            this.Controls.Add(this.dataGridViewMetas);
            this.Controls.Add(this.btnSalvar);
            this.Controls.Add(this.btnExcluir);
            this.Controls.Add(this.btnDuplicar);
            this.Controls.Add(this.btnExportarExcel);
            this.Controls.Add(this.btnExportarPDF);
            this.Controls.Add(this.btnVoltar);
            this.Controls.Add(this.btnAdicionar);
            this.Controls.Add(this.btnBuscar);
            this.Controls.Add(this.textBoxBusca);
            this.Controls.Add(this.labelLegenda);
            this.Controls.Add(this.pictureBoxLupa);
        }

        private void FormCadastroMetas_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }

            if (e.KeyCode == Keys.F2)
            {
                btnSalvar_Click(sender, e);
            }
        }

        private void TextBoxVendedor_GotFocus(object sender, EventArgs e)
        {
            if (textBoxVendedor.Text == "Nome do vendedor")
            {
                textBoxVendedor.Text = "";
                textBoxVendedor.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void TextBoxVendedor_LostFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxVendedor.Text))
            {
                textBoxVendedor.Text = "Nome do vendedor";
                textBoxVendedor.ForeColor = System.Drawing.Color.Gray;
            }
        }

        private void TextBoxBusca_GotFocus(object sender, EventArgs e)
        {
            if (textBoxBusca.Text == "Digite para buscar")
            {
                textBoxBusca.Text = "";
                textBoxBusca.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void TextBoxBusca_LostFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxBusca.Text))
            {
                textBoxBusca.Text = "Digite para buscar";
                textBoxBusca.ForeColor = System.Drawing.Color.Gray;
            }
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            if (ValidarCampos())
            {
                Meta novaMeta = new Meta(
                    textBoxVendedor.Text,
                    comboBoxTipoMeta.SelectedItem.ToString(),
                    comboBoxProduto.SelectedItem.ToString(),
                    numericValor.Value,
                    comboBoxPeriodicidade.SelectedItem.ToString()
                );
                listaMetas.Add(novaMeta);
                historicoMetas.Push(novaMeta);
                historicoOperacoes.Add($"Meta salva: {novaMeta.Vendedor} - {novaMeta.Tipo} - {novaMeta.Produto} - {novaMeta.Valor} - {novaMeta.Periodicidade}");
                AtualizarGrid();
                LimparCampos();
                dadosNaoSalvos = false;
                MessageBox.Show("Meta salva com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private bool ValidarCampos()
        {
            if (string.IsNullOrWhiteSpace(textBoxVendedor.Text) || textBoxVendedor.Text == "Nome do vendedor")
            {
                textBoxVendedor.BackColor = CorErro;
                MessageBox.Show("O campo Vendedor é obrigatório.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else
            {
                textBoxVendedor.BackColor = Color.White;
            }

            if (comboBoxTipoMeta.SelectedItem == null || comboBoxProduto.SelectedItem == null)
            {
                MessageBox.Show("Selecione um tipo de meta e um produto.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (numericValor.Value <= 0)
            {
                MessageBox.Show("O valor da meta deve ser maior que zero.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void AtualizarGrid()
        {
            dataGridViewMetas.DataSource = null;
            dataGridViewMetas.DataSource = listaMetas;
        }

        private void LimparCampos()
        {
            textBoxVendedor.Text = "Nome do vendedor";
            textBoxVendedor.ForeColor = System.Drawing.Color.Gray;
            comboBoxTipoMeta.SelectedIndex = 0;
            comboBoxProduto.SelectedIndex = 0;
            numericValor.Value = 1;
            comboBoxPeriodicidade.SelectedIndex = 0;
        }

        private void btnExcluir_Click(object sender, EventArgs e)
        {
            if (dataGridViewMetas.SelectedRows.Count > 0)
            {
                listaMetas.RemoveAt(dataGridViewMetas.SelectedRows[0].Index);
                AtualizarGrid();
            }
            else
            {
                MessageBox.Show("Selecione uma meta para excluir.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDuplicar_Click(object sender, EventArgs e)
        {
            if (dataGridViewMetas.SelectedRows.Count > 0)
            {
                Meta metaSelecionada = listaMetas[dataGridViewMetas.SelectedRows[0].Index];
                listaMetas.Add(new Meta(metaSelecionada.Vendedor, metaSelecionada.Tipo, metaSelecionada.Produto, metaSelecionada.Valor, metaSelecionada.Periodicidade));
                AtualizarGrid();
            }
            else
            {
                MessageBox.Show("Selecione uma meta para duplicar.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportarExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Metas");
                    worksheet.Cell(1, 1).Value = "Vendedor";
                    worksheet.Cell(1, 2).Value = "Tipo";
                    worksheet.Cell(1, 3).Value = "Produto";
                    worksheet.Cell(1, 4).Value = "Valor";
                    worksheet.Cell(1, 5).Value = "Periodicidade";

                    int row = 2;
                    foreach (var meta in listaMetas)
                    {
                        worksheet.Cell(row, 1).Value = meta.Vendedor;
                        worksheet.Cell(row, 2).Value = meta.Tipo;
                        worksheet.Cell(row, 3).Value = meta.Produto;
                        worksheet.Cell(row, 4).Value = meta.Valor.ToString("#,##0.00");
                        worksheet.Cell(row, 5).Value = meta.Periodicidade;
                        row++;
                    }

                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Excel Files|*.xlsx";
                    saveDialog.Title = "Salvar Excel";
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        workbook.SaveAs(saveDialog.FileName);
                        MessageBox.Show("Exportação para Excel concluída!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao exportar para Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportarPDF_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "PDF Files|*.pdf";
                saveDialog.Title = "Salvar PDF";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    Document doc = new Document();
                    PdfWriter.GetInstance(doc, new FileStream(saveDialog.FileName, FileMode.Create));
                    doc.Open();

                    // Título
                    iTextSharp.text.Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18);
                    Paragraph title = new Paragraph("Relatório de Metas", titleFont);
                    title.Alignment = Element.ALIGN_CENTER;
                    doc.Add(title);

                    // Tabela de dados
                    PdfPTable table = new PdfPTable(5);
                    table.WidthPercentage = 100;
                    table.SetWidths(new float[] { 2f, 2f, 2f, 2f, 2f });

                    // Cabeçalho
                    table.AddCell("Vendedor");
                    table.AddCell("Tipo");
                    table.AddCell("Produto");
                    table.AddCell("Valor");
                    table.AddCell("Periodicidade");

                    // Dados
                    foreach (var meta in listaMetas)
                    {
                        table.AddCell(meta.Vendedor);
                        table.AddCell(meta.Tipo);
                        table.AddCell(meta.Produto);
                        table.AddCell(meta.Valor.ToString("#,##0.00"));
                        table.AddCell(meta.Periodicidade);
                    }

                    doc.Add(table);
                    doc.Close();
                    MessageBox.Show("Exportação para PDF concluída!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao exportar para PDF: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVoltar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAdicionar_Click(object sender, EventArgs e)
        {
            // Lógica para adicionar uma nova meta
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxBusca.Text) || textBoxBusca.Text == "Digite para buscar")
            {
                textBoxBusca.BackColor = Color.FromArgb(252, 199, 194);
                MessageBox.Show("O campo de busca é obrigatório.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                textBoxBusca.BackColor = Color.White;
            }

            filtroAtual = textBoxBusca.Text;
            AplicarFiltro();
        }

        private void AplicarFiltro()
        {
            dataGridViewMetas.DataSource = listaMetas.Where(m => m.Vendedor.IndexOf(filtroAtual, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
        }

        private void pictureBoxLupa_Click(object sender, EventArgs e)
        {
            // Lógica para abrir uma janela de busca
        }

        private void FormCadastroMetas_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dadosNaoSalvos)
            {
                DialogResult result = MessageBox.Show("Existem dados não salvos. Tem certeza que deseja sair?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
        }

        private void FormCadastroMetas_Load(object sender, EventArgs e)
        {
            comboBoxTipoMeta.Select();
        }
    }

    public class Meta
    {
        public string Vendedor { get; set; }
        public string Tipo { get; set; }
        public string Produto { get; set; }
        public decimal Valor { get; set; }
        public string Periodicidade { get; set; }

        public Meta(string vendedor, string tipo, string produto, decimal valor, string periodicidade)
        {
            Vendedor = vendedor;
            Tipo = tipo;
            Produto = produto;
            Valor = valor;
            Periodicidade = periodicidade;
        }
    }
}