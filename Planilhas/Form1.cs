using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Planilhas
{
    public partial class Planilhas : Form
    {
        public Planilhas()
        {
            InitializeComponent();
        }


        // Método para abrir e carregar a planilha no DataGridView
        private void AbrirArquivoPlanilha()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Arquivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Arquivos CSV (*.csv)|*.csv|Todos os Arquivos (*.*)|*.*";
                openFileDialog.Title = "Abrir Planilha";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = openFileDialog.FileName;
                        DataTable dataTable;

                        if (Path.GetExtension(filePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
                        {
                            dataTable = CarregarCsv(filePath);
                        }
                        else
                        {
                            dataTable = CarregarExcel(filePath);
                        }

                        dgvPlanilha.DataSource = dataTable;
                        MessageBox.Show("Arquivo carregado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao abrir o arquivo:\n{ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // Método para salvar os dados do DataGridView em um arquivo
        private void SalvarArquivoPlanilha()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Arquivos Excel (*.xlsx)|*.xlsx|Arquivos CSV (*.csv)|*.csv|Todos os Arquivos (*.*)|*.*";
                saveFileDialog.Title = "Salvar Planilha";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = saveFileDialog.FileName;

                        if (Path.GetExtension(filePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
                        {
                            SalvarComoCsv(filePath);
                        }
                        else
                        {
                            SalvarComoExcel(filePath);
                        }

                        MessageBox.Show("Arquivo salvo com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao salvar o arquivo:\n{ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // Métodos auxiliares
        private DataTable CarregarCsv(string filePath)
        {
            DataTable dataTable = new DataTable();
            string[] lines = File.ReadAllLines(filePath);

            if (lines.Length > 0)
            {
                string[] headers = lines[0].Split(',');

                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header.Trim());
                }

                for (int i = 1; i < lines.Length; i++)
                {
                    string[] row = lines[i].Split(',');
                    dataTable.Rows.Add(row);
                }
            }

            return dataTable;
        }

        private DataTable CarregarExcel(string filePath)
        {
            DataTable dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                for (int col = 1; col <= columns; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                }

                for (int row = 2; row <= rows; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= columns; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        private void SalvarComoCsv(string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                for (int col = 0; col < dgvPlanilha.Columns.Count; col++)
                {
                    writer.Write(dgvPlanilha.Columns[col].HeaderText);
                    if (col < dgvPlanilha.Columns.Count - 1)
                        writer.Write(",");
                }
                writer.WriteLine();

                foreach (DataGridViewRow row in dgvPlanilha.Rows)
                {
                    for (int col = 0; col < dgvPlanilha.Columns.Count; col++)
                    {
                        writer.Write(row.Cells[col].Value?.ToString());
                        if (col < dgvPlanilha.Columns.Count - 1)
                            writer.Write(",");
                    }
                    writer.WriteLine();
                }
            }
        }

        private void SalvarComoExcel(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Planilha");

                for (int col = 0; col < dgvPlanilha.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dgvPlanilha.Columns[col].HeaderText;
                }

                for (int row = 0; row < dgvPlanilha.Rows.Count; row++)
                {
                    for (int col = 0; col < dgvPlanilha.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dgvPlanilha.Rows[row].Cells[col].Value?.ToString();
                    }
                }

                package.Save();
            }
        }

        private void abrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AbrirArquivoPlanilha();
        }

        private void salvarComoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SalvarArquivoPlanilha();
        }

        private void CriarNovaPlanilha()
        {
            // Limpa o DataGridView e define colunas padrão
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Coluna 1");
            dataTable.Columns.Add("Coluna 2");
            dataTable.Columns.Add("Coluna 3");

            dgvPlanilha.DataSource = dataTable;

            MessageBox.Show("Nova planilha criada!", "Nova Planilha", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void novoToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            // Verifica se há informações no DataGridView
            if (dgvPlanilha.Rows.Count > 0)
            {
                // Pergunta ao usuário se deseja salvar as alterações
                var result = MessageBox.Show(
                    "Existem dados na planilha atual. Deseja salvar as alterações antes de criar uma nova planilha?",
                    "Salvar Alterações",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Chama o método de salvar
                    SalvarArquivoPlanilha();

                    // Limpa o DataGridView
                    dgvPlanilha.DataSource = null;
                    dgvPlanilha.Rows.Clear();
                    dgvPlanilha.Columns.Clear();
                }
                else if (result == DialogResult.Cancel)
                {
                    // Cancela a ação de criar uma nova planilha
                    return;
                }
            }
            // Cria uma nova planilha (limpa o DataGridView)
            CriarNovaPlanilha();
        }
    }
}
