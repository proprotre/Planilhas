# Manipulação de Planilhas em C# com Interface Gráfica

- **Tecnologias Utilizadas:**
  - C# .NET Core
  - Windows Forms
  - Biblioteca [EPPlus](https://epplussoftware.com) para manipulação de arquivos Excel (.xlsx)

## Software com uma planilha carregada
![image](/Imagens/planilhasSoftware.png "Seleção de cidades")

## Funcionalidades

- **Carregar Planilha:** 
  Permite abrir um arquivo Excel e exibir os dados no `DataGridView`.

- **Salvar Planilha:** 
  Salva os dados exibidos no `DataGridView` em um arquivo Excel existente ou cria um novo.

- **Criar Nova Planilha:** 
  Gera uma planilha em branco, verificando se há dados não salvos no `DataGridView`.

- **Editar Planilha:** 
  Ao clicar duas vezes sobre uma célula é possível alterar os valores

## Interface do Usuário

A interface inclui:
- Um `DataGridView` para exibição e edição de dados.
- Botões:
  - `Carregar`: Abre e carrega um arquivo Excel no `DataGridView`.
  - `Salvar`: Salva os dados do `DataGridView` em um arquivo Excel.
  - `Novo`: Verifica alterações antes de limpar o `DataGridView` e cria uma nova planilha.

## Funcionalidades Detalhadas

### Novo
- Verifica se há alterações não salvas no `DataGridView`.
- Caso existam dados, solicita ao usuário se deseja salvar antes de limpar e criar uma nova planilha.

### Abrir
- Permite abrir um arquivo `.xlsx` selecionado pelo usuário.
- Os dados do arquivo são carregados no `DataGridView` para visualização e edição.

### Salvar Como
- Salva os dados exibidos no `DataGridView` em um arquivo Excel.
- Caso o arquivo já exista, sobrescreve a planilha com o mesmo nome.

---

## Código Principal

### Novo
```csharp
private void btnNovo_Click(object sender, EventArgs e)
{
    if (dgvPlanilha.Rows.Count > 0)
    {
        var result = MessageBox.Show(
            "Existem dados na planilha atual. Deseja salvar as alterações antes de criar uma nova planilha?",
            "Salvar Alterações",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            btnSalvar.PerformClick();
        }
        else if (result == DialogResult.Cancel)
        {
            return;
        }
    }

    CriarNovaPlanilha();
}
```
### Salvar Como
```csharp
private void btnSalvar_Click(object sender, EventArgs e)
{
    SaveFileDialog saveFileDialog = new SaveFileDialog
    {
        Filter = "Arquivos Excel (*.xlsx)|*.xlsx",
        Title = "Salvar Planilha"
    };

    if (saveFileDialog.ShowDialog() == DialogResult.OK)
    {
        string filePath = saveFileDialog.FileName;

        using (var package = new ExcelPackage())
        {
            var existingWorksheet = package.Workbook.Worksheets["Planilha"];
            if (existingWorksheet != null)
            {
                package.Workbook.Worksheets.Delete("Planilha");
            }

            var worksheet = package.Workbook.Worksheets.Add("Planilha");

            for (int col = 0; col < dgvPlanilha.Columns.Count; col++)
                worksheet.Cells[1, col + 1].Value = dgvPlanilha.Columns[col].HeaderText;

            for (int row = 0; row < dgvPlanilha.Rows.Count; row++)
            {
                for (int col = 0; col < dgvPlanilha.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1].Value = dgvPlanilha.Rows[row].Cells[col].Value;
                }
            }

            package.SaveAs(new FileInfo(filePath));
        }

        MessageBox.Show("Planilha salva com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
```
## Caminho do executável
"Planilhas/bin/Debug/net8.0-windows/Planilhas.exe"

## Autor
Gabriel Badaró