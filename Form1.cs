namespace conversor;

using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public partial class Form1 : Form
{
    private ComboBox opcaoConversao = null!;
    private Label labelStatus = null!;
    private TextBox txtCaminhoArquivo = null!;
    private Button btnSelecionarArquivo = null!;
    private Button btnConverter = null!;

    public Form1()
    {
        InitializeComponent();
        ConfigurarInterface();
    }

    public void ConfigurarInterface()
    {
        this.Text = "Conversor de Arquivos";
        this.Size = new Size(420, 260);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;

       
        Label lblTipo = new Label { Text = "Tipo de Conversão:", Location = new Point(20, 15), AutoSize = true };

      
        opcaoConversao = new ComboBox();
        opcaoConversao.Location = new Point(20, 35);
        opcaoConversao.Size = new Size(360, 30);
        opcaoConversao.DropDownStyle = ComboBoxStyle.DropDownList;
        opcaoConversao.Items.AddRange(new string[] {
            "Bloco de Texto (.txt) -> PDF (.pdf)",
            "Bloco de Texto (.txt) -> Word (.docx)",
            "Word (.docx) -> Bloco de Texto (.txt)" 
        });
        opcaoConversao.SelectedIndex = 0;

       
        Label lblArquivo = new Label { Text = "Arquivo de Origem:", Location = new Point(20, 75), AutoSize = true };

        
        txtCaminhoArquivo = new TextBox { Location = new Point(20, 95), Width = 260, ReadOnly = true };

       
        btnSelecionarArquivo = new Button();
        btnSelecionarArquivo.Text = "Buscar...";
        btnSelecionarArquivo.Location = new Point(290, 93);
        btnSelecionarArquivo.Width = 90;
        btnSelecionarArquivo.Click += BtnSelecionarArquivo_Click;

       
        btnConverter = new Button();
        btnConverter.Text = "Converter Arquivo";
        btnConverter.Location = new Point(20, 135);
        btnConverter.Size = new Size(360, 35);
        btnConverter.Click += BtnConverter_Click;

       
        labelStatus = new Label { Text = "Aguardando arquivo...", Location = new Point(20, 185), AutoSize = true };

       
        this.Controls.Add(lblTipo);
        this.Controls.Add(opcaoConversao);
        this.Controls.Add(lblArquivo);
        this.Controls.Add(txtCaminhoArquivo);
        this.Controls.Add(btnSelecionarArquivo);
        this.Controls.Add(btnConverter);
        this.Controls.Add(labelStatus);
    }

    private void BtnSelecionarArquivo_Click(object? sender, EventArgs e)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Title = "Selecione o arquivo de origem";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtCaminhoArquivo.Text = openFileDialog.FileName;
                labelStatus.Text = "Arquivo selecionado: " + openFileDialog.SafeFileName;
            }
        }
    }

    private void BtnConverter_Click(object? sender, EventArgs e)
    {
        string caminhoArquivo = txtCaminhoArquivo.Text;

        if (string.IsNullOrEmpty(caminhoArquivo) || !File.Exists(caminhoArquivo))
        {
            MessageBox.Show("Por favor, selecione um arquivo válido.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            switch (opcaoConversao.SelectedIndex)
            {
                case 0:
                    ConversorTxtParaPdf(caminhoArquivo);
                    break;

                case 1: 
                    ConversorTxtParaWord(caminhoArquivo);
                    break;

                case 2: 
                    ConversorWordParaTxt(caminhoArquivo);
                    break;

                default:
                    MessageBox.Show("Opção de conversão inválida.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Ocorreu um erro durante a conversão: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

   
    public void ConversorTxtParaPdf(string caminhoArquivo)
    {
        using (SaveFileDialog saveDialog = new SaveFileDialog { Filter = "Arquivo PDF (*.pdf)|*.pdf" })
        {
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                string conteudo = File.ReadAllText(caminhoArquivo);

                using (PdfDocument pdf = new PdfDocument())
                {
                    PdfPage page = pdf.AddPage();
                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XFont font = new XFont("Arial", 12, XFontStyle.Regular);

                    XRect rect = new XRect(40, 40, page.Width - 80, page.Height - 80);
                    gfx.DrawString(conteudo, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    pdf.Save(saveDialog.FileName);
                }

                labelStatus.Text = "Convertido para PDF com sucesso!";
                MessageBox.Show("PDF gerado com sucesso!", "Concluído", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    
    public void ConversorTxtParaWord(string caminhoArquivo)
    {
        using (SaveFileDialog saveDialog = new SaveFileDialog { Filter = "Arquivo Word (*.docx)|*.docx" })
        {
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                string conteudo = File.ReadAllText(caminhoArquivo);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(saveDialog.FileName, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                   
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text(conteudo));
                }

                labelStatus.Text = "Convertido para Word com sucesso!";
                MessageBox.Show("Documento Word gerado com sucesso!", "Concluído", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

  
    public void ConversorWordParaTxt(string caminhoArquivo)
    {
        using (SaveFileDialog saveDialog = new SaveFileDialog { Filter = "Arquivo de Texto (*.txt)|*.txt" })
        {
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                string conteudoExtraido = "";

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(caminhoArquivo, false))
                {
                    var body = wordDocument.MainDocumentPart?.Document?.Body;
                    if (body != null)
                    {
                       
                        conteudoExtraido = body.InnerText;
                    }
                }

                File.WriteAllText(saveDialog.FileName, conteudoExtraido);

                labelStatus.Text = "Convertido para Bloco de Texto com sucesso!";
                MessageBox.Show("Bloco de Texto gerado com sucesso!", "Concluído", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}