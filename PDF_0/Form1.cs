using System;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PDF_0
{
    public partial class Form1 : Form
    {
        //PDFの結合を行うプログラム
        //このプロジェクトのように、追加パッケージを入れたものは
        //実行ファイルのみで動くわけがないことに注意しなければならない(1敗)
        public Form1()
        {
            InitializeComponent();
        }
        public string DesktopFilepath = "C:\\Users\\" + Environment.UserName + "\\Desktop\\";
        private void Form1_Load(object sender, EventArgs e)
        {
            Text = "PDF結合専用ソフト";
            Open_setting();
            Size = new System.Drawing.Size(400, 100);
        }
        void Open_setting()
        {
            openFileDialog1.DefaultExt = "pdf";
            openFileDialog1.Filter = "PDF|*.pdf*";
            openFileDialog1.Title = "結合するファイルを選択する";
            saveFileDialog1.DefaultExt = "pdf";
            saveFileDialog1.Filter = "PDF|*.pdf*";
            saveFileDialog1.Title = "結合したファイルを保存する";
        }
        private void label1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel) return;
            string[] files = openFileDialog1.FileNames;
            //結合後の出力ファイル
            PdfDocument outputPdfDocument = new PdfDocument();
            foreach (string file in files)
            {
                //ファイルを開く
                PdfDocument inputPdfDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import);
                //１ページずつ出力ファイルに追加
                for (int i = 0; i < inputPdfDocument.PageCount; i++)
                {
                    outputPdfDocument.AddPage(inputPdfDocument.Pages[i]);
                }
                //ファイルを閉じる
                inputPdfDocument.Close();
            }
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel) return;
            //出力
            outputPdfDocument.Save(saveFileDialog1.FileName);
        }
    }
}
