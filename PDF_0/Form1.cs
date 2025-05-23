﻿using System;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PDF_0
{
    public partial class Form1 : Form
    {
        //PDFの結合・分割を行うプログラム
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
            openFileDialog1.Multiselect = true;
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
            }
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel) return;
            //出力
            outputPdfDocument.Save(saveFileDialog1.FileName);
        }
        private void label2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "分割するファイルを指定";
            openFileDialog1.Multiselect = false;
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel) return;
            // 元のPDFファイルのパス
            string inputFilePath = openFileDialog1.FileName;
            // PDFを読み込み（修正可能として開く）
            PdfDocument inputDocument = PdfReader.Open(inputFilePath, PdfDocumentOpenMode.Import);
            if (folderBrowserDialog1.ShowDialog() == DialogResult.Cancel) return;
            string outputFilePath = ""; 
            // 各ページごとに新しいPDFを作成して保存
            for (int idx = 0; idx < inputDocument.PageCount; idx++)
            {
                // 新しいPDFドキュメント作成
                PdfDocument outputDocument = new PdfDocument();
                outputDocument.Version = inputDocument.Version;
                outputDocument.Info.Title = $"Page {idx + 1} of {inputDocument.Info.Title}";
                outputDocument.Info.Creator = inputDocument.Info.Creator;
                // 対象ページを追加
                outputDocument.AddPage(inputDocument.Pages[idx]);
                // ファイル名を作成
                 outputFilePath = $"{folderBrowserDialog1.SelectedPath}\\output_page_{idx + 1}.pdf";
                // 保存
                outputDocument.Save(outputFilePath);
                Console.WriteLine($"Saved: {outputFilePath}");
            }
        }
    }
}
