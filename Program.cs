using System;
using System.IO;
using System.Windows.Forms; // Windowsフォームを利用
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

class ExtractTextWithFontSizeToFile
{
    [STAThread] // 必須アトリビュート
    static void Main(string[] args)
    {
        // ファイル選択ダイアログを開く
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            Title = "PDFファイルを選択してください",
            Filter = "PDFファイル (*.pdf)|*.pdf",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) // 初期表示フォルダをデスクトップに設定
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            string pdfPath = openFileDialog.FileName;

            // 保存先フォルダ: ダウンロードフォルダ
            string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

            // 出力ファイル名
            string outputFilePath = Path.Combine(downloadsPath, "ExtractedTextWithFontSize.txt");

            try
            {
                using (var pdf = PdfDocument.Open(pdfPath))
                {
                    using (StreamWriter writer = new StreamWriter(outputFilePath))
                    {
                        writer.WriteLine("--- PDFから抽出した内容 ---\n");

                        foreach (var page in pdf.GetPages())
                        {
                            writer.WriteLine($"ページ {page.Number}:\n");

                            foreach (var letter in page.Letters)
                            {
                                // 文字、フォントサイズ、座標を書き出し
                                writer.WriteLine($"テキスト: {letter.Value}, フォントサイズ: {letter.FontSize}, 座標: ({letter.GlyphRectangle.Left}, {letter.GlyphRectangle.Top})");
                            }

                            writer.WriteLine("\n-------------------\n");
                        }
                    }
                }

                // 保存成功メッセージ
                MessageBox.Show($"PDFの内容を以下の場所に保存しました:\n{outputFilePath}", "処理成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // メモ帳で出力ファイルを開く
                System.Diagnostics.Process.Start("notepad.exe", outputFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        else
        {
            MessageBox.Show("ファイル選択がキャンセルされました。", "キャンセル", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
