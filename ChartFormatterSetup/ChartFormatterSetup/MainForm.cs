using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChartFormatterSetup
{
    public partial class MainForm : Form
    {
        private readonly AddInService _service;
        private System.Windows.Forms.Timer _timer;

        public MainForm()
        {
            InitializeComponent();
            _service = new AddInService(); // インスタンス化
            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 1000; // 1秒ごとに監視
            _timer.Tick += (s, e) => RefreshUI();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            RefreshUI();
            _timer.Start();
        }

        private void RefreshUI()
        {
            bool isInstalled = _service.IsInstalled();
            bool isExcelRunning = _service.IsExcelRunning();

            if (isInstalled)
            {
                btnInstallOrUpdate.Text = "アップデート";
                btnUninstall.Visible = true;
                lblGuide.Text = "アドインは既にインストールされています。" +
                                $"\n\n現在の保存先: {_service.TargetPath}" +
                                "\n\n「アンインストール」：アドインを無効化して削除" +
                                "\n「アップデート」：アドインのバージョンを更新";
            }
            else
            {
                btnInstallOrUpdate.Text = "インストール";
                btnUninstall.Visible = false;
                lblGuide.Text = "「インストール」をクリックすると、Excelアドインを既定の場所に配置し、有効化します。" +
                                $"\n\n既定の保存先: {_service.TargetPath}";
            }

            if (isExcelRunning)
            {
                lblGuide.Text += "\n\n【注意】Excelが起動しています。実行前にExcelを閉じてください。";
                btnInstallOrUpdate.Enabled = false;
                btnUninstall.Enabled = false;
            }
            else
            {
                btnInstallOrUpdate.Enabled = true;
                btnUninstall.Enabled = true;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnInstallOrUpdate_Click(object sender, EventArgs e)
        {
            _timer.Stop();

            // 1. 実行前に確認（現在のボタンテキストを取得してメッセージに反映）
            string actionName = btnInstallOrUpdate.Text; // "インストール" または "アップデート"
            var result = MessageBox.Show(
                $"{actionName}を開始しますか？",
                "実行確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
            {
                _timer.Start();
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                // 2. サービス実行
                _service.ExtractAndInstall();

                MessageBox.Show($"{actionName}が完了しました。Excelを起動して確認してください。",
                    "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 成功したときだけ閉じる
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{actionName}失敗：\n{ex.Message}",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // 失敗したときは、RefreshUI
                _timer.Start();
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void btnUninstall_Click(object sender, EventArgs e)
        {
            _timer.Stop();

            var result = MessageBox.Show(
                "アドインを削除しますか？", "確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question
            );

            if (result != DialogResult.Yes)
            {
                _timer.Start();
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                // サービス側で無効化と削除を実行
                _service.Uninstall();

                MessageBox.Show("アドインを正常に削除しました。", "完了",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                Application.Exit();
            }
            catch (Exception ex)
            {
                // 失敗した場合はユーザーに伝え、アプリを終了させない
                MessageBox.Show(
                    $"削除に失敗しました。Excelやエクスプローラーでファイルが開かれていないか確認してください。\n\n詳細: {ex.Message}",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                _timer.Start();
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
    }
}