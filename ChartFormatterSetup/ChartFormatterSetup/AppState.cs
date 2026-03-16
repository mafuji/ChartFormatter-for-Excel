using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChartFormatterSetup
{
    public sealed class AppState
    {
        // xlamからの直接更新時、自身が配置されるフォルダ名
        private const string myFolderName = "ChartFormatter_0qc34mctq94thm9q";

        // 自爆キーワード
        private const string bomb = "bomb_mc4893thug9m8c58v59m4y9mh";

        private static readonly AppState _instance = new AppState();
        public static AppState Instance => _instance;

        // --- 状態データ ---
        public bool IsBomb { get; private set; }

        private AppState() { }

        /// <summary>
        /// アプリ起動時に一度だけ呼ぶ初期化メソッド
        /// </summary>
        public void Initialize(string[] args)
        {
            if (args.Length == 1 && args[0].Equals(bomb, StringComparison.OrdinalIgnoreCase))
            {
                IsBomb = true;
            }
        }

        /// <summary>
        /// 「今の起動パス」が「Tempフォルダ」配下にあるか
        /// </summary>
        public bool IsInTemp()
        {
            string? currentDirRaw = System.Windows.Forms.Application.StartupPath;
            string tempDirRaw = Path.GetTempPath();

            // 比較用に展開＆正規化
            string currentDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(currentDirRaw ?? ""));
            string tempDir = Path.GetFullPath(Environment.ExpandEnvironmentVariables(tempDirRaw));

            var result = currentDir.StartsWith(tempDir, StringComparison.OrdinalIgnoreCase);

            //// Test from ==============================================
            //MessageBox.Show(
            //    $"currentDir(raw): {currentDirRaw}\n" +
            //    $"tempDir(raw): {tempDirRaw}\n\n" +
            //    $"currentDir(expanded+full): {currentDir}\n" +
            //    $"tempDir(expanded+full): {tempDir}\n\n" +
            //    $"Result: {result}",
            //    "Debug step: IsInTemp"
            //);
            //// Test to ==============================================

            return result;
        }

        /// <summary>
        /// 自己削除ロジック
        /// </summary>
        public void ExecuteSelfDestruct()
        {
            // 1. 起動パスとフォルダ名の取得（nullチェックを挟む）
            string? exePath = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exePath)) return;

            string? currentDir = Path.GetDirectoryName(exePath);
            if (string.IsNullOrEmpty(currentDir)) return;

            string? currentFolderName = Path.GetFileName(currentDir);
            // ※GetFileNameは末尾がディレクトリならその名前を返すが、ルート(C:\)ならnullを返す
            if (string.IsNullOrEmpty(currentFolderName)) return;

            // 2. 実行条件の最終確認
            bool isCorrectFolder = currentFolderName.Equals(myFolderName, StringComparison.OrdinalIgnoreCase);

            //// Test from ==============================================
            //MessageBox.Show(
            //    $"exePath: {exePath}\n" +
            //    $"currentDir: {currentDir}\n" +
            //    $"currentFolderName: {currentFolderName}\n" +
            //    $"isCorrectFolder: {isCorrectFolder}\n" +
            //    $"cancel: {!IsBomb || !isCorrectFolder || !IsInTemp()}",
            //    "Debug step: SelfDestruct"
            //);
            //// Test to ==============================================

            // 全条件が揃っているかチェック
            if (!IsBomb || !isCorrectFolder || !IsInTemp())
            {
                return;
            }
#if DEBUG
    if (Debugger.IsAttached) return;
#endif
            // 3. 削除コマンドの構築

            // /c : 実行後に cmd を閉じる
            // ping : 自分のプロセスが終了するまでの待機時間 (約4秒)
            // del /f /q : 強制的に、確認なしでファイルを削除
            string command = $"/c ping 127.0.0.1 -n 5 > nul & del /f /q \"{exePath}\"";

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = command,
                WindowStyle = ProcessWindowStyle.Hidden, // ユーザーには見せない
                CreateNoWindow = true,
                UseShellExecute = false
            };

            // 4. 自爆コマンドをバックグラウンドで発射
            Process.Start(psi);
        }
    }
}
