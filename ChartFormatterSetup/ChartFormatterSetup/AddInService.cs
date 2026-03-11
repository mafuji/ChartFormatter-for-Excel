using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ChartFormatterSetup
{
    internal class AddInService
    {
        public const string AddInFileName = "ChartFormatter.xlam";
        public string TargetDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\AddIns");
        public string TargetPath => Path.Combine(TargetDir, AddInFileName);

        public bool IsInstalled() => File.Exists(TargetPath);

        // Excelプロセスが1つでもあればtrue
        public bool IsExcelRunning() => Process.GetProcessesByName("EXCEL").Length > 0;

        /// <summary>
        /// インストール・アップデート処理（Excelが終了している前提）
        /// </summary>
        public void ExtractAndInstall()
        {
            // 1. 念のための最終防衛ライン（起動中なら例外を投げる）
            if (IsExcelRunning())
                throw new Exception("Excelが起動しているため、ファイルを上書きできません。");

            // 2. フォルダ準備
            if (!Directory.Exists(TargetDir)) Directory.CreateDirectory(TargetDir);

            // 3. ファイルをリソースから展開（Excelが閉じていれば100%成功する）
            ExtractResourceToFile();

            // 4. Excelをバックグラウンドで一瞬立ち上げ、アドインリストに登録・有効化する
            ExecuteExcelAction(excelApp =>
            {
                var target = FindAddIn(excelApp);
                if (target == null)
                {
                    target = excelApp.AddIns.Add(TargetPath, false);
                }
                target.Installed = true;
            });
        }

        public void Uninstall()
        {
            if (!IsInstalled()) return;

            // 無効化
            ExecuteExcelAction(excelApp =>
            {
                var target = FindAddIn(excelApp);
                if (target != null) target.Installed = false;
            });

            // 物理削除
            try { File.Delete(TargetPath); }
            catch { Thread.Sleep(500); File.Delete(TargetPath); }
        }

        private void ExtractResourceToFile()
        {
            var assembly = Assembly.GetExecutingAssembly();
            string resourceName = $"{assembly.GetName().Name}.{AddInFileName}";
            using (Stream? stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) throw new Exception($"リソース '{resourceName}' 未検出");
                using (FileStream fileStream = new FileStream(TargetPath, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
            }
        }

        private Excel.AddIn? FindAddIn(Excel.Application app)
        {
            foreach (Excel.AddIn ai in app.AddIns)
            {
                if (ai.Name.Equals(AddInFileName, StringComparison.OrdinalIgnoreCase)) return ai;
            }
            return null;
        }

        private void ExecuteExcelAction(Action<Excel.Application> action)
        {
            Excel.Application? excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                action(excelApp);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }
    }
}