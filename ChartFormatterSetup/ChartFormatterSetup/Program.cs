namespace ChartFormatterSetup
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] rawArgs)
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            // 起動時の文脈を構築（bomb 判定 + パス保持）
            AppState.StartupContext = new StartupContext(rawArgs);

            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }
}