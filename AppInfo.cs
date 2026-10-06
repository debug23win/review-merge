using System.Reflection;
[assembly: AssemblyTitle("Свод проверки")]
[assembly: AssemblyVersion("1.3.0.0")]
[assembly: AssemblyFileVersion("1.3.0.0")]
namespace DesktopUpdates {
    internal static class AppInfo {
        public const string Repository="debug23win/review-merge",Product="review-merge",Asset="review-merge-windows.zip",Executable="Свод_проверки.exe";
        public static readonly string[] UpdateFiles={Executable,"README.md"};
    }
}
