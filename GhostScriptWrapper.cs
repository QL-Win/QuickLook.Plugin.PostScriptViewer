using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;

namespace QuickLook.Plugin.PostScriptViewer
{
    internal static class GhostScriptWrapper
    {
        private static readonly string GhostScriptPath =
            Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "gs9.52\\gswin32c.exe");

        public static MemoryStream ConvertToPdf(string file)
        {
            var arg = $@"-sDEVICE=pdfwrite -dEPSCrop -q -o- ""{file}""";

            var stdout = Execute(arg);

            return stdout;
        }

        public static List<Size> GetPageSizes(string file)
        {
            var result = new List<Size>();

            var arg =
                $@"-q -dNODISPLAY -dBATCH -dNOPAUSE -c ""/showpage{{currentpagedevice /PageSize get{{=}}forall showpage}}bind def"" -f ""{file}""";

            var stdout = Encoding.ASCII.GetString(Execute(arg)?.ToArray() ?? new byte[0]);
            if (string.IsNullOrEmpty(stdout))
                return new List<Size>();

            var lines = stdout.Replace("\r\n", "\n").Split('\n');
            for (var i = 0; i < lines.Length - 1; i += 2)
            {
                if (string.IsNullOrWhiteSpace(lines[i]) &&
                    string.IsNullOrWhiteSpace(lines[i + 1]))
                    continue;

                if (!double.TryParse(lines[i], out var width) || !double.TryParse(lines[i + 1], out var height))
                    continue;

                result.Add(new Size(width, height));
            }

            return result;
        }

        private static MemoryStream Execute(string arguments)
        {
            var proc = new Process
            {
                StartInfo =
                {
                    FileName = GhostScriptPath,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    UseShellExecute = false
                }
            };

            proc.Start();

            var ms = new MemoryStream();
            proc.StandardOutput.BaseStream.CopyTo(ms);

            return proc.ExitCode != 0 ? null : ms;
        }
    }
}