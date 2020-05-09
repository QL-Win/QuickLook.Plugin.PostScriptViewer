using System;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickLook.Common.Plugin;
using QuickLook.Plugin.PDFViewer;

namespace QuickLook.Plugin.PostScriptViewer
{
    public class Plugin : IViewer
    {
        private static readonly string[] Extensions = {".ps", ".eps"};

        private MemoryStream _buffer;

        private ContextObject _context;
        private string _path;
        private PdfViewerControl _pdfControl;

        public int Priority => 0;

        public void Init()
        {
        }

        public bool CanHandle(string path)
        {
            return !Directory.Exists(path) && Extensions.Any(path.ToLower().EndsWith);
        }

        public void Prepare(string path, ContextObject context)
        {
            _context = context;
            _path = path;

            var width = 800d;
            var height = 600d;

            var pages = GhostScriptWrapper.GetPageSizes(path);
            pages?.ForEach(p =>
            {
                width = Math.Max(width, p.Width);
                height = Math.Max(height, p.Height);
            });

            context.SetPreferredSizeFit(new Size(width, height), 0.8);
        }

        public void View(string path, ContextObject context)
        {
            _pdfControl = new PdfViewerControl();
            context.ViewerContent = _pdfControl;

            _buffer = GhostScriptWrapper.ConvertToPdf(path);
            if (_buffer == null || _buffer.Length == 0)
            {
                context.ViewerContent = new Label {Content = "Conversion to PDF failed."};
                context.IsBusy = false;
                return;
            }

            Exception exception = null;

            _pdfControl.Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    _pdfControl.LoadPdf(_buffer);

                    context.Title = $"1 / {_pdfControl.TotalPages}: {Path.GetFileName(path)}";

                    _pdfControl.CurrentPageChanged += UpdateWindowCaption;
                    context.IsBusy = false;
                }
                catch (Exception e)
                {
                    exception = e;
                }
            }), DispatcherPriority.Loaded).Wait();

            if (exception != null)
                ExceptionDispatchInfo.Capture(exception).Throw();
        }

        public void Cleanup()
        {
            _pdfControl?.Dispose();
            _pdfControl = null;
            _context = null;
            _buffer = null;
        }

        private void UpdateWindowCaption(object sender, EventArgs e2)
        {
            _context.Title = $"{_pdfControl.CurrentPage + 1} / {_pdfControl.TotalPages}: {Path.GetFileName(_path)}";
        }
    }
}