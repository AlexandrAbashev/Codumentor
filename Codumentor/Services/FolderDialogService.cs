using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Codumentor.Services
{
    public static class FolderDialogService
    {
        public static string OpenFolderDialog(string title)
        {
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = title
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                return dialog.FileName;
            }

            return null;
        }
    }
}
