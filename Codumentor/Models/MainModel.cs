using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Codumentor.Models
{
    class MainModel : BindableBase
    {
        private ObservableCollection<string> _filePaths = new ObservableCollection<string>();

        public ObservableCollection<string> FilePaths
        {
            get => _filePaths;
            private set => SetProperty(ref _filePaths, value);
        }

        public void AddFile(string filePath)
        {
            if (!FilePaths.Any(p => p.Equals(filePath, StringComparison.OrdinalIgnoreCase)))
            {
                FilePaths.Add(filePath);
            }
        }

        public void RemoveFile(string filePath)
        {
            if (filePath != null)
            {
                FilePaths.Remove(filePath);
            }
        }

        public void MoveUp(string filePath)
        {
            if (filePath == null) return;

            int oldIndex = FilePaths.IndexOf(filePath);
            if (oldIndex > 0)
            {
                FilePaths.Move(oldIndex, oldIndex - 1);
            }
        }

        public void MoveDown(string filePath)
        {
            if (filePath == null) return;

            int oldIndex = FilePaths.IndexOf(filePath);
            if (oldIndex < FilePaths.Count - 1)
            {
                FilePaths.Move(oldIndex, oldIndex + 1);
            }
        }
    }
}
