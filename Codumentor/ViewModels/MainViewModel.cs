using Codumentor.Models;
using Codumentor.Services;
using System.Collections.ObjectModel;
using System.Windows;
using MaterialDesignThemes.Wpf;
using System.ComponentModel;
using System.Drawing.Text;
using System.Collections.Specialized;

namespace Codumentor.ViewModels
{
    public class MainViewModel : BindableBase
    {
        private readonly CodeToWordExporter _exporter;
        private readonly MainModel _model = new();
        public ObservableCollection<string> FilePaths => _model.FilePaths;

        private string _folderPath;
        public string FolderPath
        {
            get => _folderPath;
            set => SetProperty(ref _folderPath, value);
        }

        private string _fileName;
        public string FileName
        {
            get => _fileName;
            set => SetProperty(ref _fileName, value);
        }

        private string _fileExtension;
        public string FileExtension
        {
            get => _fileExtension;
            set => SetProperty(ref _fileExtension, value);
        }

        private bool _isMenuOpen;
        public bool IsMenuOpen
        {
            get => _isMenuOpen;
            set => SetProperty(ref _isMenuOpen, value);
        }

        public DelegateCommand<DragEventArgs> DropCommand { get; }
        public DelegateCommand<string> MoveUpCommand { get; }
        public DelegateCommand<string> MoveDownCommand { get; }
        public DelegateCommand<string> RemoveCommand { get; }

        public AsyncDelegateCommand ExportToWordCommand { get; }
        public DelegateCommand OpenFolderDialogCommand { get; }
        public DelegateCommand ToggleMenuCommand { get; }

        public ISnackbarMessageQueue SnackMessageQueue { get; }

        public MainViewModel()
        {
            if (DesignerProperties.GetIsInDesignMode(new DependencyObject()))
            {
                FilePaths.Add(@"C:\Test\file1.cs");
                FilePaths.Add(@"C:\Test\file2.md");
                FilePaths.Add(@"D:\Projects\example.txt");
            }

            FileName = "exported_code";
            FileExtension = "pdf";

            DropCommand = new DelegateCommand<DragEventArgs>(OnDrop);
            ExportToWordCommand = new AsyncDelegateCommand(ExportToWordAsync);

            MoveUpCommand = new DelegateCommand<string>(
                filePath => _model.MoveUp(filePath),
                filePath => FilePaths.IndexOf(filePath) > 0);

            MoveDownCommand = new DelegateCommand<string>(
                filePath => _model.MoveDown(filePath),
                filePath => FilePaths.IndexOf(filePath) < FilePaths.Count - 1);

            FilePaths.CollectionChanged += OnFilesChanged;

            RemoveCommand = new DelegateCommand<string>(filePath => _model.RemoveFile(filePath));

            SnackMessageQueue = new SnackbarMessageQueue(TimeSpan.FromSeconds(2));
            OpenFolderDialogCommand = new DelegateCommand(OpenFolderDialog);

            ToggleMenuCommand = new DelegateCommand(() =>
            {
                IsMenuOpen = !IsMenuOpen;
            });

            _exporter = new CodeToWordExporter();
        }

        private void OnFilesChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            MoveDownCommand.RaiseCanExecuteChanged();
            MoveUpCommand.RaiseCanExecuteChanged();
        }

        private void OnDrop(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (var file in files)
                {
                    _model.AddFile(file);
                }
            }
        }

        private async Task ExportToWordAsync()
        {
            if (string.IsNullOrEmpty(FolderPath))
            {
                ShowMessage("Выберите папку для сохранения");
                return;
            }
            try
            {
                ShowMessage("Сохранение файла...");

                string outputPath = $@"{FolderPath}\{FileName}.{FileExtension}";
                switch (FileExtension)
                {
                    case "pdf":
                        await _exporter.ExportCodeToDocumentAsync(FilePaths, outputPath, true);
                        break;
                    case "docx":
                        await _exporter.ExportCodeToDocumentAsync(FilePaths, outputPath, false);
                        break;
                }

                ShowMessage("Файл сохранён!");
            }
            catch (Exception ex)
            {
                ShowMessage("Ошибка: " + ex.Message);
            }
        }

        private void OpenFolderDialog()
        {
            FolderPath = FolderDialogService.OpenFolderDialog("Выберите папку для сохранения");
        }

        public void ShowMessage(string message)
        {
            SnackMessageQueue.Enqueue(message);
        }
    }
}
