using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Data;
using System.Globalization;
using MahApps.Metro.Controls.Dialogs;
using System.Threading;

namespace ExcelImport
{
    public class UploadStateToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var state = (UploadState)value;
            if (state == UploadState.ReadyToUpload||state==UploadState.Success||state==UploadState.Failed)
                return true;
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class DelegateCommand : ICommand
    {
        Action _execute;
        Func<bool> _canExecute;
        public DelegateCommand(Action execute)
        : this(execute, null)
        {

        }
        public DelegateCommand(Action execute, Func<bool> canExecute)
        {
            _execute = execute;
            _canExecute = canExecute;
        }
        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute();
        }
        public void Execute(object parameter)
        {
            _execute?.Invoke();
        }
        public event EventHandler CanExecuteChanged;
        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
    public class DelegateCommand<T> : ICommand
    {
        private readonly Predicate<T> _canExecute;
        private readonly Action<T> _execute;
        public DelegateCommand(Action<T> execute)
            : this(execute, null)
        {
        }
        public DelegateCommand(Action<T> execute, Predicate<T> canExecute)
        {
            _execute = execute;
            _canExecute = canExecute;
        }
        public bool CanExecute(object parameter)
        {
            if (_canExecute == null)
                return true;

            return _canExecute((parameter == null) ? default(T) : (T)Convert.ChangeType(parameter, typeof(T)));
        }
        public void Execute(object parameter)
        {
            _execute((parameter == null) ? default(T) : (T)Convert.ChangeType(parameter, typeof(T)));
        }
        public event EventHandler CanExecuteChanged;
        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
    public enum UploadState { NotReady, ReadyToUpload, Uploading, Success, Failed }
    public class UploaderPresenter : DependencyObject
    {
        public UploaderPresenter():base()
        {
            _uploader = new BakgroundInfoUploader();
            _uploader.RunWorkerCompleted += Uploader_RunWorkerCompleted;
            _uploader.ProgressChanged += Uploader_ProgressChanged;
            SelectedFileCommand = new DelegateCommand(ExecuteSelectFileCommand);
            UploadCommand = new DelegateCommand(ExecuteUploadCommand);
            ExportCommand = new DelegateCommand(ExecuteExportCommand);
        }

        private void Uploader_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if(_controller!=null)
            {
                _controller.SetProgress(e.ProgressPercentage / 100.0);
            }
        }
        private void Uploader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                UploadingState = UploadState.Success;
                _controller.CloseAsync();
                Window.ShowMessageAsync("Upload Success","");
            }
            else
            {
                UploadingState = UploadState.Failed;
                _controller.CloseAsync();
                Window.ShowMessageAsync("Upload Fail",e.Error.Message);
            }
        }

        public static readonly DependencyProperty UploadFileNameProperty = DependencyProperty.Register("UploadFileName", typeof(string), typeof(UploaderPresenter));
        public static readonly DependencyProperty UploadStateProperty = DependencyProperty.Register("UploadingState", typeof(UploadState), typeof(UploaderPresenter));
        public static readonly DependencyProperty ReportedProgressProperty = DependencyProperty.Register("ReportedProgress", typeof(int), typeof(UploaderPresenter));
        public ICommand SelectedFileCommand { get; private set; }
        public ICommand UploadCommand { get; private set; }
        public ICommand ExportCommand { get; private set; }
        public string UploadFileName
        {
            get { return (string)GetValue(UploadFileNameProperty); }
            set { SetValue(UploadFileNameProperty, value); }
        }
        public UploadState UploadingState
        {
            get { return (UploadState)GetValue(UploadStateProperty); }
            set { SetValue(UploadStateProperty, value);}
        }
        public int ReportedProgress
        {
            get { return (int)GetValue(ReportedProgressProperty); }
            set { SetValue(ReportedProgressProperty, value); }
        }
        public MainWindow Window { get; set; }
        private void ExecuteSelectFileCommand()
        {  
            System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
            System.Windows.Forms.DialogResult r = dialog.ShowDialog();
            if(r== System.Windows.Forms.DialogResult.OK)
            {
                UploadFileName = dialog.FileName;
                _async_uploader.Path = dialog.FileName;
                UploadingState = UploadState.ReadyToUpload;
            }
        }
        private async void ExecuteUploadCommand()
        {
            _controller = await Window.ShowProgressAsync("Please wait","Uploading......",true);
            var wapper = new Wrapper(_controller);
            UploadingState = UploadState.Uploading;
            CancellationTokenSource cts = new CancellationTokenSource();
            _controller.Canceled += (sender, e) => { cts.Cancel(); };
            try
            {
                var result = await _async_uploader.UploadExistSheetsAsync(cts.Token, wapper);
                if (result != 0)
                {
                    UploadingState = UploadState.Success;
                }
                else
                {
                    throw new Exception("Don't have any sheet, make sure follow naming rule!");
                }
                await _controller.CloseAsync();
                await Window.ShowMessageAsync("Upload Success", $"{result} sheet(s) is(are) uploaded");
            }
            catch (OperationCanceledException)
            {
                await _controller.CloseAsync();
            }
            catch (Exception ex)
            {
                await _controller.CloseAsync();
                await Window.ShowMessageAsync("Error",ex.Message);
            }
        }
        private async void ExecuteExportCommand()
        {
            _controller = await Window.ShowProgressAsync("Please wait", "Exporting......", true);
            var wapper = new Wrapper(_controller);
            UploadingState = UploadState.Uploading;
            CancellationTokenSource cts = new CancellationTokenSource();
            _controller.Canceled += (sender, e) => { cts.Cancel(); };
            try
            {
                await _async_uploader.ExportAsync(cts.Token, wapper);
                await _controller.CloseAsync();
            }
            catch (OperationCanceledException)
            {
                await _controller.CloseAsync();
            }
            catch (Exception ex)
            {
                await _controller.CloseAsync();
                await Window.ShowMessageAsync("Error", ex.Message);
            }
        }
        private AsyncInfoUploader _async_uploader = new AsyncInfoUploader();
        private BakgroundInfoUploader _uploader;
        private ProgressDialogController _controller;
    }
    class Wrapper : IProgress<Double>
    {
        private ProgressDialogController _controller;
        public Wrapper(ProgressDialogController controller)
        {
            _controller = controller;
        }

        public void Report(double value)
        {
            _controller.SetProgress(value);
        }
    }
}
