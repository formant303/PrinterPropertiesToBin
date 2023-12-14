using Caliburn.Micro;
using MvvmDialogs;
using MvvmDialogs.FrameworkDialogs.SaveFile;
using PrintingProperties.Library;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace PrintingProperties.ViewModels
{
    public class ShellViewModel : Conductor<IScreen>.Collection.OneActive
    {
        // DEVMODE is a structure in Windows that represents device-specific initialization and configuration data
        // for a printer or display device.The name "DEVMODE" stands for Device Mode.This structure is used with
        // various Windows API functions and structures related to printing and display settings.

        #region Platform Invokes
        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesW", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter,
            [MarshalAs(UnmanagedType.LPWStr)] string pDeviceName,
            IntPtr pDevModeOutput, IntPtr pDevModeInput, int fMode);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalFree(IntPtr handle);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalLock(IntPtr handle);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalUnlock(IntPtr handle);
        #endregion

        private readonly IDialogService _dialogService;
        PrinterSettings? _currentPrintSettings = new();
        PrintDocument? _pd = new PrintDocument();
        readonly PrintSettings? _objPrinterSetup = new();
        PrinterSetup? _currentSetup = new PrinterSetup();
        byte[]? _devModeArray;

        /// <summary>
        /// Get Application Window Name
        /// </summary>
        /// <param name="lpClassName"></param>
        /// <param name="lpWindowName"></param>
        /// <returns></returns>
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        IntPtr hWnd = (IntPtr)FindWindow("Printer Settings to Bin converter", string.Empty);

        #region Collections
        private List<InstalledPrinterModel>? _installedPrinters = new();
        public List<InstalledPrinterModel>? InstalledPrinters
        {
            get { return _installedPrinters; }
            set
            {
                _installedPrinters = value;
                NotifyOfPropertyChange();
            }
        }
        #endregion

        #region ComboBox Properties
        private InstalledPrinterModel? _selectedPrinter;
        public InstalledPrinterModel? SelectedPrinter
        {
            get { return _selectedPrinter; }
            set
            {
                _selectedPrinter = value;
                NotifyOfPropertyChange();

                if (_selectedPrinter?.Name != null)
                {
                    _pd = new PrintDocument();
                    _pd.PrinterSettings.PrinterName = _selectedPrinter.Name;

                    LoadPrinterSettings();
                }
            }
        }
        #endregion

        #region TextBox Properties     
        private string _storeOneStatus = "Free";
        private string _storeTwoStatus = "Free";
        private int _storeOneBytes;
        private int _storeTwoBytes;

        public int StoreTwoBytes
        {
            get { return _storeTwoBytes; }
            set { _storeTwoBytes = value; NotifyOfPropertyChange(); }
        }
        public int StoreOneBytes
        {
            get { return _storeOneBytes; }
            set { _storeOneBytes = value; NotifyOfPropertyChange(); }
        }
        public string StoreOneStatus
        {
            get { return _storeOneStatus; }
            set { _storeOneStatus = value; NotifyOfPropertyChange(); }
        }
        public string StoreTwoStatus
        {
            get { return _storeTwoStatus; }
            set { _storeTwoStatus = value; NotifyOfPropertyChange(() => StoreTwoStatus); }
        }
        #endregion

        #region General Properties
        private string? _currentPrinterName;
        public string? CurrentPrinterName
        {
            get { return _currentPrinterName; }
            set { _currentPrinterName = value; NotifyOfPropertyChange(() => CurrentPrinterName); }
        }

        private bool _isPrintPropetiesEnabled = true;
        public bool IsPrintPropertiesEnabled
        {
            get { return _isPrintPropetiesEnabled; }
            set { _isPrintPropetiesEnabled = value; NotifyOfPropertyChange(() => IsPrintPropertiesEnabled); }
        }

        private bool _storeOneInUse = false;
        public bool StoreOneInUse
        {
            get { return _storeOneInUse; }
            set { _storeOneInUse = value; NotifyOfPropertyChange(() => StoreOneInUse); }
        }

        private bool _storeTwoInUse = false;
        public bool StoreTwoInUse
        {
            get { return _storeTwoInUse; }
            set { _storeTwoInUse = value; NotifyOfPropertyChange(() => StoreTwoInUse); }
        }
        #endregion

        public ShellViewModel(IDialogService dialogService)
        {
            InitialseCurrentSetup();
            FindPrinters();
            _dialogService = dialogService;
        }

        /// <summary>
        /// Clears the printer setup arraylist and puts 2 blank items in it
        /// </summary>
        private void InitialseCurrentSetup()
        {
            _objPrinterSetup?.alPrinterSetup.Clear();
            for (int i = 0; i <= 2; i++)
            {
                PrinterSetup? temp = new();
                temp = null;
                _objPrinterSetup?.alPrinterSetup.Add(temp);
            }
        }

        private void FindPrinters()
        {
            InstalledPrinters?.Clear();

            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                InstalledPrinters?.Add(new InstalledPrinterModel { Name = printer });
            }
            if (InstalledPrinters?.Count > 0)
            {
                SelectedPrinter = InstalledPrinters[0];
            }

            LoadPrinterSettings();
        }

        /// <summary>
        /// Loads the settings currently in the printer for the selected printer
        /// </summary>
        private void LoadPrinterSettings()
        {
            if (_pd != null)
            {
                _currentPrintSettings = _pd.PrinterSettings;
                CurrentPrinterName = _currentPrintSettings.PrinterName;
            }
            else
            {
                _dialogService.ShowMessageBox(this, "PrintDocument (pd) was null");
                return;
            }
        }

        public void SaveToBin()
        {
            if (CurrentPrinterName != null)
            {
                var printerName = Regex.Replace(CurrentPrinterName, @"\s+", ""); //remove spaces

                var settings = new SaveFileDialogSettings
                {
                    Title = "Save Printer Settings File",
                    Filter = "Devmode File|*.Bin",
                    InitialDirectory = @"C:\Andy",
                    FileName = $"{printerName}<your description>"
                };

                bool? success = _dialogService.ShowSaveFileDialog(this, settings);

                if (success == true && _currentPrintSettings != null)
                {
                    GetDevmode(_currentPrintSettings, 1, settings.FileName);
                }
            }
        }

        #region UI Button Methods
        public void StoreSetup(int location)
        {
            _currentSetup = new PrinterSetup();

            if (_objPrinterSetup != null)
            {
                if (location == 1 && StoreOneStatus == "Free")
                {
                    Store(location);
                    _objPrinterSetup.alPrinterSetup[location - 1] = _currentSetup;
                    StoreOneStatus = "In Use";
                    StoreOneInUse = true;
                }
                else if (location == 2 && StoreTwoStatus == "Free")
                {
                    Store(location);
                    _objPrinterSetup.alPrinterSetup[location - 1] = _currentSetup;
                    StoreTwoStatus = "In Use";
                    StoreTwoInUse = true;
                }
            }
            else
            {
                _dialogService.ShowMessageBox(this, "PrintSettings (objPrinterSetup) was null");
                return;
            }

        }
        public void StoreRecall(int location)
        {
            _currentSetup = _objPrinterSetup?.alPrinterSetup[location - 1] as PrinterSetup;

            if (_currentSetup == null)
            {
                _dialogService.ShowMessageBox(this, $"To Recall location {location} Set it first.");
            }
            else
            {
                RecallFromArrayList();
            }
        }
        public void StoreClear(int location)
        {
            if (_objPrinterSetup != null)
            {
                if (location == 1 && StoreOneStatus == "In Use")
                {
                    _currentSetup = new PrinterSetup();
                    _objPrinterSetup.alPrinterSetup[0] = null;
                    StoreOneStatus = "Free";
                    StoreOneInUse = false;
                    StoreOneBytes = 0;
                }
                else if (location == 2 && StoreTwoStatus == "In Use")
                {
                    _currentSetup = new PrinterSetup();
                    _objPrinterSetup.alPrinterSetup[1] = null;
                    StoreTwoStatus = "Free";
                    StoreTwoInUse = false;
                    StoreTwoBytes = 0;
                }
            }
        }
        public void PrinterProperties()
        {
            try
            {
                if (SelectedPrinter != null)
                {
                    if (_currentPrintSettings != null)
                    {
                        IntPtr ipDevMode = _currentPrintSettings.GetHdevmode();
                        _currentPrintSettings.DefaultPageSettings.CopyToHdevmode(ipDevMode);
                        _currentPrintSettings = OpenPrinterPropertiesDialog(_currentPrintSettings);
                    }
                }
            }
            catch (Exception ex)
            {
                _dialogService.ShowMessageBox(this, $"Error setting CurrentPrintSettings. It's possible some of the settings you are trying to save are not supported.\\n{ex}");
                return;
            }
        }
        #endregion

        /// <summary>
        /// Recalls and applies stored printer setup
        /// </summary>
        private void RecallFromArrayList()
        {
            if (_currentPrintSettings != null && _currentSetup != null)
            {
                _currentPrintSettings.PrinterName = _currentSetup.NameOfPrinter;
                SetDevmode(_currentPrintSettings, 2, "");
            }

        }
        private void SetDevmode(PrinterSettings printerSettings, int mode, String Filename)
        {
            ///int mode
            ///1 = Load devmode structure from file
            ///2 = Load devmode structure from arraylist
            IntPtr hDevMode = IntPtr.Zero;                        // a handle to our current DEVMODE
            IntPtr pDevMode = IntPtr.Zero;                        // a pointer to our current DEVMODE
            Byte[] devModeData;

            try
            {
                if (_currentSetup != null)
                {
                    _devModeArray = _currentSetup.Devmodearray;
                    // Obtain the current DEVMODE position in memory
                    hDevMode = printerSettings.GetHdevmode(printerSettings.DefaultPageSettings);

                    // Obtain a lock on the handle and get an actual pointer so Windows won't move it around
                    pDevMode = GlobalLock(hDevMode);

                    // Overwrite our current DEVMODE in memory with the one we saved.
                    // They should be the same size since we haven't like upgraded the OS
                    // or anything.

                    if (mode == 1) // Load DEVMODE structure from file
                    {
                        using (FileStream fs = new FileStream(Filename, FileMode.Open, FileAccess.Read))
                        {
                            devModeData = new byte[fs.Length];
                            fs.Read(devModeData, 0, devModeData.Length);
                        }
                    }
                    else if (mode == 2 && _devModeArray != null) // Load DEVMODE structure from array
                    {
                        devModeData = _devModeArray;
                    }
                    else
                    {
                        // Handle unsupported mode
                        return;
                    }

                    for (int i = 0; i < devModeData.Length; ++i)
                    {
                        Marshal.WriteByte(pDevMode, i, devModeData[i]);
                    }

                    GlobalUnlock(hDevMode);

                    // Tell our printer settings to use the one we just overwrote
                    printerSettings.SetHdevmode(hDevMode);
                    printerSettings.DefaultPageSettings.SetHdevmode(hDevMode);

                    // It's copied to our printer settings, so we can free the OS-level one
                    GlobalFree(hDevMode);
                }
            }
            catch (Exception ex)
            {
                if (hDevMode != IntPtr.Zero)
                {
                    _dialogService.ShowMessageBox(this, $"SetDevmode Issue: {ex.Message}");
                    GlobalUnlock(hDevMode);

                    GlobalFree(hDevMode);
                    hDevMode = IntPtr.Zero;
                }
            }

        }
        private void Store(int store)
        {
            if (_currentPrintSettings != null && _currentSetup != null)
            {
                GetDevmode(_currentPrintSettings, 2, "");
                _currentSetup.NameOfPrinter = _currentPrintSettings.PrinterName;
                _currentSetup.PaperSize = _currentPrintSettings.DefaultPageSettings.PaperSize;
                _currentSetup.PrintQuality = _currentPrintSettings.DefaultPageSettings.PrinterResolution;
                _currentSetup.PaperSource = _currentPrintSettings.DefaultPageSettings.PaperSource;
                _currentSetup.CanDuplex = _currentPrintSettings.CanDuplex;
                _currentSetup.DSided = _currentPrintSettings.Duplex;
                _currentSetup.Devmodearray = _devModeArray;

                if (store == 1)
                {
                    StoreOneBytes = _devModeArray!.Length;
                }
                else if (store == 2)
                {
                    StoreTwoBytes = _devModeArray!.Length;
                }
            }
        }

        /// <summary>
        /// Shows the printer settings dialogue that comes with the printer driver
        /// </summary>
        /// <param name="printerSettings"></param>
        /// <returns></returns>
        private PrinterSettings OpenPrinterPropertiesDialog(PrinterSettings printerSettings)
        {
            IsPrintPropertiesEnabled = false;

            IntPtr hDevMode = IntPtr.Zero;
            IntPtr devModeData = IntPtr.Zero;
            IntPtr hPrinter = IntPtr.Zero;
            String pName = printerSettings.PrinterName;
            try
            {
                hDevMode = printerSettings.GetHdevmode(printerSettings.DefaultPageSettings);
                IntPtr pDevMode = GlobalLock(hDevMode);
                int sizeNeeded = DocumentProperties(hWnd, IntPtr.Zero, pName, devModeData, pDevMode, 0);//get needed size and allocate memory 

                if (sizeNeeded < 0)
                {
                    _dialogService.ShowMessageBox(this, "PrinterSettings Issue: Cant get size of devmode structure");
                    Marshal.FreeHGlobal(devModeData);
                    Marshal.FreeHGlobal(hDevMode);
                    devModeData = IntPtr.Zero;
                    hDevMode = IntPtr.Zero;
                    return printerSettings;
                }
                devModeData = Marshal.AllocHGlobal(sizeNeeded);

                //show the native dialog 
                int returncode = DocumentProperties(hWnd, IntPtr.Zero, pName, devModeData, pDevMode, 14);
                if (returncode < 0) //Failure to display native dialogue
                {
                    _dialogService.ShowMessageBox(this, "PrinterSettings Issue: Got devmode, but the dialogue got stuck");
                    Marshal.FreeHGlobal(devModeData);
                    Marshal.FreeHGlobal(hDevMode);
                    devModeData = IntPtr.Zero;
                    hDevMode = IntPtr.Zero;

                    IsPrintPropertiesEnabled = true;

                    return printerSettings;
                }
                if (returncode == 2) //User clicked "Cancel"
                {
                    GlobalUnlock(hDevMode);//unlocks the memory
                    if (hDevMode != IntPtr.Zero)
                    {
                        Marshal.FreeHGlobal(hDevMode); //Frees the memory
                        hDevMode = IntPtr.Zero;
                    }
                    if (devModeData != IntPtr.Zero)
                    {
                        GlobalFree(devModeData);
                        devModeData = IntPtr.Zero;
                    }
                    IsPrintPropertiesEnabled = true;
                }
                IsPrintPropertiesEnabled = true;
                GlobalUnlock(hDevMode);//unlocks the memory

                if (hDevMode != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(hDevMode); //Frees the memory
                    hDevMode = IntPtr.Zero;
                }
                if (devModeData != IntPtr.Zero)
                {
                    printerSettings.SetHdevmode(devModeData);
                    printerSettings.DefaultPageSettings.SetHdevmode(devModeData);
                    GlobalFree(devModeData);
                    devModeData = IntPtr.Zero;
                }
            }
            catch (Exception ex)
            {
                IsPrintPropertiesEnabled = true;
                _dialogService.ShowMessageBox(this, "PrinterSettings Issue: An error has occurred, caught and chucked back\n" + ex.Message);
            }
            finally
            {
                if (hDevMode != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(hDevMode);
                }
                if (devModeData != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(devModeData);
                }
            }

            //IsPrintPropertiesEnabled = true;
            return printerSettings;
        }

        /// <summary>
        /// Grabs the devmode data in memory and stores in arraylist
        /// </summary>
        /// <param name="printerSettings"></param>
        /// <param name="mode"></param>
        /// <param name="Filename"></param>
        private void GetDevmode(PrinterSettings printerSettings, int mode, String Filename)
        {
            IntPtr hDevMode = IntPtr.Zero; // handle to the DEVMODE
            IntPtr pDevMode = IntPtr.Zero; // pointer to the DEVMODE
            IntPtr hwnd = hWnd;
            try
            {
                // Get a handle to a DEVMODE for the default printer settings
                hDevMode = printerSettings.GetHdevmode(printerSettings.DefaultPageSettings);

                // Obtain a lock on the handle and get an actual pointer so Windows won't
                // move it around while we're futzing with it
                pDevMode = GlobalLock(hDevMode);
                int sizeNeeded = DocumentProperties(hwnd, IntPtr.Zero, printerSettings.PrinterName, IntPtr.Zero, pDevMode, 0);
                if (sizeNeeded <= 0)
                {
                    _dialogService.ShowMessageBox(this, "Can't get size of devmode structure");
                    GlobalUnlock(hDevMode);
                    GlobalFree(hDevMode);
                    return;
                }
                _devModeArray = new byte[sizeNeeded];    //Copies the buffer into a byte array
                if (mode == 1)  //Save devmode structure to file
                {
                    FileStream fs = new FileStream(Filename, FileMode.Create);
                    for (int i = 0; i < sizeNeeded; ++i)
                    {
                        fs.WriteByte(Marshal.ReadByte(pDevMode, i));
                    }
                    fs.Close();
                    fs.Dispose();
                }
                if (mode == 2)  //Save devmode structure to Byte array and arraylist
                {
                    for (int i = 0; i < sizeNeeded; ++i)
                    {
                        _devModeArray[i] = (byte)(Marshal.ReadByte(pDevMode, i));    //Copies the array to an arraylist where it can be recalled
                    }
                }

                // Unlock the handle, we're done futzing around with memory
                GlobalUnlock(hDevMode);
                // And to boot, we don't need that DEVMODE anymore, either
                GlobalFree(hDevMode);
                hDevMode = IntPtr.Zero;
            }
            catch (Exception ex)
            {
                if (hDevMode != IntPtr.Zero)
                {
                    _dialogService.ShowMessageBox(this, $"{ex.Message}");
                    GlobalUnlock(hDevMode);
                    // And to boot, we don't need that DEVMODE anymore, either
                    GlobalFree(hDevMode);
                    hDevMode = IntPtr.Zero;
                }
            }
        }
    }

    public class InstalledPrinterModel
    {
        public string? Name { get; set; }
    }
}