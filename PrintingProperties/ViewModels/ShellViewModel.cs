using Caliburn.Micro;
using MvvmDialogs;
using MvvmDialogs.FrameworkDialogs.SaveFile;
using PrintingProperties.Library;
using PrintingProperties.Views;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace PrintingProperties.ViewModels
{
    public class ShellViewModel : Conductor<IScreen>.Collection.OneActive
    {
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
        PrinterSettings CurrentPrintSettings = new PrinterSettings();
        PrintDocument pd = new PrintDocument();
        PrintSettings objPrinterSetup = new PrintSettings();
        PrinterSetup CurrentSetup = new PrinterSetup();
        PrinterSetup DefaultSetup = new PrinterSetup();
        BinaryFormatter formatter = new BinaryFormatter();
        bool? nuInhibit = false;
        byte[]? Comparator1;
        byte[]? Comparator2;
        byte[]? DevModeArray;

        #region ComboBox Collections
        private List<InstalledPrinterModel>? _installedPrinters = new List<InstalledPrinterModel>();
        public List<InstalledPrinterModel>? InstalledPrinters
        {
            get { return _installedPrinters; }
            set
            {
                _installedPrinters = value;
                NotifyOfPropertyChange(() => InstalledPrinters);
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
                NotifyOfPropertyChange(() => SelectedPrinter);

                if (_selectedPrinter != null)
                {
                    pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = _selectedPrinter.Name;

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
            set { _storeTwoBytes = value; NotifyOfPropertyChange(() => StoreTwoBytes); }
        }
        public int StoreOneBytes
        {
            get { return _storeOneBytes; }
            set { _storeOneBytes = value; NotifyOfPropertyChange(() => StoreOneBytes); }
        }
        public string StoreOneStatus
        {
            get { return _storeOneStatus; }
            set { _storeOneStatus = value; NotifyOfPropertyChange(() => StoreOneStatus); }
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
        /// Clears the printer setup arraylist and puts 10 blank items in it
        /// </summary>
        private void InitialseCurrentSetup()
        {
            objPrinterSetup.alPrinterSetup.Clear();
            for (int i = 0; i <= 9; i++)
            {
                PrinterSetup temp = new PrinterSetup();
                temp = null;
                objPrinterSetup.alPrinterSetup.Add(temp);
            }
        }

        private void FindPrinters()
        {
            InstalledPrinters.Clear();

            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                InstalledPrinters.Add(new InstalledPrinterModel { Name = printer });
            }
            if (InstalledPrinters.Count > 0)
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
            try
            {
                nuInhibit = true;
                CurrentPrintSettings = pd.PrinterSettings;
                CurrentPrinterName = CurrentPrintSettings.PrinterName;

                //PaperSizes.Clear();
                //foreach (PaperSize PS in CurrentPrintSettings.PaperSizes)
                //{
                //    if (Enum.IsDefined(PS.Kind.GetType(), PS.Kind))
                //    {
                //        PaperSizes.Add(new PaperSizesModel { Name = PS });
                //    }
                //}
                //if (PaperSizes.Count > 0)
                //{
                //    SelectedPaperSize = PaperSizes[0];
                //    //cbxPaperSize.Text = CurrentPrintSettings.DefaultPageSettings.PaperSize.PaperName;
                //}
                //Trace.WriteLine($"Paper Size Count (After): {PaperSizes.Count}");

                //PrinterResolutions.Clear();
                //foreach (PrinterResolution PR in CurrentPrintSettings.PrinterResolutions)
                //{
                //    if (Enum.IsDefined(PR.Kind.GetType(), PR.Kind))
                //    {
                //        PrinterResolutions.Add(new PrinterResolutionModel { Name = PR });
                //    }
                //}
                //if (PrinterResolutions.Count > 0)
                //{
                //    SelectedPrinterResolution = PrinterResolutions[0];
                //    //cbxPrintQuality.Text = CurrentPrintSettings.DefaultPageSettings.PrinterResolution.Kind.ToString();
                //    //tbxHRes.Text = CurrentPrintSettings.DefaultPageSettings.PrinterResolution.X.ToString();
                //    //tbxVres.Text = CurrentPrintSettings.DefaultPageSettings.PrinterResolution.Y.ToString();
                //}
                //Trace.WriteLine($"Printer Resolution Count (After): {PrinterResolutions.Count}");

                //PaperSources.Clear();
                //foreach (PaperSource PSC in CurrentPrintSettings.PaperSources)
                //{
                //    if (Enum.IsDefined(PSC.Kind.GetType(), PSC.Kind))
                //    {
                //        PaperSources.Add(new PaperSourcesModel { Name = PSC });
                //    }
                //}
                //if (PaperSources.Count > 0)
                //{
                //    SelectedPaperSource = PaperSources[0];
                //}
                //else
                //{
                //    //SelectedPaperSource.Name = "None - (Default)";
                //}
                //Trace.WriteLine($"Paper Sources Count (After): {PaperSources.Count}");
                //if (CurrentPrintSettings.DefaultPageSettings.Landscape == true)
                //{
                //    rbLandscape.Checked = true;
                //}
                //if (CurrentPrintSettings.DefaultPageSettings.Landscape == false)
                //{
                //    rbPortrait.Checked = true;
                //}

                //DefaultSetup = SetPrinterDefault(CurrentPrintSettings);
                //MaximisePrintableArea();
                //nuInhibit = false;
                //UpdatePrintForm(CurrentPrintSettings, CurrentSetup);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                //nuInhibit = false;
            }
        }
               
        public void SaveToBin()
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
            if (success == true)
            {
                GetDevmode(CurrentPrintSettings, 1, settings.FileName);
            }
        }

        #region Store Methods
        public void Store1()
        {
            CurrentSetup = new PrinterSetup();
            if(StoreOneStatus == "Free")
            {
                Store(1);
                objPrinterSetup.alPrinterSetup[0] = CurrentSetup;
                StoreOneStatus = "In Use";
                StoreOneInUse = true;
         
            }
        }
        public void Store2()
        {
            CurrentSetup = new PrinterSetup();
            if (StoreTwoStatus == "Free")
            {
                Store(2);
                objPrinterSetup.alPrinterSetup[1] = CurrentSetup;
                StoreTwoStatus = "In Use";
                StoreTwoInUse = true;

            }
        }
        public void Store1Recall()
        {
            CurrentSetup = (PrinterSetup)objPrinterSetup.alPrinterSetup[0];
            if (CurrentSetup == null)
            {
                _dialogService.ShowMessageBox(this, "To Recall this location Set it first!");
            }
            else
            {
                Recall();                
            }
        }
        public void Store2Recall()
        {
            CurrentSetup = (PrinterSetup)objPrinterSetup.alPrinterSetup[1];
            if (CurrentSetup == null)
            {
                _dialogService.ShowMessageBox(this, "To Recall this location Set it first!");
            }
            else
            {
                Recall();
            }
        }
        public void Store1Clear()
        {
            if (StoreOneStatus == "In Use")
            {
                CurrentSetup = new PrinterSetup();
                objPrinterSetup.alPrinterSetup[0] = null;
                StoreOneStatus = "Free";
                StoreOneInUse = false;
                StoreOneBytes = 0;
            }
        }
        public void Store2Clear()
        {
            if (StoreTwoStatus == "In Use")
            {
                CurrentSetup = new PrinterSetup();
                objPrinterSetup.alPrinterSetup[1] = null;
                StoreTwoStatus = "Free";
                StoreTwoInUse = false;
                StoreTwoBytes = 0;
            }
        }
        private void Recall()
        {
            CurrentPrintSettings.PrinterName = CurrentSetup.NameOfPrinter;
            SetDevmode(CurrentPrintSettings, 2, "");
            
        }
        private void SetDevmode(PrinterSettings printerSettings, int mode, String Filename)//Grabs the data in arraylist and chucks it back into memory "Crank the suckers out"
        {
            ///int mode
            ///1 = Load devmode structure from file
            ///2 = Load devmode structure from arraylist
            IntPtr hDevMode = IntPtr.Zero;                        // a handle to our current DEVMODE
            IntPtr pDevMode = IntPtr.Zero;                          // a pointer to our current DEVMODE
            Byte[] Temparray;
            try
            {
                DevModeArray = CurrentSetup.Devmodearray;
                // Obtain the current DEVMODE position in memory
                hDevMode = printerSettings.GetHdevmode(printerSettings.DefaultPageSettings);

                // Obtain a lock on the handle and get an actual pointer so Windows won't move
                // it around while we're futzing with it
                pDevMode = GlobalLock(hDevMode);

                // Overwrite our current DEVMODE in memory with the one we saved.
                // They should be the same size since we haven't like upgraded the OS
                // or anything.


                if (mode == 1)  //Load devmode structure from file
                {
                    FileStream fs = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                    Temparray = new byte[fs.Length];
                    fs.Read(Temparray, 0, Temparray.Length);
                    fs.Close();
                    fs.Dispose();
                    for (int i = 0; i < Temparray.Length; ++i)
                    {
                        Marshal.WriteByte(pDevMode, i, Temparray[i]);
                    }
                }
                if (mode == 2)  //Load devmode structure from arraylist
                {
                    for (int i = 0; i < DevModeArray.Length; ++i)
                    {
                        Marshal.WriteByte(pDevMode, i, DevModeArray[i]);
                    }
                }
                // We're done futzing
                GlobalUnlock(hDevMode);

                // Tell our printer settings to use the one we just overwrote
                printerSettings.SetHdevmode(hDevMode);
                printerSettings.DefaultPageSettings.SetHdevmode(hDevMode);

                // It's copied to our printer settings, so we can free the OS-level one
                GlobalFree(hDevMode);
            }
            catch (Exception ex)
            {
                if (hDevMode != IntPtr.Zero)
                {
                    _dialogService.ShowMessageBox(this, $"SetDevmode Issue: {ex.Message}");
                    GlobalUnlock(hDevMode);
                    // And to boot, we don't need that DEVMODE anymore, either
                    GlobalFree(hDevMode);
                    hDevMode = IntPtr.Zero;
                }
            }

        }
        private void Store(int store)
        {
            GetDevmode(CurrentPrintSettings, 2, "");
            CurrentSetup.NameOfPrinter = CurrentPrintSettings.PrinterName;
            CurrentSetup.PaperSize = CurrentPrintSettings.DefaultPageSettings.PaperSize;
            CurrentSetup.PrintQuality = CurrentPrintSettings.DefaultPageSettings.PrinterResolution;
            CurrentSetup.PaperSource = CurrentPrintSettings.DefaultPageSettings.PaperSource;
            CurrentSetup.CanDuplex = CurrentPrintSettings.CanDuplex;
            CurrentSetup.DSided = CurrentPrintSettings.Duplex;
            CurrentSetup.Devmodearray = DevModeArray;

            if(store == 1)
            {
                StoreOneBytes = DevModeArray.Length;
            }
            else if(store == 2)
            {
                StoreTwoBytes = DevModeArray.Length;
            }
            //Trace.WriteLine($"Dev Mode Arry Length: {DevModeArray.Length}");
        }
        #endregion

        /// <summary>
        /// Get Application Window Handle
        /// </summary>
        /// <param name="lpClassName"></param>
        /// <param name="lpWindowName"></param>
        /// <returns></returns>
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        IntPtr hWnd = (IntPtr)FindWindow("Printer Settings to Bin converter", null);

        public void PrinterProperties()
        {
            IntPtr ipDevMode = CurrentPrintSettings.GetHdevmode();
            CurrentPrintSettings.DefaultPageSettings.CopyToHdevmode(ipDevMode);
            CurrentPrintSettings = OpenPrinterPropertiesDialog(CurrentPrintSettings);
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
                    _dialogService.ShowMessageBox(this, "Cant get size of devmode structure");
                    GlobalUnlock(hDevMode);
                    GlobalFree(hDevMode);
                    return;
                }
                DevModeArray = new byte[sizeNeeded];    //Copies the buffer into a byte array
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
                        DevModeArray[i] = (byte)(Marshal.ReadByte(pDevMode, i));    //Copies the array to an arraylist where it can be recalled
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

    public record InstalledPrinterModel
    {
        public string? Name { get; set; }
    }
}

