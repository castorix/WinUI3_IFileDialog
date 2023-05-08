using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

using static WinUI3_IFileDialog.CFileDialog;
using GlobalStructures;
using static GlobalStructures.GlobalTools;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace WinUI3_IFileDialog
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public delegate int SUBCLASSPROC(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, uint dwRefData);

        [DllImport("Comctl32.dll", SetLastError = true)]
        public static extern bool SetWindowSubclass(IntPtr hWnd, SUBCLASSPROC pfnSubclass, uint uIdSubclass, uint dwRefData);

        [DllImport("Comctl32.dll", SetLastError = true)]
        public static extern int DefSubclassProc(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam);

        COMDLG_FILTERSPEC[] cdfs = new COMDLG_FILTERSPEC[]
        {
            new COMDLG_FILTERSPEC("All files (*.*)", "*.*"),
            new COMDLG_FILTERSPEC("JPG Joint Photographic Experts Group (*.jpg)", "*.jpg"),
            new COMDLG_FILTERSPEC("PNG Portable Network Graphics (*.png)", "*.png"),
            new COMDLG_FILTERSPEC("GIF Graphics Interchange Format (*.gif)", "*.gif"),
            new COMDLG_FILTERSPEC("BMP Windows Bitmap (*.bmp)", "*.bmp"),
            new COMDLG_FILTERSPEC("TIF Tagged Image File Format (*.tif)", "*.tif")
        };

        private IntPtr hWndMain = IntPtr.Zero;
        private Microsoft.UI.Windowing.AppWindow _apw;
        private SUBCLASSPROC SubClassDelegate;
        public const int ID_PUSHBUTTON = 601;
        ObservableCollection<File> files = new ObservableCollection<File>();

        public MainWindow()
        {
            this.InitializeComponent();
            hWndMain = WinRT.Interop.WindowNative.GetWindowHandle(this);
            Microsoft.UI.WindowId myWndId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWndMain);
            _apw = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(myWndId);
            _apw.Resize(new Windows.Graphics.SizeInt32(1000, 530));
            _apw.Move(new Windows.Graphics.PointInt32(400, 300));
            this.Title = "WinUI 3 - Test IFileDialog";
            // For custom Button
            SubClassDelegate = new SUBCLASSPROC(WindowSubClass);
            bool bRet = SetWindowSubclass(hWndMain, SubClassDelegate, 0, 0);
        }

        private async void btnChooseFiles_Click(object sender, RoutedEventArgs e)
        {
            // Documents
            string sFolder = "shell:::{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}";
            FOS nOptions = (FOS.FOS_FORCESHOWHIDDEN | FOS.FOS_ALLOWMULTISELECT | FOS.FOS_NODEREFERENCELINKS);
            string sFile = null;
            object oFileNames = new System.Collections.ObjectModel.Collection<string>();
            IShellItem si = null;
            IShellItemArray sia = null;
            CFileDialog cfd = new CFileDialog(nOptions, "Choose files", sFolder);
            HRESULT hr = cfd.SetFileTypes(cdfs, 1);
            if ((cfd.ShowDialog(hWndMain, true) == DialogResult.OK))
            {
                files.Clear();
                Windows.Storage.StorageFile file = null;
                Windows.Storage.FileProperties.StorageItemThumbnail iconThumbnail = null;               
                if (cfd.FileName != null)
                {
                    sFile = cfd.FileName;
                    si = cfd.ShellItem;
                }
                if (cfd.FileNames != null)
                {
                    sia = cfd.ShellItemArray;
                    oFileNames = cfd.FileNames;
                    int nCount = ((System.Collections.ObjectModel.Collection<CFileDialog.FileNameType>)oFileNames).Count;
                    for (int n = 0; n < nCount; n++)
                    {
                        file = null;                       
                        var fn = ((System.Collections.ObjectModel.Collection<CFileDialog.FileNameType>)oFileNames).ElementAt(n);
                        try
                        {
                            file = await Windows.Storage.StorageFile.GetFileFromPathAsync(fn.FileName);
                        }
                        catch (Exception ex)
                        {
                            // For example for Hidden files
                            // Message = "Accès refusé. (0x80070005 (E_ACCESSDENIED))"
                            file = null;
                        }
                        if (file != null)
                        {
                            iconThumbnail = await file.GetThumbnailAsync(Windows.Storage.FileProperties.ThumbnailMode.SingleItem, 32);
                        }
                        else
                        {
                            iconThumbnail = null;
                        }
                        var bi = new BitmapImage();
                        if (iconThumbnail != null)
                            bi.SetSource(iconThumbnail);
                        files.Add(new File(fn.FileName, fn.IsFolder?"Folder":"File", bi));
                    }
                }
            }
        }

        private async void btnChooseDirectory_Click(object sender, RoutedEventArgs e)
        {
            // This PC
            string sFolder = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
            FOS nOptions = (FOS.FOS_FORCESHOWHIDDEN | FOS.FOS_PICKFOLDERS | FOS.FOS_NODEREFERENCELINKS);
            string sFile = null;
            object oFileNames = new System.Collections.ObjectModel.Collection<string>();
            IShellItem si = null;
            CFileDialog cfd = new CFileDialog(nOptions, "Choose directory", sFolder);
            if ((cfd.ShowDialog(hWndMain, true) == DialogResult.OK))
            {
                files.Clear();
                Windows.Storage.StorageFolder folder = null;
                Windows.Storage.FileProperties.StorageItemThumbnail iconThumbnail = null;
                if (cfd.FileName != null)
                {
                    sFile = cfd.FileName;
                    si = cfd.ShellItem;
                }
                if (cfd.FileNames != null)
                {
                    oFileNames = cfd.FileNames;
                    var fn = ((System.Collections.ObjectModel.Collection<CFileDialog.FileNameType>)oFileNames).ElementAt(0);                   
                    try
                    {
                        folder = await Windows.Storage.StorageFolder.GetFolderFromPathAsync(fn.FileName);                       
                    }
                    catch (Exception ex)
                    {
                        folder = null;
                    }
                    if (folder != null)
                    {
                        iconThumbnail = await folder.GetThumbnailAsync(Windows.Storage.FileProperties.ThumbnailMode.SingleItem, 32);
                    }
                    else
                    {
                        iconThumbnail = null;
                    }
                    var bi = new BitmapImage();
                    if (iconThumbnail != null)
                        bi.SetSource(iconThumbnail);
                    files.Add(new File(fn.FileName, fn.IsFolder ? "Folder" : "File", bi));
                }
            }
        }

        private async void btnChooseFilesDirectories_Click(object sender, RoutedEventArgs e)
        {
            // This PC
            string sFolder = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
            FOS nOptions = (FOS.FOS_FORCESHOWHIDDEN | FOS.FOS_ALLOWMULTISELECT | FOS.FOS_NODEREFERENCELINKS);
            string sFile = null;
            object oFileNames = new System.Collections.ObjectModel.Collection<string>();
            IShellItem si = null;
            IShellItemArray sia = null;
            CFileDialog cfd = new CFileDialog(nOptions, "Choose files & directories", sFolder);
            cfd.AddPushButton(ID_PUSHBUTTON, " Return selected items ");
            if ((cfd.ShowDialog(hWndMain, true) == DialogResult.OK))
            {
                files.Clear();
                Windows.Storage.StorageFile file = null;
                Windows.Storage.StorageFolder folder = null;
                Windows.Storage.FileProperties.StorageItemThumbnail iconThumbnail = null;
                if (cfd.FileName != null)
                {
                    sFile = cfd.FileName;
                    si = cfd.ShellItem;
                }
                if (cfd.FileNames != null)
                {
                    sia = cfd.ShellItemArray;
                    oFileNames = cfd.FileNames;
                    int nCount = ((System.Collections.ObjectModel.Collection<CFileDialog.FileNameType>)oFileNames).Count;
                    for (int n = 0; n < nCount; n++)
                    {
                        file = null;
                        folder = null;
                        var fn = ((System.Collections.ObjectModel.Collection<CFileDialog.FileNameType>)oFileNames).ElementAt(n);
                        try
                        {
                            if (fn.IsFolder)
                                folder = await Windows.Storage.StorageFolder.GetFolderFromPathAsync(fn.FileName);
                            else
                                file = await Windows.Storage.StorageFile.GetFileFromPathAsync(fn.FileName);
                        }
                        catch (Exception ex)
                        {
                            file = null;
                            folder = null;
                        }
                        if (file != null || folder != null)
                        {
                            if (file != null)
                                iconThumbnail = await file.GetThumbnailAsync(Windows.Storage.FileProperties.ThumbnailMode.SingleItem, 32);
                            if (folder != null)
                                iconThumbnail = await folder.GetThumbnailAsync(Windows.Storage.FileProperties.ThumbnailMode.SingleItem, 32);
                        }
                        else
                        {
                            iconThumbnail = null;
                        }
                        var bi = new BitmapImage();
                        if (iconThumbnail != null)
                            bi.SetSource(iconThumbnail);
                        files.Add(new File(fn.FileName, fn.IsFolder ? "Folder" : "File", bi));
                    }
                }
            }
        }

        private async void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            // Documents
            string sFolder = "shell:::{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}";
            FOS nOptions = (FOS.FOS_FORCESHOWHIDDEN | FOS.FOS_NODEREFERENCELINKS);
            string sFile = null;
            object oFileNames = new System.Collections.ObjectModel.Collection<string>();
            IShellItem si = null;
            CFileDialog cfd = new CFileDialog(nOptions, "Save file", sFolder);
            HRESULT hr = cfd.SetFileTypes(cdfs, 1);
            hr = cfd.SetDefaultExtension("jpg");
            if ((cfd.ShowDialog(hWndMain, false) == DialogResult.OK))
            {
                files.Clear();
                Windows.Storage.StorageFile file = null;
                Windows.Storage.FileProperties.StorageItemThumbnail iconThumbnail = null;
                if (cfd.FileName != null)
                {
                    sFile = cfd.FileName;
                    si = cfd.ShellItem;
                }
                if (cfd.FileNames != null)
                {
                    oFileNames = cfd.FileNames;
                    var fn = ((System.Collections.ObjectModel.Collection<CFileDialog.FileNameType>)oFileNames).ElementAt(0);
                    try
                    {
                        file = await Windows.Storage.StorageFile.GetFileFromPathAsync(fn.FileName);
                    }
                    catch (Exception ex)
                    {
                        file = null;
                    }
                    if (file != null)
                    {
                        iconThumbnail = await file.GetThumbnailAsync(Windows.Storage.FileProperties.ThumbnailMode.SingleItem, 32);
                    }
                    else
                    {
                        iconThumbnail = null;
                    }
                    var bi = new BitmapImage();
                    if (iconThumbnail != null)
                        bi.SetSource(iconThumbnail);
                    files.Add(new File(fn.FileName, fn.IsFolder ? "Folder" : "File", bi));
                }
            }            
        }

        private int WindowSubClass(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, uint dwRefData)
        {
            switch (uMsg)
            {
                case WM_APP_FILEDIALOG:
                    {
                        if (wParam == (IntPtr)ID_PUSHBUTTON)
                        {
                            CDialogEventSink cdes = Marshal.GetObjectForIUnknown(lParam) as CDialogEventSink;
                            CFileDialog pcfd = cdes.parentCFD;
                            IFileDialogCustomize pfdc = pcfd.m_pfdc;
                            IShellItemArray psia = null;
                            IntPtr ppvOut = IntPtr.Zero;
                            Guid SID_SFolderView = new Guid("cde725b0-ccc9-4519-917e-325d72fab4ce");
                            Guid IID_IForlderView = typeof(IFolderView2).GUID;
                            IntPtr pUnknown = Marshal.GetIUnknownForObject(pfdc);
                            HRESULT hr = IUnknown_QueryService(pUnknown, ref SID_SFolderView, ref IID_IForlderView, out ppvOut);
                            if ((hr == HRESULT.S_OK))
                            {
                                IFolderView2 pfv2 = (IFolderView2)Marshal.GetObjectForIUnknown(ppvOut);
                                hr = pfv2.GetSelection(true, ref psia);
                                if ((hr == HRESULT.S_OK))
                                {
                                    int nCount = 0;
                                    hr = psia.GetCount(ref nCount);
                                    for (int i = 0; i <= nCount - 1; i++)
                                    {
                                        IShellItem si = null;
                                        hr = psia.GetItemAt(i, ref si);
                                        System.Text.StringBuilder sbResult = new System.Text.StringBuilder(MAX_PATH);
                                        hr = si.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, ref sbResult);
                                        if ((hr == HRESULT.S_OK))
                                        {
                                            bool bFolder = false;
                                            uint nAttributes = 0;
                                            hr = si.GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, ref nAttributes);
                                            if ((nAttributes & SFGAO_FOLDER) == SFGAO_FOLDER & (nAttributes & SFGAO_FILESYSTEM) == SFGAO_FILESYSTEM & (nAttributes & SFGAO_STREAM) != SFGAO_STREAM)
                                                bFolder = true;
                                            else
                                                bFolder = false;
                                            pcfd.FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
                                        }
                                        else
                                        {
                                            hr = si.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, ref sbResult);
                                            if ((hr == HRESULT.S_OK))
                                            {
                                                bool bFolder = false;
                                                uint nAttributes = 0;
                                                hr = si.GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, ref nAttributes);
                                                if ((nAttributes & SFGAO_FOLDER) == SFGAO_FOLDER & (nAttributes & SFGAO_FILESYSTEM) == SFGAO_FILESYSTEM & (nAttributes & SFGAO_STREAM) != SFGAO_STREAM)
                                                    bFolder = true;
                                                else
                                                    bFolder = false;
                                                pcfd.FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
                                            }
                                        }                                        
                                    }
                                    Marshal.ReleaseComObject(psia);
                                }
                                Marshal.ReleaseComObject(pfv2);

                                IFileDialog pfd;
                                pfd = (IFileDialog)pfdc;
                                pfd.Close((int)HRESULT.S_OK);
                            }                            
                        }                       
                    }
                    break;
            }
            return DefSubclassProc(hWnd, uMsg, wParam, lParam);
        }     
    }

    public class File
    {
        #region Properties
        public string Name { get; set; }
        public string Type { get; set; }
        public BitmapImage biFileThumbnail { get; set; }
        #endregion

        public File(string sName, string sType, BitmapImage biThumbnail)
        {
            Name = sName;
            Type = sType;
            biFileThumbnail = biThumbnail;
        }
    }
}
