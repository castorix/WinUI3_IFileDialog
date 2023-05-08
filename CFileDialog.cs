using Microsoft.UI.Xaml.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using GlobalStructures;
using static GlobalStructures.GlobalTools;
using Microsoft.UI.Xaml.Media.Imaging;
using System.Collections.ObjectModel;

namespace WinUI3_IFileDialog
{
    public class CFileDialog
    {
        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        public const int WM_APP = 0x8000;
        public const int WM_APP_FILEDIALOG = WM_APP + 1;

        public CFileDialog(FOS nOptions = 0, string sTitle = null, string sFolder = null)
        {
            Options = nOptions;
            Title = sTitle;
            Folder = sFolder;
            //FileNames = new System.Collections.ObjectModel.Collection<string>();
            FileNames = new System.Collections.ObjectModel.Collection<FileNameType>();
        }

        public FOS Options { get; set; }
        public string Title { get; set; }
        public string Folder { get; set; }
        public string FileName { get; set; }   
        public IShellItem ShellItem { get; set; }
        public IShellItemArray ShellItemArray { get; set; }

        COMDLG_FILTERSPEC[] m_rgFilterSpec;
        uint m_nFileTypeIndex = 0;
        string m_sDefaultExtension = "";
        ObservableCollection<PushButton> m_Buttons = new ObservableCollection<PushButton>();
        IntPtr m_hWndOwner = IntPtr.Zero;

        public class PushButton
        {
            #region Properties
            public uint Id { get; set; }
            public string Label { get; set; }
            #endregion

            public PushButton(uint nId, string sLabel)
            {               
                Id = nId;
                Label = sLabel;
            }
        }

        //public System.Collections.ObjectModel.Collection<string> FileNames { get; set; }
        public System.Collections.ObjectModel.Collection<FileNameType> FileNames { get; set; }
        public class FileNameType
        {
            public string FileName { get; set; }
            public bool IsFolder { get; set; }
        }

        private CDialogEventSink eventSink;
        public IFileDialogCustomize m_pfdc = null;

        public DialogResult ShowDialog(IntPtr hWndOwner, bool bOpen)
        {
            HRESULT hr = HRESULT.E_FAIL;
            m_hWndOwner = hWndOwner;
            IFileOpenDialog fod = null;
            IFileSaveDialog fsd = null;            
            if (bOpen)
            {
                fod = (IFileOpenDialog)new FileOpenDialog();
                FOS nOptions = 0;

                hr = fod.GetOptions(out nOptions);
                nOptions = nOptions | Options;
                hr = fod.SetOptions(nOptions);

                if (!string.IsNullOrEmpty(Title))
                    hr = fod.SetTitle(Title);
               
                if (m_rgFilterSpec != null)
                {
                    uint nFileTypes = (uint)m_rgFilterSpec.Count();
                    fod.SetFileTypes(nFileTypes, m_rgFilterSpec);
                    fod.SetFileTypeIndex(m_nFileTypeIndex);
                }
            }
            else
            {
                fsd = (IFileSaveDialog)new FileSaveDialog();
                FOS nOptions = 0;

                hr = fsd.GetOptions(out nOptions);
                nOptions = nOptions | Options;
                hr = fsd.SetOptions(nOptions);

                if (hr == HRESULT.S_OK)
                {
                    if (!string.IsNullOrEmpty(Title))
                        hr = fsd.SetTitle(Title);

                    if (m_rgFilterSpec != null)
                    {
                        uint nFileTypes = (uint)m_rgFilterSpec.Count();
                        fsd.SetFileTypes(nFileTypes, m_rgFilterSpec);
                        fsd.SetFileTypeIndex(m_nFileTypeIndex);
                    }
                }
                else
                    return DialogResult.Abort;
            }
            try
            {
                if (!string.IsNullOrEmpty(Folder))
                {
                    Guid GUID_IShellItem = typeof(IShellItem).GUID;
                    IntPtr psi = IntPtr.Zero;
                    hr = SHCreateItemFromParsingName(Folder, IntPtr.Zero, ref GUID_IShellItem, ref psi);
                    if ((hr == HRESULT.S_OK))
                    {
                        IShellItem si = (IShellItem)Marshal.GetObjectForIUnknown(psi);
                        if (bOpen)
                            hr = fod.SetFolder(si);
                        else
                            hr = fsd.SetFolder(si);
                        Marshal.ReleaseComObject(si);
                    }
                }

                uint nCookie;
                eventSink = new CDialogEventSink(this);
                if (bOpen)
                    hr = fod.Advise(eventSink, out nCookie);
                else
                    hr = fsd.Advise(eventSink, out nCookie);
                eventSink.Cookie = nCookie;
              
                if (bOpen)
                    m_pfdc = (IFileDialogCustomize)fod;
                else
                    m_pfdc = (IFileDialogCustomize)fsd;
                
                foreach (var pb in m_Buttons)
                {
                    hr = m_pfdc.AddPushButton(pb.Id, pb.Label);                  
                }

                if (bOpen)
                {
                    hr = fod.Show(hWndOwner);
                    if ((hr == HRESULT.S_OK))
                    {
                        IShellItem pShellItemResult = null;
                        System.Text.StringBuilder sbResult = new System.Text.StringBuilder(MAX_PATH);
                        hr = fod.GetResult(out pShellItemResult);
                        if ((hr == HRESULT.S_OK))
                        {
                            ShellItem = pShellItemResult;
                            hr = pShellItemResult.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, ref sbResult);
                            if ((hr == HRESULT.S_OK))
                            {
                                FileName = sbResult.ToString();

                                bool bFolder = false;
                                uint nAttributes = 0;
                                hr = ShellItem.GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, ref nAttributes);
                                if ((nAttributes & SFGAO_FOLDER) == SFGAO_FOLDER & (nAttributes & SFGAO_FILESYSTEM) == SFGAO_FILESYSTEM & (nAttributes & SFGAO_STREAM) != SFGAO_STREAM)
                                    bFolder = true;
                                else
                                    bFolder = false;
                                FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
                            }
                            else
                            {
                                hr = pShellItemResult.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, ref sbResult);
                                if ((hr == HRESULT.S_OK))
                                {
                                    FileName = sbResult.ToString();

                                    bool bFolder = false;
                                    uint nAttributes = 0;
                                    hr = ShellItem.GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, ref nAttributes);
                                    if ((nAttributes & SFGAO_FOLDER) == SFGAO_FOLDER & (nAttributes & SFGAO_FILESYSTEM) == SFGAO_FILESYSTEM & (nAttributes & SFGAO_STREAM) != SFGAO_STREAM)
                                        bFolder = true;
                                    else
                                        bFolder = false;
                                    FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
                                }
                            }
                            //Marshal.ReleaseComObject(pShellItemResult);
                        }
                        else if (hr == HRESULT.E_UNEXPECTED)
                        {
                            IShellItemArray pShellItemArrayResult = null;
                            hr = fod.GetResults(out pShellItemArrayResult);
                            if ((hr == HRESULT.S_OK))
                            {
                                ShellItemArray = pShellItemArrayResult;
                                int nCount = 0;
                                hr = pShellItemArrayResult.GetCount(ref nCount);
                                for (int i = 0; i <= nCount - 1; i++)
                                {
                                    IShellItem si = null;
                                    hr = pShellItemArrayResult.GetItemAt(i, ref si);
                                    if (hr == HRESULT.S_OK)
                                    {
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
                                            FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
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
                                                FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (hr == HRESULT.E_CANCELLED)
                        return DialogResult.Cancel;

                    return DialogResult.OK;
                }
                else
                {
                    hr = fsd.SetDefaultExtension(m_sDefaultExtension);
                    hr = fsd.Show(hWndOwner);
                    if ((hr == HRESULT.S_OK))
                    {
                        IShellItem pShellItemResult = null;
                        System.Text.StringBuilder sbResult = new System.Text.StringBuilder(MAX_PATH);
                        hr = fsd.GetResult(out pShellItemResult);
                        if ((hr == HRESULT.S_OK))
                        {
                            ShellItem = pShellItemResult;
                            hr = pShellItemResult.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, ref sbResult);
                            if ((hr == HRESULT.S_OK))
                            {
                                FileName = sbResult.ToString();

                                bool bFolder = false;
                                uint nAttributes = 0;
                                hr = ShellItem.GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, ref nAttributes);
                                if ((nAttributes & SFGAO_FOLDER) == SFGAO_FOLDER & (nAttributes & SFGAO_FILESYSTEM) == SFGAO_FILESYSTEM & (nAttributes & SFGAO_STREAM) != SFGAO_STREAM)
                                    bFolder = true;
                                else
                                    bFolder = false;
                                FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
                            }
                            else
                            {
                                hr = pShellItemResult.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, ref sbResult);
                                if ((hr == HRESULT.S_OK))
                                {
                                    FileName = sbResult.ToString();

                                    bool bFolder = false;
                                    uint nAttributes = 0;
                                    hr = ShellItem.GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, ref nAttributes);
                                    if ((nAttributes & SFGAO_FOLDER) == SFGAO_FOLDER & (nAttributes & SFGAO_FILESYSTEM) == SFGAO_FILESYSTEM & (nAttributes & SFGAO_STREAM) != SFGAO_STREAM)
                                        bFolder = true;
                                    else
                                        bFolder = false;
                                    FileNames.Add(new FileNameType() { FileName = sbResult.ToString(), IsFolder = bFolder });
                                }
                            }
                            //Marshal.ReleaseComObject(pShellItemResult);
                        }
                        else if (hr == HRESULT.E_UNEXPECTED)
                        {

                        }
                    }
                    else if (hr == HRESULT.E_CANCELLED)
                        return DialogResult.Cancel;

                    return DialogResult.OK;
                }
            }
            finally
            {
                if (bOpen)
                    Marshal.ReleaseComObject(fod);
                else
                    Marshal.ReleaseComObject(fsd);
            }
        }

        public HRESULT SetFileTypes([In][MarshalAs(UnmanagedType.LPArray)] COMDLG_FILTERSPEC[] rgFilterSpec, uint nFileTypeIndex)
        {
            HRESULT hr = HRESULT.S_OK;
            m_rgFilterSpec = rgFilterSpec;
            m_nFileTypeIndex = nFileTypeIndex;
            return hr;
        }

        public HRESULT SetDefaultExtension(string sDefaultExtension)
        {
            HRESULT hr = HRESULT.S_OK;
            m_sDefaultExtension = sDefaultExtension;
            return hr;
        }        

        public void AddPushButton(uint nID, string sLabel)
        {
            m_Buttons.Add(new PushButton(nID, sLabel));            
        }

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern HRESULT SHCreateItemFromParsingName(string pszPath, IntPtr pbc, ref Guid riid, ref IntPtr ppv);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern HRESULT SHGetKnownFolderItem(ref Guid rfid, KNOWN_FOLDER_FLAG flags, IntPtr hToken, ref Guid riid, ref IntPtr ppv);

        public enum KNOWN_FOLDER_FLAG : uint
        {
            KF_FLAG_DEFAULT = 0x00000000,
            KF_FLAG_FORCE_APP_DATA_REDIRECTION = 0x00080000,
            KF_FLAG_RETURN_FILTER_REDIRECTION_TARGET = 0x00040000,
            KF_FLAG_FORCE_PACKAGE_REDIRECTION = 0x00020000,
            KF_FLAG_NO_PACKAGE_REDIRECTION = 0x00010000,
            KF_FLAG_FORCE_APPCONTAINER_REDIRECTION = 0x00020000,
            KF_FLAG_NO_APPCONTAINER_REDIRECTION = 0x00010000,
            KF_FLAG_CREATE = 0x00008000,
            KF_FLAG_DONT_VERIFY = 0x00004000,
            KF_FLAG_DONT_UNEXPAND = 0x00002000,
            KF_FLAG_NO_ALIAS = 0x00001000,
            KF_FLAG_INIT = 0x00000800,
            KF_FLAG_DEFAULT_PATH = 0x00000400,
            KF_FLAG_NOT_PARENT_RELATIVE = 0x00000200,
            KF_FLAG_SIMPLE_IDLIST = 0x00000100,
            KF_FLAG_ALIAS_ONLY = 0x80000000,
        }

        public const int SFGAO_DROPTARGET = 0x100; // Objects are drop target
        public const int SFGAO_FOLDER = 0x20000000; // Support BindToObject(IID_IShellFolder)
        public const int SFGAO_FILESYSTEM = 0x40000000; // Is a win32 file system Object (file/folder/root)
        public const int SFGAO_STREAM = 0x400000; // Supports BindToObject(IID_IStream)

        public class CDialogEventSink : IFileDialogEvents, IFileDialogControlEvents
        {
            public CFileDialog parentCFD;

            public CDialogEventSink(CFileDialog fileDialog)
            {
                this.parentCFD = fileDialog;
            }

            public uint Cookie { get; set; }

            public HRESULT OnFileOk([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd)
            {
                return HRESULT.S_OK;
            }

            public HRESULT OnFolderChanging([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd, [In][MarshalAs(UnmanagedType.Interface)] IShellItem psiFolder)
            {
                //System.Text.StringBuilder sbResult = new System.Text.StringBuilder(MAX_PATH);
                return HRESULT.S_OK;
            }

            // If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
            public void OnFolderChange([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd)
            {
            }

            public void OnSelectionChange([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd)
            {
            }

            public void OnShareViolation([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd, [In][MarshalAs(UnmanagedType.Interface)] IShellItem psi, out FDE_SHAREVIOLATION_RESPONSE pResponse)
            {
                //throw new NotImplementedException();
                pResponse = FDE_SHAREVIOLATION_RESPONSE.FDESVR_DEFAULT;
            }

            public void OnTypeChange([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd)
            {
                //throw new NotImplementedException();
            }

            public void OnOverwrite([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd, [In][MarshalAs(UnmanagedType.Interface)] IShellItem psi, out FDE_OVERWRITE_RESPONSE pResponse)
            {
                //throw new NotImplementedException();
                pResponse = FDE_OVERWRITE_RESPONSE.FDEOR_DEFAULT;
            }

            public HRESULT OnItemSelected([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl, [In] int dwIDItem)
            {
                //throw new NotImplementedException();
                return HRESULT.S_OK;
            }

            // If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
            public void OnButtonClicked([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl)
            {
                IntPtr pClassPtr = Marshal.GetIUnknownForObject(this);
                SendMessage(parentCFD.m_hWndOwner, WM_APP_FILEDIALOG, (IntPtr)dwIDCtl, (IntPtr)pClassPtr);
                return;
            }

            public HRESULT OnCheckButtonToggled([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl, [In] bool bChecked)
            {
                //throw new NotImplementedException();
                return HRESULT.S_OK;
            }

            public HRESULT OnControlActivating([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl)
            {
                //throw new NotImplementedException();
                return HRESULT.S_OK;
            }
        }

        public enum FOS : uint
        {
            FOS_OVERWRITEPROMPT = 0x2,
            FOS_STRICTFILETYPES = 0x4,
            FOS_NOCHANGEDIR = 0x8,
            FOS_PICKFOLDERS = 0x20,
            FOS_FORCEFILESYSTEM = 0x40,
            FOS_ALLNONSTORAGEITEMS = 0x80,
            FOS_NOVALIDATE = 0x100,
            FOS_ALLOWMULTISELECT = 0x200,
            FOS_PATHMUSTEXIST = 0x800,
            FOS_FILEMUSTEXIST = 0x1000,
            FOS_CREATEPROMPT = 0x2000,
            FOS_SHAREAWARE = 0x4000,
            FOS_NOREADONLYRETURN = 0x8000,
            FOS_NOTESTFILECREATE = 0x10000,
            FOS_HIDEMRUPLACES = 0x20000,
            FOS_HIDEPINNEDPLACES = 0x40000,
            FOS_NODEREFERENCELINKS = 0x100000,
            FOS_OKBUTTONNEEDSINTERACTION = 0x200000, // Only enable the OK button If the user has done something In the view.        
            FOS_DONTADDTORECENT = 0x2000000,
            FOS_FORCESHOWHIDDEN = 0x10000000,
            FOS_DEFAULTNOMINIMODE = 0x20000000, // (Not used In Win7)        
            FOS_FORCEPREVIEWPANEON = 0x40000000,
            FOS_SUPPORTSTREAMABLEITEMS = 0x80000000 // Indicates the caller will use BHID_Stream To open contents, no need To download the file
        }

        [ComImport()]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
        public interface IShellItem
        {
            HRESULT BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, ref IntPtr ppv);
            HRESULT GetParent(ref IShellItem ppsi);
            [PreserveSig]
            HRESULT GetDisplayName(SIGDN sigdnName, ref System.Text.StringBuilder ppszName);
            HRESULT GetAttributes(uint sfgaoMask, ref uint psfgaoAttribs);
            HRESULT Compare(IShellItem psi, uint hint, ref int piOrder);
        }

        public enum SIGDN : uint
        {
            SIGDN_NORMALDISPLAY = 0x0,
            SIGDN_PARENTRELATIVEPARSING = 0x80018001,
            SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,
            SIGDN_PARENTRELATIVEEDITING = 0x80031001,
            SIGDN_DESKTOPABSOLUTEEDITING = 0x8004C000,
            SIGDN_FILESYSPATH = 0x80058000,
            SIGDN_URL = 0x80068000,
            SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8007C001,
            SIGDN_PARENTRELATIVE = 0x80080001
        }

        [ComImport()]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("b63ea76d-1f85-456f-a19c-48159efa858b")]
        public interface IShellItemArray
        {
            HRESULT BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, ref IntPtr ppvOut);
            HRESULT GetPropertyStore(GETPROPERTYSTOREFLAGS flags, ref Guid riid, ref IntPtr ppv);
            HRESULT GetPropertyDescriptionList(PROPERTYKEY keyType, ref Guid riid, ref IntPtr ppv);
            // Function GetAttributes(AttribFlags As SIATTRIBFLAGS, sfgaoMask As SFGAOF, ByRef psfgaoAttribs As SFGAOF) As HRESULT
            HRESULT GetAttributes(SIATTRIBFLAGS AttribFlags, int sfgaoMask, ref int psfgaoAttribs);
            HRESULT GetCount(ref int pdwNumItems);
            HRESULT GetItemAt(int dwIndex, ref IShellItem ppsi);

            // Function EnumItems(ByRef ppenumShellItems As IEnumShellItems) As HRESULT
            HRESULT EnumItems(ref IntPtr ppenumShellItems);
        }

        public enum GETPROPERTYSTOREFLAGS
        {
            GPS_DEFAULT = 0,
            GPS_HANDLERPROPERTIESONLY = 0x1,
            GPS_READWRITE = 0x2,
            GPS_TEMPORARY = 0x4,
            GPS_FASTPROPERTIESONLY = 0x8,
            GPS_OPENSLOWITEM = 0x10,
            GPS_DELAYCREATION = 0x20,
            GPS_BESTEFFORT = 0x40,
            GPS_NO_OPLOCK = 0x80,
            GPS_PREFERQUERYPROPERTIES = 0x100,
            GPS_EXTRINSICPROPERTIES = 0x200,
            GPS_EXTRINSICPROPERTIESONLY = 0x400,
            GPS_MASK_VALID = 0x7FF
        }

        public enum SIATTRIBFLAGS
        {
            SIATTRIBFLAGS_AND = 0x1,
            SIATTRIBFLAGS_OR = 0x2,
            SIATTRIBFLAGS_APPCOMPAT = 0x3,
            SIATTRIBFLAGS_MASK = 0x3,
            SIATTRIBFLAGS_ALLITEMS = 0x4000
        }

        [ComImport]
        [Guid("b4db1657-70d7-485e-8e3e-6fcb5a5c1802")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IModalWindow
        {
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall, MethodCodeType = System.Runtime.CompilerServices.MethodCodeType.Runtime)]
            [PreserveSig]
            uint Show([In] IntPtr hwndOwner);
        }

        [ComImport]
        [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileDialog : IModalWindow
        {
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall, MethodCodeType = System.Runtime.CompilerServices.MethodCodeType.Runtime)]
            [PreserveSig]
            new HRESULT Show([In] IntPtr parent);
            HRESULT SetFileTypes([In] uint cFileTypes, [In][MarshalAs(UnmanagedType.LPArray)] COMDLG_FILTERSPEC[] rgFilterSpec);
            HRESULT SetFileTypeIndex([In] uint iFileType);
            HRESULT GetFileTypeIndex(out uint piFileType);
            HRESULT Advise([In][MarshalAs(UnmanagedType.Interface)] IFileDialogEvents pfde, out uint pdwCookie);
            HRESULT Unadvise([In] uint dwCookie);
            HRESULT SetOptions([In] FOS fos);
            HRESULT GetOptions(out FOS pfos);
            HRESULT SetDefaultFolder([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            HRESULT SetFolder([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            HRESULT GetFolder([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            HRESULT GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            HRESULT SetFileName([In][MarshalAs(UnmanagedType.LPWStr)] string pszName);
            HRESULT GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            HRESULT SetTitle([In][MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            HRESULT SetOkButtonLabel([In][MarshalAs(UnmanagedType.LPWStr)] string pszText);
            HRESULT SetFileNameLabel([In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            [PreserveSig]
            HRESULT GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            HRESULT AddPlace([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi, FDAP fdap);
            HRESULT SetDefaultExtension([In][MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            HRESULT Close([MarshalAs(UnmanagedType.Error)] int hr);
            HRESULT SetClientGuid([In] ref Guid guid);
            HRESULT ClearClientData();
            HRESULT SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);
        }

        public enum FDAP
        {
            FDAP_BOTTOM,
            FDAP_TOP
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct COMDLG_FILTERSPEC
        {
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszSpec;

            public COMDLG_FILTERSPEC(string pszName, string pszSpec)
            {
                this.pszName = pszName;
                this.pszSpec = pszSpec;
            }
        }

        [ComImport]
        [Guid("973510DB-7D7F-452B-8975-74A85828D354")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileDialogEvents
        {
            [PreserveSig]
            HRESULT OnFileOk([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);

            [PreserveSig]
            HRESULT OnFolderChanging([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd, [In][MarshalAs(UnmanagedType.Interface)] IShellItem psiFolder);

            // If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
            void OnFolderChange([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);
            void OnSelectionChange([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);
            void OnShareViolation([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd, [In][MarshalAs(UnmanagedType.Interface)] IShellItem psi, out FDE_SHAREVIOLATION_RESPONSE pResponse);
            void OnTypeChange([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);
            void OnOverwrite([In][MarshalAs(UnmanagedType.Interface)] IFileDialog pfd, [In][MarshalAs(UnmanagedType.Interface)] IShellItem psi, out FDE_OVERWRITE_RESPONSE pResponse);
        }

        public enum FDE_SHAREVIOLATION_RESPONSE
        {
            FDESVR_DEFAULT = 0x0,
            FDESVR_ACCEPT = 0x1,
            FDESVR_REFUSE = 0x2
        }

        public enum FDE_OVERWRITE_RESPONSE
        {
            FDEOR_DEFAULT = 0x0,
            FDEOR_ACCEPT = 0x1,
            FDEOR_REFUSE = 0x2
        }

        [ComImport]
        [Guid("36116642-D713-4b97-9B83-7484A9D00433")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileDialogControlEvents
        {
            HRESULT OnItemSelected([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl, [In] int dwIDItem);

            // If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
            void OnButtonClicked([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl);
            HRESULT OnCheckButtonToggled([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl, [In] bool bChecked);
            HRESULT OnControlActivating([In][MarshalAs(UnmanagedType.Interface)] IFileDialogCustomize pfdc, [In] int dwIDCtl);
        }

        [ComImport]
        [Guid("e6fdd21a-163f-4975-9c8c-a69f1ba37034")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileDialogCustomize
        {
            HRESULT EnableOpenDropDown([In] int dwIDCtl);
            HRESULT AddMenu([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            HRESULT AddPushButton([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            HRESULT AddComboBox([In] uint dwIDCtl);
            HRESULT AddRadioButtonList([In] uint dwIDCtl);
            HRESULT AddCheckButton([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel, [In] bool bChecked);
            HRESULT AddEditBox([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszText);
            HRESULT AddSeparator([In] uint dwIDCtl);
            HRESULT AddText([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszText);
            HRESULT SetControlLabel([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            HRESULT GetControlState([In] uint dwIDCtl, out CDCONTROLSTATEF pdwState);
            HRESULT SetControlState([In] uint dwIDCtl, [In] CDCONTROLSTATEF dwState);
            HRESULT GetEditBoxText([In] uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] out string ppszText);
            HRESULT SetEditBoxText([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszText);
            HRESULT GetCheckButtonState([In] uint dwIDCtl, out bool pbChecked);
            HRESULT SetCheckButtonState([In] uint dwIDCtl, [In] bool bChecked);
            HRESULT AddControlItem([In] uint dwIDCtl, [In] uint dwIDItem, [In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            HRESULT RemoveControlItem([In] uint dwIDCtl, [In] uint dwIDItem);
            HRESULT RemoveAllControlItems([In] uint dwIDCtl);
            HRESULT GetControlItemState([In] uint dwIDCtl, [In] uint dwIDItem, out CDCONTROLSTATEF pdwState);
            HRESULT SetControlItemState([In] uint dwIDCtl, [In] uint dwIDItem, [In] CDCONTROLSTATEF dwState);
            HRESULT GetSelectedControlItem([In] uint dwIDCtl, out uint pdwIDItem);
            HRESULT SetSelectedControlItem([In] uint dwIDCtl, [In] uint dwIDItem);
            HRESULT StartVisualGroup([In] uint dwIDCtl, [In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            HRESULT EndVisualGroup();
            HRESULT MakeProminent([In] uint dwIDCtl);
            HRESULT SetControlItemText([In] uint dwIDCtl, [In] uint dwIDItem, [In] string pszLabel);
        }

        public enum CDCONTROLSTATEF
        {
            CDCS_INACTIVE = 0,
            CDCS_ENABLED = 0x1,
            CDCS_VISIBLE = 0x2,
            CDCS_ENABLEDVISIBLE = 0x3
        }

        [ComImport]
        [Guid("d57c7288-d4ad-4768-be02-9d969532d960")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IFileOpenDialog : IFileDialog
        {
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall, MethodCodeType = System.Runtime.CompilerServices.MethodCodeType.Runtime)]
            [PreserveSig]
            new HRESULT Show([In] IntPtr hwndOwner);
            new HRESULT SetFileTypes([In] uint cFileTypes, [In][MarshalAs(UnmanagedType.LPArray)] COMDLG_FILTERSPEC[] rgFilterSpec);
            new HRESULT SetFileTypeIndex([In] uint iFileType);
            new HRESULT GetFileTypeIndex(out uint piFileType);
            new HRESULT Advise([In][MarshalAs(UnmanagedType.Interface)] IFileDialogEvents pfde, out uint pdwCookie);
            new HRESULT Unadvise([In] uint dwCookie);
            new HRESULT SetOptions([In] FOS fos);
            new HRESULT GetOptions(out FOS pfos);
            new HRESULT SetDefaultFolder([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            new HRESULT SetFolder([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            new HRESULT GetFolder([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            new HRESULT GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            new HRESULT SetFileName([In][MarshalAs(UnmanagedType.LPWStr)] string pszName);
            new HRESULT GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            new HRESULT SetTitle([In][MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            new HRESULT SetOkButtonLabel([In][MarshalAs(UnmanagedType.LPWStr)] string pszText);
            new HRESULT SetFileNameLabel([In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            [PreserveSig]
            new HRESULT GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            new HRESULT AddPlace([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi, FDAP fdap);
            new HRESULT SetDefaultExtension([In][MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            new HRESULT Close([MarshalAs(UnmanagedType.Error)] int hr);
            new HRESULT SetClientGuid([In] ref Guid guid);
            new HRESULT ClearClientData();
            new HRESULT SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);
            [PreserveSig]
            HRESULT GetResults([MarshalAs(UnmanagedType.Interface)] out IShellItemArray ppenum);
            HRESULT GetSelectedItems([MarshalAs(UnmanagedType.Interface)] out IShellItemArray ppsai);
        }

        [ComImport]
        [Guid("84bccd23-5fde-4cdb-aea4-af64b83d78ab")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IFileSaveDialog : IFileDialog
        {
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall, MethodCodeType = System.Runtime.CompilerServices.MethodCodeType.Runtime)]
            [PreserveSig]
            new HRESULT Show([In] IntPtr hwndOwner);
            new HRESULT SetFileTypes([In] uint cFileTypes, [In][MarshalAs(UnmanagedType.LPArray)] COMDLG_FILTERSPEC[] rgFilterSpec);
            new HRESULT SetFileTypeIndex([In] uint iFileType);
            new HRESULT GetFileTypeIndex(out uint piFileType);
            new HRESULT Advise([In][MarshalAs(UnmanagedType.Interface)] IFileDialogEvents pfde, out uint pdwCookie);
            new HRESULT Unadvise([In] uint dwCookie);
            [PreserveSig]
            new HRESULT SetOptions([In] FOS fos);
            new HRESULT GetOptions(out FOS pfos);
            new HRESULT SetDefaultFolder([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            new HRESULT SetFolder([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            new HRESULT GetFolder([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            new HRESULT GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            new HRESULT SetFileName([In][MarshalAs(UnmanagedType.LPWStr)] string pszName);
            new HRESULT GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            new HRESULT SetTitle([In][MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            new HRESULT SetOkButtonLabel([In][MarshalAs(UnmanagedType.LPWStr)] string pszText);
            new HRESULT SetFileNameLabel([In][MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            [PreserveSig]
            new HRESULT GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            new HRESULT AddPlace([In][MarshalAs(UnmanagedType.Interface)] IShellItem psi, FDAP fdap);
            new HRESULT SetDefaultExtension([In][MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            new HRESULT Close([MarshalAs(UnmanagedType.Error)] int hr);
            new HRESULT SetClientGuid([In] ref Guid guid);
            new HRESULT ClearClientData();
            new HRESULT SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);
            [PreserveSig]

            HRESULT SetSaveAsItem(IShellItem psi);
            HRESULT SetProperties(IPropertyStore pStore);
            HRESULT SetCollectedProperties(IPropertyDescriptionList pList, bool fAppendDefault);
            HRESULT GetProperties(out IPropertyStore ppStore);
            HRESULT ApplyProperties(IShellItem psi, IPropertyStore pStore, IntPtr hwnd, IFileOperationProgressSink pSink);
        }

        [ComImport, Guid("1f9fc1d0-c39b-4b26-817f-011967d3440e"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyDescriptionList
        {
            HRESULT GetCount(out uint pcElem);
            HRESULT GetAt(uint iElem, ref Guid riid, out IntPtr ppv);
        }

        [ComImport]
        [Guid("04b0f1a7-9490-44bc-96e1-4296a31252e2")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileOperationProgressSink
        {
            void StartOperations();
            void FinishOperations(HRESULT hrResult);
            void PreRenameItem(uint dwFlags, IShellItem psiItem, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName);
            void PostRenameItem(uint dwFlags, IShellItem psiItem, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName, HRESULT hrRename, IShellItem psiNewlyCreated);
            void PreMoveItem(uint dwFlags, IShellItem psiItem, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName);
            void PostMoveItem(uint dwFlags, IShellItem psiItem, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName, HRESULT hrMove, IShellItem psiNewlyCreated);
            void PreCopyItem(uint dwFlags, IShellItem psiItem, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName);
            void PostCopyItem(uint dwFlags, IShellItem psiItem, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName, HRESULT hrCopy, IShellItem psiNewlyCreated);
            void PreDeleteItem(uint dwFlags, IShellItem psiItem);
            void PostDeleteItem(uint dwFlags, IShellItem psiItem, HRESULT hrDelete, IShellItem psiNewlyCreated);
            void PreNewItem(uint dwFlags, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName);
            void PostNewItem(uint dwFlags, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName, [MarshalAs(UnmanagedType.LPWStr)] string pszTemplateName, uint dwFileAttributes, HRESULT hrNew, IShellItem psiNewItem);
            void UpdateProgress(uint iWorkTotal, uint iWorkSoFar);
            void ResetTimer();
            void PauseTimer();
            void ResumeTimer();
        }

        [ComImport]
        [Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        private class FileOpenDialog
        {
        }

        [ComImport]
        [Guid("C0B4E2F3-BA21-4773-8DBA-335EC946EB8B")]
        private class FileSaveDialog
        {
        }

        [ComImport]
        [Guid("cde725b0-ccc9-4519-917e-325d72fab4ce")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFolderView
        {
            HRESULT GetCurrentViewMode(ref uint pViewMode);
            HRESULT SetCurrentViewMode(uint ViewMode);
            HRESULT GetFolder(ref Guid riid, ref IntPtr ppv);
            HRESULT Item(int iItemIndex, ref IntPtr ppidl);
            HRESULT ItemCount(uint uFlags, ref int pcItems);
            HRESULT Items(uint uFlags, ref Guid riid, ref IntPtr ppv);
            HRESULT GetSelectionMarkedItem(ref int piItem);
            HRESULT GetFocusedItem(ref int piItem);
            HRESULT GetItemPosition(IntPtr pidl, ref Windows.Foundation.Point ppt);
            HRESULT GetSpacing(ref Windows.Foundation.Point ppt);
            HRESULT GetDefaultSpacing(ref Windows.Foundation.Point ppt);
            HRESULT GetAutoArrange();
            HRESULT SelectItem(int iItem, int dwFlags);
            HRESULT SelectAndPositionItems(uint cidl, IntPtr apidl, Windows.Foundation.Point apt, int dwFlags);
        }

        [ComImport]
        [Guid("1af3a467-214f-4298-908e-06b03e0b39f9")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFolderView2 : IFolderView
        {
            new HRESULT GetCurrentViewMode(ref uint pViewMode);
            new HRESULT SetCurrentViewMode(uint ViewMode);
            new HRESULT GetFolder(ref Guid riid, ref IntPtr ppv);
            new HRESULT Item(int iItemIndex, ref IntPtr ppidl);
            new HRESULT ItemCount(uint uFlags, ref int pcItems);
            new HRESULT Items(uint uFlags, ref Guid riid, ref IntPtr ppv);
            new HRESULT GetSelectionMarkedItem(ref int piItem);
            new HRESULT GetFocusedItem(ref int piItem);
            new HRESULT GetItemPosition(IntPtr pidl, ref Windows.Foundation.Point ppt);
            new HRESULT GetSpacing(ref Windows.Foundation.Point ppt);
            new HRESULT GetDefaultSpacing(ref Windows.Foundation.Point ppt);
            new HRESULT GetAutoArrange();
            new HRESULT SelectItem(int iItem, int dwFlags);
            new HRESULT SelectAndPositionItems(uint cidl, IntPtr apidl, Windows.Foundation.Point apt, int dwFlags);
            HRESULT SetGroupBy(PROPERTYKEY key, bool fAscending);

            HRESULT GetGroupBy(ref PROPERTYKEY pkey, ref bool pfAscending);

            // DEPRECATED
            HRESULT SetViewProperty(IntPtr pidl, PROPERTYKEY propkey, PROPVARIANT propvar);
            // DEPRECATED
            HRESULT GetViewProperty(IntPtr pidl, PROPERTYKEY propkey, ref PROPVARIANT ppropvar);
            // DEPRECATED
            HRESULT SetTileViewProperties(IntPtr pidl, string pszPropList);
            // DEPRECATED
            HRESULT SetExtendedTileViewProperties(IntPtr pidl, string pszPropList);

            HRESULT SetText(FVTEXTTYPE iType, string pwszText);
            HRESULT SetCurrentFolderFlags(int dwMask, int dwFlags);
            HRESULT GetCurrentFolderFlags(ref int pdwFlags);
            HRESULT GetSortColumnCount(ref int pcColumns);
            HRESULT SetSortColumns(SORTCOLUMN rgSortColumns, int cColumns);
            HRESULT GetSortColumns(ref SORTCOLUMN rgSortColumns, int cColumns);
            HRESULT GetItem(int iItem, ref Guid riid, ref IntPtr ppv);
            HRESULT GetVisibleItem(int iStart, bool fPrevious, ref int piItem);
            HRESULT GetSelectedItem(int iStart, ref int piItem);
            HRESULT GetSelection(bool fNoneImpliesFolder, ref IShellItemArray ppsia);
            HRESULT GetSelectionState(IntPtr pidl, ref int pdwFlags);
            HRESULT InvokeVerbOnSelection(string pszVerb);
            HRESULT SetViewModeAndIconSize(FOLDERVIEWMODE uViewMode, int iImageSize);
            HRESULT GetViewModeAndIconSize(ref FOLDERVIEWMODE puViewMode, ref int piImageSize);
            HRESULT SetGroupSubsetCount(uint cVisibleRows);
            HRESULT GetGroupSubsetCount(ref uint pcVisibleRows);
            HRESULT SetRedraw(bool fRedrawOn);
            HRESULT IsMoveInSameFolder();
            HRESULT DoRename();
        }

        public enum FVTEXTTYPE
        {
            FVST_EMPTYTEXT = 0
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SORTCOLUMN
        {
            public PROPERTYKEY propkey;
            public SORTDIRECTION direction;
        }

        public enum SORTDIRECTION
        {
            SORT_DESCENDING = -1,
            SORT_ASCENDING = 1
        }

        public enum FOLDERVIEWMODE : int
        {
            FVM_AUTO = -1,
            FVM_FIRST = 1,
            FVM_ICON = 1,
            FVM_SMALLICON = 2,
            FVM_LIST = 3,
            FVM_DETAILS = 4,
            FVM_THUMBNAIL = 5,
            FVM_TILE = 6,
            FVM_THUMBSTRIP = 7,
            FVM_CONTENT = 8,
            FVM_LAST = 8
        }

        [DllImport("Shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT IUnknown_QueryService(IntPtr punk, ref Guid guidService, ref Guid riid, out IntPtr ppvOut);


        [ComVisible(true)]
        public enum DialogResult
        {
            //
            // Summary:
            //     Nothing is returned from the dialog box. This means that the modal dialog continues
            //     running.
            None,
            //
            // Summary:
            //     The dialog box return value is OK (usually sent from a button labeled OK).
            OK,
            //
            // Summary:
            //     The dialog box return value is Cancel (usually sent from a button labeled Cancel).
            Cancel,
            //
            // Summary:
            //     The dialog box return value is Abort (usually sent from a button labeled Abort).
            Abort,
            //
            // Summary:
            //     The dialog box return value is Retry (usually sent from a button labeled Retry).
            Retry,
            //
            // Summary:
            //     The dialog box return value is Ignore (usually sent from a button labeled Ignore).
            Ignore,
            //
            // Summary:
            //     The dialog box return value is Yes (usually sent from a button labeled Yes).
            Yes,
            //
            // Summary:
            //     The dialog box return value is No (usually sent from a button labeled No).
            No
        }
    }
}
